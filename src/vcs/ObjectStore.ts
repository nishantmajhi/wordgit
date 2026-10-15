/* global Word */

import {
  emptyRepo,
  NS,
  parseRepoXml,
  RepoFormatError,
  repoToXml,
  RepoXml,
  XmlParse,
} from "./RepoFormat";

export { NS, RepoFormatError } from "./RepoFormat";
export type { RepoXml, CommitXml, Encoding } from "./RepoFormat";

/**
 * The only thing `ObjectStore` needs from the host: the Custom XML Parts in
 * WordGit's namespace, and a way to write them.
 *
 * Office.js is behind this interface rather than in front of it so that the
 * rules about *when* it is safe to write can be tested. Those rules are the
 * part with teeth -- see `save()`.
 */
export interface CustomXmlPartHost {
  /** Every part in the WordGit namespace, in document order. */
  list(): Promise<Array<{ id: string; xml: string }>>;
  /** Replace a part's content in one operation. Returns false if the host cannot. */
  setXml(id: string, xml: string): Promise<boolean>;
  /** Add a new part. Returns its id. */
  add(xml: string): Promise<string>;
  remove(id: string): Promise<void>;
}

/** Office.js implementation of `CustomXmlPartHost`. */
export class WordCustomXmlParts implements CustomXmlPartHost {
  private context: Word.RequestContext;

  constructor(context: Word.RequestContext) {
    this.context = context;
  }

  async list(): Promise<Array<{ id: string; xml: string }>> {
    const items = this.context.document.customXmlParts.getByNamespace(NS);
    items.load("items");
    await this.context.sync();

    const pending = items.items.map((part) => ({ part, xml: part.getXml() }));
    await this.context.sync();
    return pending.map(({ part, xml }) => ({ id: part.id, xml: xml.value }));
  }

  /**
   * `CustomXmlPart.setXml` is WordApi 1.4. The original's comment said Office.js
   * had no way to update a part and deleted it before adding a replacement;
   * there is one, and it removes the window in which the document has no
   * history at all. The runtime check is for hosts old enough not to have it.
   */
  async setXml(id: string, xml: string): Promise<boolean> {
    const part = this.context.document.customXmlParts.getItem(id) as Word.CustomXmlPart & {
      setXml?: (xml: string) => void;
    };
    if (typeof part.setXml !== "function") return false;
    part.setXml(xml);
    await this.context.sync();
    return true;
  }

  async add(xml: string): Promise<string> {
    const part = this.context.document.customXmlParts.add(xml);
    part.load("id");
    await this.context.sync();
    return part.id;
  }

  async remove(id: string): Promise<void> {
    this.context.document.customXmlParts.getItem(id).delete();
    await this.context.sync();
  }
}

/**
 * Reads and writes the repository stored in the document's Custom XML Parts.
 *
 * One instance is one load-modify-save cycle. `load()` remembers which part it
 * read and whether it could be understood, and `save()` refuses to write when
 * it could not -- the caller has to decide what to do about a damaged history,
 * and silently replacing it is not one of the options.
 */
export class ObjectStore {
  private host: CustomXmlPartHost;
  private parse: XmlParse | undefined;
  private partId: string | null = null;
  private readable = true;

  constructor(host: CustomXmlPartHost | Word.RequestContext, parse?: XmlParse) {
    this.host = isHost(host) ? host : new WordCustomXmlParts(host);
    this.parse = parse;
  }

  /** True once `load()` has found a part it could not interpret. */
  get damaged(): boolean {
    return !this.readable;
  }

  /**
   * Load the repository, or return an empty one if this document has never had
   * a repository in it.
   *
   * Throws `RepoFormatError` if the document *does* carry WordGit data that
   * cannot be read. That distinction is the point: the original returned an
   * empty repository in both cases, and the task pane then offered to
   * initialise a document whose history was merely unreadable, overwriting it.
   */
  async load(): Promise<RepoXml> {
    const parts = await this.host.list();
    if (parts.length === 0) {
      this.partId = null;
      this.readable = true;
      return emptyRepo();
    }

    // With more than one part the extra ones are debris from a host that could
    // not replace a part in place, where a new part was added and the old one
    // outlived the attempt to remove it. The newest write is always the larger
    // history, so the part with the most commits is the one to keep.
    let best: { id: string; repo: RepoXml } | null = null;
    let firstError: RepoFormatError | null = null;
    for (const part of parts) {
      try {
        const repo = this.parseXml(part.xml);
        if (!best || repo.objects.length > best.repo.objects.length) {
          best = { id: part.id, repo };
        }
      } catch (e) {
        if (!firstError && e instanceof RepoFormatError) firstError = e;
      }
    }

    if (!best) {
      this.partId = parts[0].id;
      this.readable = false;
      throw firstError ?? new RepoFormatError("WordGit data in this document could not be read.");
    }

    this.partId = best.id;
    this.readable = true;
    return best.repo;
  }

  /**
   * Write the repository back.
   *
   * Prefers replacing the existing part in place. Where the host has no such
   * operation, the new part is added *before* the old one is removed: the
   * original deleted first, so an error, a closed document or a crash in the
   * window between the two calls left the document with no history at all.
   */
  async save(repo: RepoXml): Promise<void> {
    if (!this.readable) {
      throw new RepoFormatError(
        "Refusing to write: this document already contains WordGit data that could not be " +
          "read, and writing would replace it. Keep a copy of the file before going further."
      );
    }

    const xml = repoToXml(repo);

    if (this.partId !== null) {
      if (await this.host.setXml(this.partId, xml)) return;
      const replacement = await this.host.add(xml);
      await this.host.remove(this.partId);
      this.partId = replacement;
      return;
    }

    this.partId = await this.host.add(xml);
  }

  private parseXml(xml: string): RepoXml {
    return this.parse ? parseRepoXml(xml, this.parse) : parseRepoXml(xml);
  }
}

function isHost(v: CustomXmlPartHost | Word.RequestContext): v is CustomXmlPartHost {
  return typeof (v as CustomXmlPartHost).list === "function";
}
