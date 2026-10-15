/**
 * RepoFormat.ts - the storage contract for a WordGit repository.
 *
 * This module is deliberately free of Office.js and of any DOM global at load
 * time, so the format can be exercised without Word and without a browser.
 * `ObjectStore` is the half that talks to the document; this is the half that
 * decides what a repository *is*.
 *
 * The data lives inside users' .docx files and cannot be migrated centrally,
 * so both directions matter: anything this writes must still be readable by a
 * later version, and anything an earlier version wrote must still be readable
 * here. See README.md, "Storage format", for the compatibility rules.
 */

export const NS = "http://schemas.wordgit.local/vcs/1";

/** Sentinel parent of the first commit. Not a real object. */
export const ROOT_PARENT = "00root";

export const ENCODINGS = ["dmp", "snapshot", "dmp-lz", "snapshot-lz"] as const;
export type Encoding = (typeof ENCODINGS)[number];

export interface RepoXml {
  head: string;
  author: string;
  refs: Array<{ name: string; tip: string }>;
  objects: Array<CommitXml>;
}

export interface CommitXml {
  hash: string;
  /** Hash of the first parent, or ROOT_PARENT for the initial commit. */
  parent: string;
  /** Hash of the second parent for a merge commit, otherwise "". */
  parent2: string;
  author: string;
  timestamp: string;
  message: string;
  /**
   * Payload string. Meaning depends on `encoding`:
   *   - "snapshot"    : raw OOXML (uncompressed)
   *   - "dmp"         : raw diff-match-patch patch text (uncompressed)
   *   - "snapshot-lz" : OOXML compressed with LZString.compressToBase64()
   *   - "dmp-lz"      : dmp patch text compressed with LZString.compressToBase64()
   */
  delta: string;
  encoding: Encoding;
}

export function emptyRepo(): RepoXml {
  return { head: "refs/heads/main", author: "Author", refs: [], objects: [] };
}

/**
 * Thrown when a document contains WordGit data that cannot be interpreted.
 *
 * This is emphatically not the same as "this document has no repository", and
 * the difference is the whole reason the class exists. The original code
 * returned an empty repository for both cases, so a single damaged byte in the
 * stored XML presented in the task pane as an uninitialised document -- and
 * the next Init, or the next Alt+C, replaced the part and destroyed the
 * history for good. Callers must treat this as "stop, do not write".
 */
export class RepoFormatError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "RepoFormatError";
  }
}

/**
 * Characters XML 1.0 cannot represent at all, not even as a numeric character
 * reference: C0 controls other than tab, newline and carriage return, plus the
 * two non-characters. A payload containing one cannot be stored, and finding
 * that out at read time -- when the document is the only copy -- is far worse
 * than refusing at write time.
 */
const ILLEGAL_XML_CHAR = new RegExp("[\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F\\uFFFE\\uFFFF]");

/** Escape for an XML attribute value. */
function escAttr(s: string): string {
  return escText(s).replace(/"/g, "&quot;").replace(/\t/g, "&#9;").replace(/\n/g, "&#10;");
}

/**
 * Escape for XML element content.
 *
 * Carriage return has to become a numeric character reference. A parser
 * performs line-ending normalisation before anything else sees the text, so a
 * literal CR in the source is silently rewritten to LF -- which means a payload
 * containing CRLF does not survive a save/load cycle, and the next delta is
 * then computed against a base that no longer matches what was committed.
 * `&#13;` survives, because normalisation happens before reference expansion.
 * That is also why the payload is no longer written inside a CDATA section:
 * CDATA cannot carry a character reference, so CR cannot be preserved in one.
 */
function escText(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\r/g, "&#13;");
}

export function repoToXml(repo: RepoXml): string {
  const illegal = findIllegalChar(repo);
  if (illegal) throw new RepoFormatError(illegal);

  const refs = repo.refs
    .map((r) => `<branch name="${escAttr(r.name)}" tip="${escAttr(r.tip)}"/>`)
    .join("\n    ");
  const objects = repo.objects
    .map(
      (c) => `<commit hash="${escAttr(c.hash)}" parent="${escAttr(c.parent)}"${
        c.parent2 ? ` parent2="${escAttr(c.parent2)}"` : ""
      }>
      <author>${escText(c.author)}</author>
      <timestamp>${escText(c.timestamp)}</timestamp>
      <message>${escText(c.message)}</message>
      <delta encoding="${escAttr(c.encoding)}">${escText(c.delta)}</delta>
    </commit>`
    )
    .join("\n    ");
  return `<?xml version="1.0" encoding="UTF-8"?>
<wordgit:repo xmlns:wordgit="${NS}">
  <HEAD>${escText(repo.head)}</HEAD>
  <config><author>${escText(repo.author)}</author></config>
  <refs>
    ${refs}
  </refs>
  <objects>
    ${objects}
  </objects>
</wordgit:repo>`;
}

function describeIllegal(where: string, value: string): string | null {
  const at = value.search(ILLEGAL_XML_CHAR);
  if (at < 0) return null;
  const code = value.charCodeAt(at).toString(16).toUpperCase().padStart(4, "0");
  return `${where} contains U+${code}, which XML cannot store.`;
}

function findIllegalChar(repo: RepoXml): string | null {
  let found = describeIllegal("HEAD", repo.head) ?? describeIllegal("the author name", repo.author);
  for (const r of repo.refs) {
    found = found ?? describeIllegal(`branch name "${r.name}"`, r.name);
  }
  for (const c of repo.objects) {
    found =
      found ??
      describeIllegal(`the message of commit ${c.hash}`, c.message) ??
      describeIllegal(`the author of commit ${c.hash}`, c.author) ??
      describeIllegal(`the payload of commit ${c.hash}`, c.delta);
  }
  return found;
}

/**
 * How `parseRepoXml` turns a string into a document. Injectable so the format
 * can be tested without a browser.
 */
export type XmlParse = (xml: string) => Document;

const domParse: XmlParse = (xml) => new DOMParser().parseFromString(xml, "application/xml");

/**
 * Parse stored repository XML, or throw `RepoFormatError`.
 *
 * Every rejection here is something that cannot be interpreted, not something
 * that merely looks unusual: an unparseable document, a root element that is
 * not a WordGit repository, a commit with no hash or an unrecognised payload
 * encoding, a branch with no name. Being lenient about the rest is deliberate,
 * because the alternative is refusing to open a document that an older version
 * wrote and that the user can still read.
 */
export function parseRepoXml(xml: string, parse: XmlParse = domParse): RepoXml {
  let doc: Document;
  try {
    doc = parse(xml);
  } catch (e) {
    throw new RepoFormatError(`WordGit data is not valid XML: ${(e as Error).message}`);
  }

  // A browser's DOMParser does not throw on malformed XML; it hands back a
  // document whose root is a <parsererror> (Firefox) or an <html> wrapping one
  // (Chromium). Checking the root element catches every shape of that, as well
  // as XML that parses perfectly but is not ours.
  const root = doc.documentElement;
  if (!root || root.localName !== "repo") {
    throw new RepoFormatError(
      `WordGit data is not a repository (root element is <${root ? root.nodeName : "nothing"}>).`
    );
  }

  const head = text(doc.querySelector("HEAD")) ?? "refs/heads/main";
  const author = text(doc.querySelector("config > author")) ?? "Author";

  const refs: RepoXml["refs"] = [];
  const seenBranch = new Set<string>();
  forEach(doc.querySelectorAll("refs > branch"), (el) => {
    const name = el.getAttribute("name");
    if (!name) throw new RepoFormatError("A <branch> entry has no name.");
    if (seenBranch.has(name)) throw new RepoFormatError(`Branch "${name}" is listed twice.`);
    seenBranch.add(name);
    refs.push({ name, tip: el.getAttribute("tip") ?? "" });
  });

  const objects: CommitXml[] = [];
  const seenHash = new Set<string>();
  forEach(doc.querySelectorAll("objects > commit"), (el) => {
    const hash = el.getAttribute("hash");
    if (!hash) throw new RepoFormatError("A <commit> entry has no hash.");
    if (seenHash.has(hash)) throw new RepoFormatError(`Commit ${hash} is stored twice.`);
    seenHash.add(hash);

    const parent = el.getAttribute("parent");
    if (!parent) throw new RepoFormatError(`Commit ${hash} has no parent pointer.`);

    const deltaEl = el.querySelector("delta");
    if (!deltaEl) throw new RepoFormatError(`Commit ${hash} has no payload.`);
    const encoding = deltaEl.getAttribute("encoding") ?? "";
    if (!(ENCODINGS as readonly string[]).includes(encoding)) {
      throw new RepoFormatError(
        `Commit ${hash} uses payload encoding "${encoding}", which this version does not ` +
          `understand. The document may have been written by a newer WordGit.`
      );
    }

    objects.push({
      hash,
      parent,
      parent2: el.getAttribute("parent2") ?? "",
      author: text(el.querySelector("author")) ?? "",
      timestamp: text(el.querySelector("timestamp")) ?? "",
      message: text(el.querySelector("message")) ?? "",
      delta: deltaEl.textContent ?? "",
      encoding: encoding as Encoding,
    });
  });

  return { head, author, refs, objects };
}

function text(el: Element | null): string | null {
  return el ? (el.textContent ?? "") : null;
}

/** NodeList.forEach is not present on every host object this runs against. */
function forEach(list: ArrayLike<Element>, fn: (el: Element) => void): void {
  for (let i = 0; i < list.length; i++) fn(list[i]);
}
