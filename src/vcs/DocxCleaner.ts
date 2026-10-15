/**
 * DocxCleaner.ts - strip WordGit's history out of a .docx.
 *
 * A repository lives in Custom XML Parts inside the document, so it travels
 * with the file. That is the whole point of the design, and it is also why
 * there has to be a way to hand someone a copy without it. Everything here
 * works on the zip container directly; nothing in it touches Office.js.
 */

import JSZip from "jszip";
import { NS } from "./RepoFormat";

const ITEM_PATH = /^customXml\/item(\d+)\.xml$/;

/** Read the name and attributes of the first element in an XML string. */
function firstElement(xml: string): { name: string; attrs: string } | null {
  // Skip the declaration, any DOCTYPE, comments and processing instructions.
  const re = /<(?!\?|!)([^\s/>]+)([^>]*)>/;
  const m = re.exec(xml);
  return m ? { name: m[1], attrs: m[2] } : null;
}

/**
 * True if this Custom XML Part is a WordGit repository.
 *
 * Checking the root element rather than searching the whole text for the
 * namespace matters: a user's own Custom XML Part that merely quotes the
 * namespace string -- a document *about* WordGit, for instance -- would
 * otherwise be deleted from their file along with ours.
 */
export function isWordGitPart(xml: string): boolean {
  const root = firstElement(xml);
  if (!root) return false;
  const localName = root.name.includes(":") ? root.name.split(":").pop() : root.name;
  return localName === "repo" && root.attrs.includes(NS);
}

/** Read one attribute out of a raw start-tag attribute string. */
function attr(attrs: string, name: string): string | null {
  const re = new RegExp(`\\b${name}\\s*=\\s*("([^"]*)"|'([^']*)')`);
  const m = re.exec(attrs);
  if (!m) return null;
  return m[2] !== undefined ? m[2] : m[3];
}

/**
 * Remove every `<tag>` element whose `attrName` value satisfies `matches`.
 *
 * Done by scanning whole elements and reading the attribute, rather than by
 * building a regex around the value. The original interpolated the part name
 * into a pattern of the form `<Override[^>]…PartName="/customXml/item1.xml"…`,
 * which matches an Override for `item1.xml` but also matches one whose
 * PartName merely contains that text, and which silently fails to match an
 * element written with a separate closing tag rather than a self-closing one.
 */
function removeElements(
  xml: string,
  tag: string,
  attrName: string,
  matches: (value: string) => boolean
): string {
  const re = new RegExp(`<${tag}\\b[^>]*?/?>`, "g");
  return xml.replace(re, (whole) => {
    const attrs = whole.slice(tag.length + 1, whole.length - 1);
    const value = attr(attrs, attrName);
    return value !== null && matches(value) ? "" : whole;
  });
}

/** The `customXml/...` path a relationship or content-type entry refers to. */
function normalisePartPath(value: string): string {
  return value.replace(/^\.\.\//, "").replace(/^\//, "");
}

export interface CleanResult {
  /** The paths removed from the package. */
  removed: string[];
}

/**
 * Remove every WordGit Custom XML Part from an open .docx, along with the
 * package bookkeeping that refers to it.
 *
 * Returns the list of removed paths; an empty list means the document carried
 * no WordGit data and nothing was changed.
 */
export async function stripWordGitParts(zip: JSZip): Promise<CleanResult> {
  const itemPaths = Object.keys(zip.files).filter((p) => ITEM_PATH.test(p));

  const removed: string[] = [];
  for (const path of itemPaths) {
    if (!isWordGitPart(await zip.file(path)!.async("text"))) continue;
    const item = path.replace(/^customXml\//, "").replace(/\.xml$/, "");

    // The rels file names the itemProps part that holds the schema references
    // for this item. Leaving that behind leaves an orphan part in the package
    // and an entry for it in [Content_Types].xml, which is exactly the kind of
    // debris a "clean copy" is supposed not to have.
    const relsPath = `customXml/_rels/${item}.xml.rels`;
    const relsFile = zip.file(relsPath);
    const props: string[] = [];
    if (relsFile) {
      const relsXml = await relsFile.async("text");
      const re = /<Relationship\b[^>]*?\/?>/g;
      let m: RegExpExecArray | null;
      while ((m = re.exec(relsXml)) !== null) {
        const target = attr(m[0], "Target");
        if (target) props.push(`customXml/${normalisePartPath(target)}`);
      }
    }

    for (const p of [path, relsPath, ...props]) {
      if (!zip.file(p)) continue;
      zip.remove(p);
      removed.push(p);
    }
  }

  if (removed.length === 0) return { removed: [] };

  const gone = new Set(removed.map(normalisePartPath));

  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    const ct = await ctFile.async("text");
    zip.file(
      "[Content_Types].xml",
      removeElements(ct, "Override", "PartName", (v) => gone.has(normalisePartPath(v)))
    );
  }

  const relsPath = "word/_rels/document.xml.rels";
  const docRels = zip.file(relsPath);
  if (docRels) {
    const rels = await docRels.async("text");
    zip.file(
      relsPath,
      removeElements(rels, "Relationship", "Target", (v) => gone.has(normalisePartPath(v)))
    );
  }

  return { removed };
}

/**
 * Derive "<basename>_clean.docx" from a document URL or path.
 *
 * A query string or fragment is dropped first: a document opened from a
 * SharePoint or OneDrive location has one, and it would otherwise end up in
 * the suggested filename.
 */
export function cleanFileName(url: string | null | undefined): string {
  const raw = (url || "document.docx").split("#")[0].split("?")[0];
  let last = raw.split(/[\\/]/).pop() || "document.docx";
  try {
    last = decodeURIComponent(last);
  } catch {
    // A name that is not valid percent-encoding is used as it stands.
  }
  if (!last) last = "document.docx";
  const dot = last.lastIndexOf(".");
  const base = dot > 0 ? last.slice(0, dot) : last;
  return `${base}_clean.docx`;
}
