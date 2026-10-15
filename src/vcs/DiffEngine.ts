import * as dmp from "diff-match-patch";

const DMP = new dmp.diff_match_patch();

/**
 * Attributes Word regenerates on every save, which therefore differ between
 * two copies of the same content and would otherwise make every merge look
 * like a rewrite. Only used in the merge path - stored deltas are kept as-is.
 *
 * `w:id` is deliberately NOT in this list, although the original had it there.
 * It is not volatile: it is the identity of a bookmark, a comment anchor, a
 * tracked insertion or deletion, and a footnote reference. Stripping it
 * unpaired every `w:bookmarkStart` from its `w:bookmarkEnd` and detached every
 * comment from the text it was about -- and because the merge result was
 * stored as the new snapshot and pushed into the document, the loss was
 * permanent from the first merge onwards.
 *
 * The paragraph identifiers are matched on any prefix because in real OOXML
 * they are `w14:paraId` and `w14:textId`, in the 2010 extension namespace. The
 * original matched only `w:paraId` and `w:textId`, so neither alternative ever
 * matched anything.
 */
const VOLATILE_ATTR_RE = /\s+[A-Za-z0-9]*:?(?:rsid[A-Za-z]*|paraId|textId)="[^"]*"/g;

export function normaliseOoxml(ooxml: string): string {
  return ooxml.replace(VOLATILE_ATTR_RE, "");
}

/** A span of the base text that a patch set rewrites, as [start, end). */
export interface EditRange {
  start: number;
  end: number;
}

export class DiffEngine {
  /** Create a dmp patch string from oldText -> newText. */
  makePatch(oldText: string, newText: string): string {
    return DMP.patch_toText(DMP.patch_make(oldText, newText));
  }

  /** Apply a dmp patch string to baseText. Throws if any hunk fails. */
  applyPatch(baseText: string, patchText: string): string {
    const { result, applied } = this.applyPatchWithResults(baseText, patchText);
    const failed = applied.filter((ok) => !ok).length;
    if (failed > 0) {
      throw new Error(`patch_apply: ${failed} of ${applied.length} hunks failed.`);
    }
    return result;
  }

  /** Apply a dmp patch string to baseText, returning result and per-hunk success flags. */
  applyPatchWithResults(
    baseText: string,
    patchText: string
  ): { result: string; applied: boolean[] } {
    const patches = DMP.patch_fromText(patchText);
    const [result, applied] = DMP.patch_apply(patches, baseText);
    return { result, applied: applied as boolean[] };
  }

  /**
   * The spans of `oldText` that this patch set rewrites.
   *
   * Each dmp hunk carries the offset and length of the region it matches in the
   * old text, including the few characters of context dmp keeps on either side.
   * Including that context is wanted here: two edits close enough to share
   * context are close enough to be worth refusing to merge silently.
   */
  editRanges(oldText: string, newText: string): EditRange[] {
    // @types/diff-match-patch declares patch_make as returning the patch
    // *constructor* rather than the patch, so the fields have to be reached
    // through a cast. The values are the ones the library documents.
    const patches = DMP.patch_make(oldText, newText) as unknown as dmp.patch_obj[];
    return patches.map((p) => ({
      start: p.start1 ?? 0,
      end: (p.start1 ?? 0) + (p.length1 ?? 0),
    }));
  }

  /** Sanity check: round-trip a patch and verify it reproduces newText. */
  verify(oldText: string, newText: string): boolean {
    try {
      return this.applyPatch(oldText, this.makePatch(oldText, newText)) === newText;
    } catch {
      return false;
    }
  }
}

/** True if any range in `a` touches or overlaps any range in `b`. */
export function rangesOverlap(a: EditRange[], b: EditRange[]): boolean {
  for (const x of a) {
    for (const y of b) {
      if (x.start < y.end && y.start < x.end) return true;
    }
  }
  return false;
}
