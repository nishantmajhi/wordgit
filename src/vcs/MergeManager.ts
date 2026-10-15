import { DiffEngine, normaliseOoxml, rangesOverlap } from "./DiffEngine";

/** Thrown when two branches cannot be combined without a person deciding. */
export class MergeConflictError extends Error {
  constructor(message: string) {
    super(`CONFLICT: ${message}`);
    this.name = "MergeConflictError";
  }
}

export class MergeManager {
  private diff: DiffEngine;

  constructor(diff: DiffEngine) {
    this.diff = diff;
  }

  /**
   * Three-way merge of OOXML strings.
   *
   * All three inputs are normalised first, so that revision-save identifiers
   * Word regenerates on every save do not read as edits. The result is
   * therefore normalised OOXML; Word puts its own identifiers back when the
   * content is inserted.
   *
   * Throws `MergeConflictError` if the two sides cannot be combined. The caller
   * must catch it and leave the document alone.
   *
   * The conflict test is a check that the two sides edited disjoint regions of
   * the base, not a check that diff-match-patch managed to apply both patch
   * sets. Those are very different questions. `patch_apply` is deliberately
   * fuzzy: it searches for each hunk's context near where it expects it, and
   * when both sides have rewritten the same sentence it will often find
   * somewhere plausible to put the second one and report success. The original
   * trusted that report, so two people editing the same paragraph got a silent,
   * mangled merge rather than a conflict -- the one outcome the doc comment
   * said it was built to avoid.
   *
   * Applying both patch sets in each order and requiring the same answer is the
   * second half of the guard. Disjoint edits commute; if they do not, a hunk
   * landed somewhere other than where it belonged, whatever it reported.
   */
  threeWayMerge(baseRaw: string, oursRaw: string, theirsRaw: string): string {
    const base = normaliseOoxml(baseRaw);
    const ours = normaliseOoxml(oursRaw);
    const theirs = normaliseOoxml(theirsRaw);

    if (ours === base) return theirs;
    if (theirs === base) return ours;
    if (ours === theirs) return ours;

    const oursRanges = this.diff.editRanges(base, ours);
    const theirsRanges = this.diff.editRanges(base, theirs);
    if (rangesOverlap(oursRanges, theirsRanges)) {
      throw new MergeConflictError(
        "both branches changed the same part of the document. Merge them by hand, or " +
          "restore one side and re-apply the other."
      );
    }

    const oursPatch = this.diff.makePatch(base, ours);
    const theirsPatch = this.diff.makePatch(base, theirs);

    const forward = this.applyBoth(base, oursPatch, "ours", theirsPatch, "theirs");
    const backward = this.applyBoth(base, theirsPatch, "theirs", oursPatch, "ours");

    if (forward !== backward) {
      throw new MergeConflictError(
        "the two sets of changes do not combine to the same document in either order, " +
          "so the result cannot be trusted. Merge them by hand."
      );
    }
    return forward;
  }

  private applyBoth(
    base: string,
    firstPatch: string,
    firstName: string,
    secondPatch: string,
    secondName: string
  ): string {
    const first = this.diff.applyPatchWithResults(base, firstPatch);
    this.requireAllApplied(first.applied, firstName);
    const second = this.diff.applyPatchWithResults(first.result, secondPatch);
    this.requireAllApplied(second.applied, secondName);
    return second.result;
  }

  private requireAllApplied(applied: boolean[], side: string): void {
    const failed = applied.filter((ok) => !ok).length;
    if (failed > 0) {
      throw new MergeConflictError(
        `${failed} of ${applied.length} change(s) from "${side}" could not be placed. ` +
          "Manual resolution required."
      );
    }
  }
}
