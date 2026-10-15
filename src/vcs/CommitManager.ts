/* global Word */

import { CommitXml, RepoFormatError, ROOT_PARENT } from "./RepoFormat";
import { ObjectStore } from "./ObjectStore";
import { DiffEngine } from "./DiffEngine";
import { MergeConflictError, MergeManager } from "./MergeManager";
import {
  branchNameFromHead,
  CommitInfo,
  currentRef,
  encodeDelta,
  findCommit,
  findRef,
  logFor,
  mergeBase,
  reconstruct,
} from "./History";

export type { CommitInfo } from "./History";

/**
 * Identifier for a new commit.
 *
 * Not a content hash, despite the name the original gave it: WordGit never
 * needs to recognise identical content, only to tell two commits apart, and a
 * real hash of a whole OOXML snapshot would cost more than it is worth here.
 * The leading millisecond count keeps identifiers roughly sortable; the random
 * tail is what actually separates two commits made in the same millisecond.
 */
function newCommitId(taken: ReadonlySet<string>): string {
  for (;;) {
    const stamp = Date.now().toString(36).padStart(8, "0").slice(-8);
    let tail = "";
    // Math.random().toString(36) is not a fixed length -- 0.5 renders as "0.i",
    // which is one character after the prefix. The original sliced four
    // characters out of it regardless and padded the shortfall with zeroes,
    // quietly narrowing the random part on exactly the values where it mattered.
    while (tail.length < 4) tail += Math.random().toString(36).slice(2);
    const id = stamp + tail.slice(0, 4);
    if (!taken.has(id)) return id;
  }
}

function idsIn(repo: { objects: Array<{ hash: string }> }): Set<string> {
  return new Set(repo.objects.map((c) => c.hash));
}

/**
 * The document, as the VCS needs it: something to read OOXML out of and
 * something to push OOXML back into. Office.js lives behind this so that a
 * whole commit-branch-merge sequence can be run without Word.
 */
export interface DocumentIO {
  getOoxml(): Promise<string>;
  /** Replace the body. `tracked` inserts as tracked changes for the user to review. */
  replaceBody(ooxml: string, tracked: boolean): Promise<void>;
  /** Best guess at who is editing, or null if the host will not say. */
  authorName?(): Promise<string | null>;
}

/** Office.js implementation of `DocumentIO`. */
export class WordDocumentIO implements DocumentIO {
  async getOoxml(): Promise<string> {
    return Word.run(async (context) => {
      const ooxml = context.document.body.getOoxml();
      await context.sync();
      return ooxml.value;
    });
  }

  async replaceBody(ooxml: string, tracked: boolean): Promise<void> {
    await Word.run(async (context) => {
      context.document.changeTrackingMode = tracked
        ? Word.ChangeTrackingMode.trackMineOnly
        : Word.ChangeTrackingMode.off;
      context.document.body.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();
      if (tracked) {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        await context.sync();
      }
    });
  }

  async authorName(): Promise<string | null> {
    try {
      return await Word.run(async (context) => {
        const props = context.document.properties;
        props.load("lastAuthor,author");
        await context.sync();
        return props.lastAuthor || props.author || null;
      });
    } catch {
      return null;
    }
  }
}

export class CommitManager {
  private store: ObjectStore;
  private diff: DiffEngine;
  private io: DocumentIO;

  constructor(store: ObjectStore, diff: DiffEngine, io: DocumentIO) {
    this.store = store;
    this.diff = diff;
    this.io = io;
  }

  /**
   * Initialise a new repository: a main branch and a root commit holding a
   * full snapshot. Does nothing if one already exists.
   *
   * `store.load()` throws when the document holds WordGit data that cannot be
   * read, and that throw is allowed through on purpose. The original could not
   * tell that case from an empty document and would initialise over the top of
   * it, which replaced the Custom XML Part and destroyed the history.
   */
  async init(): Promise<void> {
    const repo = await this.store.load();
    if (repo.refs.length > 0) return;

    const author = (await this.resolveAuthor()) ?? repo.author;
    const { delta, encoding } = encodeDelta(await this.io.getOoxml(), "snapshot");
    const hash = newCommitId(idsIn(repo));

    repo.author = author;
    repo.refs.push({ name: "main", tip: hash });
    repo.head = "refs/heads/main";
    repo.objects.push({
      hash,
      parent: ROOT_PARENT,
      parent2: "",
      author,
      timestamp: new Date().toISOString(),
      message: "Initial commit",
      delta,
      encoding,
    });
    await this.store.save(repo);
  }

  /**
   * Commit the current document state to the current branch, as a
   * diff-match-patch delta against the previous commit. Returns the new hash.
   */
  async commit(message: string): Promise<string> {
    const repo = await this.store.load();
    const ref = currentRef(repo);
    if (!ref) throw new Error("No branch found. Did you Init?");

    const currentOoxml = await this.io.getOoxml();
    const parentHash = ref.tip;
    const parentCommit = findCommit(repo, parentHash);

    const encoded = parentCommit
      ? encodeDelta(
          this.diff.makePatch(reconstruct(repo, parentHash, this.diff), currentOoxml),
          "dmp"
        )
      : encodeDelta(currentOoxml, "snapshot");

    // Asked per commit, not once at init: a document that passes between two
    // people should say who wrote which commit, which is the reason the format
    // carries an author on the commit as well as one in the config.
    const commitAuthor = (await this.resolveAuthor()) ?? repo.author;
    const hash = newCommitId(idsIn(repo));
    repo.objects.push({
      hash,
      parent: parentCommit ? parentHash : ROOT_PARENT,
      parent2: "",
      author: commitAuthor,
      timestamp: new Date().toISOString(),
      message,
      delta: encoded.delta,
      encoding: encoded.encoding,
    });
    ref.tip = hash;
    await this.store.save(repo);
    return hash;
  }

  /** Rebuild the OOXML at a given commit. */
  async reconstruct(hash: string): Promise<string> {
    return reconstruct(await this.store.load(), hash, this.diff);
  }

  /**
   * Restore the document to a past commit with change tracking on, so the user
   * reviews the difference and can reject it.
   */
  async checkout(hash: string): Promise<void> {
    await this.io.replaceBody(await this.reconstruct(hash), true);
  }

  /** Swap document content silently. Used when switching branches. */
  async checkoutSilent(hash: string): Promise<void> {
    await this.io.replaceBody(await this.reconstruct(hash), false);
  }

  /** The commit log for the current branch, newest first. */
  async log(): Promise<CommitInfo[]> {
    return logFor(await this.store.load());
  }

  /**
   * Merge another branch into the current one.
   *
   * On conflict this throws and the document is left untouched; nothing is
   * written to the repository either, so the branches stay where they were.
   */
  async merge(theirBranch: string): Promise<string> {
    const repo = await this.store.load();
    const ourBranchName = branchNameFromHead(repo.head);
    const ourRef = findRef(repo, ourBranchName);
    const theirRef = findRef(repo, theirBranch);

    if (!ourRef) throw new Error(`Current branch "${ourBranchName}" not found.`);
    if (!theirRef) throw new Error(`Branch "${theirBranch}" not found.`);

    const ourTip = ourRef.tip;
    const theirTip = theirRef.tip;
    if (ourTip === theirTip) throw new Error("Branches are already at the same commit.");

    const baseHash = mergeBase(repo, ourTip, theirTip);
    if (!baseHash) throw new Error("No common ancestor found - cannot merge.");

    // Already merged: their tip is behind ours and has nothing new to offer.
    if (baseHash === theirTip) {
      throw new Error(`Branch "${theirBranch}" is already merged into "${ourBranchName}".`);
    }

    if (baseHash === ourTip) {
      // Fast-forward: our branch has nothing of its own, so move the pointer.
      ourRef.tip = theirTip;
      await this.store.save(repo);
      await this.io.replaceBody(reconstruct(repo, theirTip, this.diff), false);
      return theirTip;
    }

    const merged = new MergeManager(this.diff).threeWayMerge(
      reconstruct(repo, baseHash, this.diff),
      reconstruct(repo, ourTip, this.diff),
      reconstruct(repo, theirTip, this.diff)
    );

    const commitAuthor = (await this.resolveAuthor()) ?? repo.author;
    const hash = newCommitId(idsIn(repo));
    const encoded = encodeDelta(merged, "snapshot");
    const mergeCommit: CommitXml = {
      hash,
      parent: ourTip,
      // Recording the second parent is what makes the merged branch part of
      // this branch's history. Without it the log showed none of the work that
      // had just been merged in, the graph drew no merge edge, and merging the
      // same branch again replayed the whole thing because the search for a
      // common ancestor could never see past our own first-parent chain.
      parent2: theirTip,
      author: commitAuthor,
      timestamp: new Date().toISOString(),
      message: `Merge branch '${theirBranch}' into ${ourBranchName}`,
      delta: encoded.delta,
      encoding: encoded.encoding,
    };

    repo.objects.push(mergeCommit);
    ourRef.tip = hash;
    await this.store.save(repo);
    await this.io.replaceBody(merged, true);
    return hash;
  }

  private async resolveAuthor(): Promise<string | null> {
    if (!this.io.authorName) return null;
    const name = await this.io.authorName();
    return name && name.trim() ? name.trim() : null;
  }
}

export { MergeConflictError, RepoFormatError };
