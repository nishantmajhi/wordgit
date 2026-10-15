/**
 * History.ts - the decisions a commit graph makes, with nothing in them that
 * talks to Word.
 *
 * Every function here takes a whole `RepoXml` and returns an answer. That
 * shape is deliberate: the original code put these walks inside
 * `CommitManager`, where each one re-read the Custom XML Part through
 * Office.js first -- so a single merge loaded and re-parsed the entire
 * repository six times, and none of it could be tested without Word.
 */

import LZString from "lz-string";
import { CommitXml, Encoding, RepoXml, ROOT_PARENT } from "./RepoFormat";
import { DiffEngine } from "./DiffEngine";

/** "refs/heads/main" -> "main". Anything else is returned unchanged. */
export function branchNameFromHead(head: string): string {
  return head.startsWith("refs/heads/") ? head.slice("refs/heads/".length) : head;
}

export function isSnapshot(encoding: Encoding): boolean {
  return encoding === "snapshot" || encoding === "snapshot-lz";
}

/** Decompress a delta payload according to its encoding. Returns raw OOXML/patch text. */
export function decodeDelta(delta: string, encoding: Encoding): string {
  if (encoding === "snapshot-lz" || encoding === "dmp-lz") {
    // An empty payload compresses to an empty string, which is legitimate:
    // an empty document, or a commit that changed nothing. Only a non-empty
    // payload that fails to decompress means corruption.
    if (delta === "") return "";
    const decompressed = LZString.decompressFromBase64(delta);
    if (decompressed === null || decompressed === "") {
      throw new Error("Failed to decompress commit data - the stored history may be damaged.");
    }
    return decompressed;
  }
  return delta; // "snapshot" | "dmp" - already plain text
}

/**
 * Compress a payload for storage.
 *
 * Compression is skipped when it does not pay for itself, which it does not on
 * short payloads: LZString's base64 form is about a third larger than its raw
 * output, so a small patch can come back longer than it went in. Storing the
 * larger of the two forms would be a pure loss, and both forms are declared by
 * the encoding, so either is readable.
 */
export function encodeDelta(
  raw: string,
  kind: "snapshot" | "dmp"
): { delta: string; encoding: Encoding } {
  const compressed = LZString.compressToBase64(raw);
  if (!compressed || compressed.length >= raw.length) {
    return { delta: raw, encoding: kind };
  }
  return { delta: compressed, encoding: kind === "snapshot" ? "snapshot-lz" : "dmp-lz" };
}

export function findCommit(repo: RepoXml, hash: string): CommitXml | undefined {
  return repo.objects.find((c) => c.hash === hash);
}

export function findRef(repo: RepoXml, name: string): { name: string; tip: string } | undefined {
  return repo.refs.find((r) => r.name === name);
}

export function currentRef(repo: RepoXml): { name: string; tip: string } | undefined {
  return findRef(repo, branchNameFromHead(repo.head));
}

/**
 * The chain of commits that has to be replayed to rebuild `hash`, oldest
 * first, starting at the most recent snapshot at or below it.
 *
 * Stopping at a snapshot rather than walking all the way to the root is what
 * makes a merge commit worth storing as a snapshot: everything before it is
 * already folded in, so replaying it again is wasted work. The original walked
 * to the root every time and then insisted the root be a snapshot, which meant
 * reconstructing the tip of a long branch replayed every delta ever made.
 */
export function replayChain(repo: RepoXml, hash: string): CommitXml[] {
  const chain: CommitXml[] = [];
  const seen = new Set<string>();
  let current = findCommit(repo, hash);
  if (!current) throw new Error(`Commit ${hash} not found in this document's history.`);

  while (current) {
    if (seen.has(current.hash)) {
      throw new Error(`Commit ${current.hash} is its own ancestor - the history is damaged.`);
    }
    seen.add(current.hash);
    chain.unshift(current);
    if (isSnapshot(current.encoding)) return chain;
    if (current.parent === ROOT_PARENT) break;
    const parent: CommitXml | undefined = findCommit(repo, current.parent);
    if (!parent) {
      throw new Error(
        `Commit ${current.hash} refers to parent ${current.parent}, which is not in this ` +
          `document. The history is incomplete.`
      );
    }
    current = parent;
  }

  throw new Error(
    `No snapshot found below commit ${hash}, so its content cannot be rebuilt. ` +
      `The history is damaged.`
  );
}

/** Rebuild the OOXML at a given commit by replaying its chain. */
export function reconstruct(repo: RepoXml, hash: string, diff: DiffEngine): string {
  const chain = replayChain(repo, hash);
  // replayChain stops at the first snapshot it reaches, so chain[0] is always
  // the snapshot and everything after it is a patch.
  let ooxml = decodeDelta(chain[0].delta, chain[0].encoding);
  for (let i = 1; i < chain.length; i++) {
    ooxml = diff.applyPatch(ooxml, decodeDelta(chain[i].delta, chain[i].encoding));
  }
  return ooxml;
}

/** Both parents of a commit that exist as objects, first parent first. */
function parentsOf(repo: RepoXml, c: CommitXml): CommitXml[] {
  const out: CommitXml[] = [];
  for (const h of [c.parent, c.parent2]) {
    if (!h || h === ROOT_PARENT) continue;
    const p = findCommit(repo, h);
    if (p) out.push(p);
  }
  return out;
}

/** Every commit reachable from `hash` by following parent pointers, including itself. */
export function ancestors(repo: RepoXml, hash: string): Set<string> {
  const out = new Set<string>();
  const stack = [hash];
  while (stack.length > 0) {
    const h = stack.pop()!;
    if (!h || h === ROOT_PARENT || out.has(h)) continue;
    const c = findCommit(repo, h);
    if (!c) continue;
    out.add(h);
    for (const p of parentsOf(repo, c)) stack.push(p.hash);
  }
  return out;
}

/**
 * The merge base of two commits: the ancestor of `theirs` nearest to it that
 * is also an ancestor of `ours`.
 *
 * Breadth-first from `theirs` rather than a straight parent walk, because a
 * merge commit has two parents and the base may lie up the second one.
 */
export function mergeBase(repo: RepoXml, ourTip: string, theirTip: string): string | null {
  const ours = ancestors(repo, ourTip);
  const queue = [theirTip];
  const seen = new Set<string>();
  while (queue.length > 0) {
    const h = queue.shift()!;
    if (!h || h === ROOT_PARENT || seen.has(h)) continue;
    seen.add(h);
    if (ours.has(h)) return h;
    const c = findCommit(repo, h);
    if (!c) continue;
    for (const p of parentsOf(repo, c)) queue.push(p.hash);
  }
  return null;
}

export interface CommitInfo {
  hash: string;
  parent: string;
  parent2: string;
  author: string;
  timestamp: string;
  message: string;
  encoding: Encoding;
}

function toInfo(c: CommitXml): CommitInfo {
  return {
    hash: c.hash,
    parent: c.parent,
    parent2: c.parent2,
    author: c.author,
    timestamp: c.timestamp,
    message: c.message,
    encoding: c.encoding,
  };
}

function timeOf(c: CommitXml): number {
  const t = Date.parse(c.timestamp);
  return Number.isNaN(t) ? 0 : t;
}

/**
 * The commit log for a tip, newest first.
 *
 * Both parents are followed, so commits that arrived through a merge are
 * listed rather than vanishing -- the original walked the first-parent chain
 * only, and since merges recorded no second parent at all, the entire merged-in
 * branch was invisible in the log even though its work was in the document.
 *
 * The order is topological with the timestamp as the tie-break: of the commits
 * whose children have all been listed, the newest goes next. Sorting purely by
 * timestamp would be simpler, but a clock that is wrong by a minute -- or two
 * machines editing the same document -- would then print a commit above its own
 * parent, which is the one thing a history view must never do.
 */
export function logFrom(repo: RepoXml, tip: string): CommitInfo[] {
  const reachable = ancestors(repo, tip);

  // How many not-yet-listed children each commit has, counting only children
  // inside the reachable set.
  const pending = new Map<string, number>();
  for (const h of reachable) pending.set(h, 0);
  for (const h of reachable) {
    const c = findCommit(repo, h)!;
    for (const p of parentsOf(repo, c)) {
      if (pending.has(p.hash)) pending.set(p.hash, pending.get(p.hash)! + 1);
    }
  }

  const out: CommitInfo[] = [];
  const ready = new Set<string>();
  for (const [h, n] of pending) if (n === 0) ready.add(h);

  while (ready.size > 0) {
    let pick: CommitXml | null = null;
    for (const h of ready) {
      const c = findCommit(repo, h)!;
      if (!pick || timeOf(c) > timeOf(pick) || (timeOf(c) === timeOf(pick) && c.hash > pick.hash)) {
        pick = c;
      }
    }
    ready.delete(pick!.hash);
    out.push(toInfo(pick!));
    for (const p of parentsOf(repo, pick!)) {
      if (!pending.has(p.hash)) continue;
      const left = pending.get(p.hash)! - 1;
      pending.set(p.hash, left);
      if (left === 0) ready.add(p.hash);
    }
  }

  return out;
}

/** The commit log for the branch HEAD points at, newest first. */
export function logFor(repo: RepoXml): CommitInfo[] {
  const ref = currentRef(repo);
  if (!ref || !ref.tip) return [];
  return logFrom(repo, ref.tip);
}
