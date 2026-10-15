"use strict";

// The commit-graph walks, and the payload encoding underneath them.

const test = require("node:test");
const assert = require("node:assert/strict");

const { HistoryModule, DiffEngine, makeOoxml } = require("./harness");
const {
  ancestors,
  branchNameFromHead,
  decodeDelta,
  encodeDelta,
  logFrom,
  mergeBase,
  reconstruct,
  replayChain,
} = HistoryModule;

function commit(hash, parent, over = {}) {
  return {
    hash,
    parent,
    parent2: "",
    author: "Ada",
    timestamp: "2026-01-01T09:00:00.000Z",
    message: hash,
    delta: "",
    encoding: "dmp",
    ...over,
  };
}

function repoOf(objects, refs = []) {
  return { head: "refs/heads/main", author: "Ada", refs, objects };
}

// ── Payload encoding ─────────────────────────────────────────────────────────

test("a compressed payload round-trips", () => {
  const raw = makeOoxml(["one", "two", "three"]);
  const { delta, encoding } = encodeDelta(raw, "snapshot");
  assert.equal(encoding, "snapshot-lz");
  assert.equal(decodeDelta(delta, encoding), raw);
});

test("compression is skipped when it would make the payload longer", () => {
  // Short strings come out of LZString's base64 form larger than they went in,
  // and the repository is rewritten whole on every commit, so storing the
  // larger of the two forms costs on every future write as well.
  const { delta, encoding } = encodeDelta("x", "dmp");
  assert.equal(encoding, "dmp");
  assert.equal(delta, "x");
  assert.ok(encodeDelta("x", "dmp").delta.length <= 1);
});

test("an empty compressed payload decodes to empty rather than failing", () => {
  assert.equal(decodeDelta("", "snapshot-lz"), "");
});

test("a corrupted compressed payload is reported, not returned as rubbish", () => {
  assert.throws(() => decodeDelta("!!!not base64!!!", "dmp-lz"), /damaged/);
});

test("branchNameFromHead strips only the refs/heads prefix", () => {
  assert.equal(branchNameFromHead("refs/heads/main"), "main");
  assert.equal(branchNameFromHead("refs/heads/feature/refs/heads/x"), "feature/refs/heads/x");
  assert.equal(branchNameFromHead("detached"), "detached");
});

// ── Replay ───────────────────────────────────────────────────────────────────

test("replay stops at the most recent snapshot rather than walking to the root", () => {
  const repo = repoOf([
    commit("a", "00root", { encoding: "snapshot", delta: "A" }),
    commit("b", "a"),
    commit("c", "b", { encoding: "snapshot", delta: "C" }),
    commit("d", "c"),
  ]);
  assert.deepEqual(
    replayChain(repo, "d").map((c) => c.hash),
    ["c", "d"]
  );
});

test("reconstructing replays the patches after the snapshot", () => {
  const diff = new DiffEngine();
  const v1 = makeOoxml(["one", "two"]);
  const v2 = makeOoxml(["one", "two changed"]);
  const v3 = makeOoxml(["one", "two changed", "three"]);

  const repo = repoOf([
    commit("a", "00root", { encoding: "snapshot", delta: v1 }),
    commit("b", "a", { delta: diff.makePatch(v1, v2) }),
    commit("c", "b", { delta: diff.makePatch(v2, v3) }),
  ]);

  assert.equal(reconstruct(repo, "a", diff), v1);
  assert.equal(reconstruct(repo, "b", diff), v2);
  assert.equal(reconstruct(repo, "c", diff), v3);
});

test("a chain with no snapshot under it is reported as damaged", () => {
  const repo = repoOf([commit("a", "00root"), commit("b", "a")]);
  assert.throws(() => replayChain(repo, "b"), /damaged/);
});

test("a missing parent is named rather than silently truncating the replay", () => {
  const repo = repoOf([commit("b", "missing")]);
  assert.throws(() => replayChain(repo, "b"), /refers to parent missing/);
});

test("a commit that is its own ancestor is reported rather than looping forever", () => {
  const repo = repoOf([commit("a", "b"), commit("b", "a")]);
  assert.throws(() => replayChain(repo, "a"), /its own ancestor/);
});

test("reconstructing a commit that is not there names it", () => {
  assert.throws(() => replayChain(repoOf([]), "nope"), /nope not found/);
});

// ── Ancestry ─────────────────────────────────────────────────────────────────
//
//   a ── b ── d(merge of c)
//    \        /
//     ── c ──

function diamond() {
  return repoOf(
    [
      commit("a", "00root", { encoding: "snapshot", delta: "A", timestamp: "2026-01-01T00:00:00Z" }),
      commit("b", "a", { timestamp: "2026-01-02T00:00:00Z" }),
      commit("c", "a", { timestamp: "2026-01-03T00:00:00Z" }),
      commit("d", "b", {
        parent2: "c",
        encoding: "snapshot",
        delta: "D",
        timestamp: "2026-01-04T00:00:00Z",
      }),
    ],
    [
      { name: "main", tip: "d" },
      { name: "side", tip: "c" },
    ]
  );
}

test("ancestors follow both parents of a merge", () => {
  assert.deepEqual([...ancestors(diamond(), "d")].sort(), ["a", "b", "c", "d"]);
});

test("ancestors of a commit before the merge do not include the other side", () => {
  assert.deepEqual([...ancestors(diamond(), "b")].sort(), ["a", "b"]);
});

test("the merge base of two branches is their fork point", () => {
  const repo = repoOf([
    commit("a", "00root"),
    commit("b", "a"),
    commit("c", "a"),
    commit("d", "b"),
  ]);
  assert.equal(mergeBase(repo, "d", "c"), "a");
});

test("the merge base of an already-merged branch is that branch's own tip", () => {
  // Which is how "already merged" is told apart from "needs merging".
  assert.equal(mergeBase(diamond(), "d", "c"), "c");
});

test("the merge base is found through a second parent", () => {
  //   a ── b ── d(merge of c) ── e      f branches off c
  const repo = repoOf([
    commit("a", "00root"),
    commit("b", "a"),
    commit("c", "a"),
    commit("d", "b", { parent2: "c" }),
    commit("e", "d"),
    commit("f", "c"),
  ]);
  assert.equal(mergeBase(repo, "e", "f"), "c");
});

test("unrelated histories have no merge base", () => {
  const repo = repoOf([commit("a", "00root"), commit("x", "00root")]);
  assert.equal(mergeBase(repo, "a", "x"), null);
});

// ── Log ──────────────────────────────────────────────────────────────────────

test("the log includes the commits that arrived through a merge", () => {
  assert.deepEqual(
    logFrom(diamond(), "d").map((c) => c.hash),
    ["d", "c", "b", "a"]
  );
});

test("the log never lists a commit above its own parent, whatever the clocks say", () => {
  // b claims to be older than its own parent a, which a wrong clock or a
  // second machine will produce. Sorting by timestamp alone would invert them.
  const repo = repoOf(
    [
      commit("a", "00root", { timestamp: "2026-06-01T00:00:00Z" }),
      commit("b", "a", { timestamp: "2020-01-01T00:00:00Z" }),
    ],
    [{ name: "main", tip: "b" }]
  );
  assert.deepEqual(
    logFrom(repo, "b").map((c) => c.hash),
    ["b", "a"]
  );
});

test("the log orders independent branches by date", () => {
  const repo = repoOf([
    commit("a", "00root", { timestamp: "2026-01-01T00:00:00Z" }),
    commit("b", "a", { timestamp: "2026-01-05T00:00:00Z" }),
    commit("c", "a", { timestamp: "2026-01-03T00:00:00Z" }),
    commit("d", "b", { parent2: "c", timestamp: "2026-01-06T00:00:00Z" }),
  ]);
  assert.deepEqual(
    logFrom(repo, "d").map((c) => c.hash),
    ["d", "b", "c", "a"]
  );
});

test("an unparseable timestamp does not drop a commit from the log", () => {
  const repo = repoOf([commit("a", "00root", { timestamp: "" }), commit("b", "a")]);
  assert.deepEqual(
    logFrom(repo, "b").map((c) => c.hash),
    ["b", "a"]
  );
});

test("the log of a tip that does not exist is empty rather than an error", () => {
  assert.deepEqual(logFrom(repoOf([]), "nope"), []);
});
