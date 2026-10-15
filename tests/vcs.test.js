"use strict";

// The whole thing, end to end, over the in-memory document in harness.js:
// init, commit, branch, switch, merge, log, restore. None of this was runnable
// outside Word before, which is why none of it was tested.

const test = require("node:test");
const assert = require("node:assert/strict");

const { makeVcs, makeOoxml, rejectsWith, FakeParts, RepoFormat } = require("./harness");

const V1 = makeOoxml(["Introduction", "Body text", "Conclusion"]);

function paragraphs(ooxml) {
  return [...ooxml.matchAll(/<w:t>([^<]*)<\/w:t>/g)].map((m) => m[1]);
}

async function started(ooxml = V1) {
  const vcs = makeVcs({ ooxml });
  await vcs.commits.init();
  return vcs;
}

// ── Init and commit ──────────────────────────────────────────────────────────

test("init creates main with a root snapshot of the document", async () => {
  const vcs = await started();
  const repo = await vcs.store.load();
  assert.deepEqual(repo.refs.map((r) => r.name), ["main"]);
  assert.equal(repo.head, "refs/heads/main");
  assert.equal(repo.objects.length, 1);
  assert.equal(repo.objects[0].parent, "00root");
  assert.equal(repo.objects[0].message, "Initial commit");
  assert.equal(await vcs.commits.reconstruct(repo.objects[0].hash), V1);
});

test("init is idempotent", async () => {
  const vcs = await started();
  await vcs.commits.init();
  assert.equal((await vcs.store.load()).objects.length, 1);
});

test("init records the document's author when the host supplies one", async () => {
  const vcs = makeVcs({ ooxml: V1 });
  vcs.doc.author = "Grace Hopper";
  await vcs.commits.init();
  const repo = await vcs.store.load();
  assert.equal(repo.author, "Grace Hopper");
  assert.equal(repo.objects[0].author, "Grace Hopper");
});

test("init falls back to a placeholder when the host will not say", async () => {
  assert.equal((await (await started()).store.load()).author, "Author");
});

test("a commit reconstructs to exactly what was in the document", async () => {
  const vcs = await started();
  const v2 = makeOoxml(["Introduction", "Body text, revised", "Conclusion"]);
  vcs.doc.ooxml = v2;
  const hash = await vcs.commits.commit("Revise the body");
  assert.equal(await vcs.commits.reconstruct(hash), v2);
});

test("a long chain of commits each reconstruct correctly", async () => {
  const vcs = await started();
  const expected = new Map();
  for (let i = 1; i <= 12; i++) {
    const text = Array.from({ length: i }, (_, j) => `paragraph ${j + 1}`);
    vcs.doc.ooxml = makeOoxml(text);
    expected.set(await vcs.commits.commit(`commit ${i}`), vcs.doc.ooxml);
  }
  for (const [hash, ooxml] of expected) {
    assert.equal(await vcs.commits.reconstruct(hash), ooxml, `commit ${hash}`);
  }
});

test("committing without a repository says so", async () => {
  const vcs = makeVcs({ ooxml: V1 });
  await rejectsWith(() => vcs.commits.commit("x"), "Did you Init?");
});

test("two commits in the same millisecond get different identifiers", async () => {
  const vcs = await started();
  const realNow = Date.now;
  try {
    Date.now = () => 1_800_000_000_000;
    vcs.doc.ooxml = makeOoxml(["a"]);
    const first = await vcs.commits.commit("one");
    vcs.doc.ooxml = makeOoxml(["b"]);
    const second = await vcs.commits.commit("two");
    assert.notEqual(first, second);
  } finally {
    Date.now = realNow;
  }
});

test("a commit identifier is never reused even if the random tail repeats", async () => {
  const vcs = await started();
  const realNow = Date.now;
  const realRandom = Math.random;
  try {
    Date.now = () => 1_800_000_000_000;
    Math.random = () => 0.5; // always the same tail
    vcs.doc.ooxml = makeOoxml(["a"]);
    const first = await vcs.commits.commit("one");
    let attempts = 0;
    Math.random = () => {
      attempts++;
      return attempts > 2 ? 0.75 : 0.5;
    };
    vcs.doc.ooxml = makeOoxml(["b"]);
    const second = await vcs.commits.commit("two");
    assert.notEqual(first, second);
  } finally {
    Date.now = realNow;
    Math.random = realRandom;
  }
});

// ── Restore ──────────────────────────────────────────────────────────────────

test("restoring puts the old content back, as tracked changes", async () => {
  const vcs = await started();
  const root = (await vcs.store.load()).objects[0].hash;
  vcs.doc.ooxml = makeOoxml(["Introduction", "Rewritten", "Conclusion"]);
  await vcs.commits.commit("rewrite");

  await vcs.commits.checkout(root);
  assert.equal(vcs.doc.ooxml, V1);
  assert.equal(vcs.doc.writes.at(-1).tracked, true);
});

test("switching branches replaces content without tracking", async () => {
  const vcs = await started();
  await vcs.branches.createBranch("draft");
  vcs.doc.ooxml = makeOoxml(["Introduction", "Main only", "Conclusion"]);
  await vcs.commits.commit("main work");

  const tip = await vcs.branches.switchBranch("draft");
  await vcs.commits.checkoutSilent(tip);
  assert.equal(vcs.doc.ooxml, V1);
  assert.equal(vcs.doc.writes.at(-1).tracked, false);
});

// ── Branches ─────────────────────────────────────────────────────────────────

test("a new branch starts at the current tip", async () => {
  const vcs = await started();
  const tip = (await vcs.store.load()).refs[0].tip;
  await vcs.branches.createBranch("draft");
  assert.equal(await vcs.branches.getTip("draft"), tip);
});

test("creating a branch that exists is refused", async () => {
  const vcs = await started();
  await vcs.branches.createBranch("draft");
  await rejectsWith(() => vcs.branches.createBranch("draft"), "already exists");
});

test("creating a branch before init is refused", async () => {
  const vcs = makeVcs({ ooxml: V1 });
  await rejectsWith(() => vcs.branches.createBranch("draft"), "no repository");
});

test("a branch name is trimmed, and control characters are refused", async () => {
  const vcs = await started();
  await vcs.branches.createBranch("  spaced  ");
  assert.ok((await vcs.branches.listBranches()).some((r) => r.name === "spaced"));
  await rejectsWith(() => vcs.branches.createBranch("a\nb"), "control characters");
  await rejectsWith(() => vcs.branches.createBranch("   "), "needs a name");
});

test("a branch name containing a quote survives a round trip", async () => {
  const vcs = await started();
  await vcs.branches.createBranch(`say "hi"`);
  assert.equal(await vcs.branches.currentBranch(), "main");
  assert.ok((await vcs.branches.listBranches()).some((r) => r.name === `say "hi"`));
});

test("switching to a branch that does not exist is refused", async () => {
  const vcs = await started();
  await rejectsWith(() => vcs.branches.switchBranch("nope"), "not found");
});

// ── Merge ────────────────────────────────────────────────────────────────────

/** main and draft each edit a different paragraph of the same document. */
async function diverged() {
  const vcs = await started(makeOoxml(["Introduction", "Body text", "Conclusion"]));
  await vcs.branches.createBranch("draft");

  vcs.doc.ooxml = makeOoxml(["Introduction rewritten on main", "Body text", "Conclusion"]);
  await vcs.commits.commit("main edits the introduction");

  const tip = await vcs.branches.switchBranch("draft");
  await vcs.commits.checkoutSilent(tip);
  vcs.doc.ooxml = makeOoxml(["Introduction", "Body text", "Conclusion rewritten on draft"]);
  await vcs.commits.commit("draft edits the conclusion");

  await vcs.branches.switchBranch("main");
  await vcs.commits.checkoutSilent(await vcs.branches.getTip("main"));
  return vcs;
}

test("a merge of disjoint edits keeps both", async () => {
  const vcs = await diverged();
  await vcs.commits.merge("draft");
  assert.deepEqual(paragraphs(vcs.doc.ooxml), [
    "Introduction rewritten on main",
    "Body text",
    "Conclusion rewritten on draft",
  ]);
});

test("a merge commit records both parents", async () => {
  const vcs = await diverged();
  const mainBefore = await vcs.branches.getTip("main");
  const draftTip = await vcs.branches.getTip("draft");
  const hash = await vcs.commits.merge("draft");

  const repo = await vcs.store.load();
  const merge = repo.objects.find((c) => c.hash === hash);
  assert.equal(merge.parent, mainBefore);
  assert.equal(merge.parent2, draftTip);
});

test("after a merge the log shows the work from both branches", async () => {
  const vcs = await diverged();
  await vcs.commits.merge("draft");
  const messages = (await vcs.commits.log()).map((c) => c.message);
  assert.ok(messages.includes("main edits the introduction"), messages.join(" | "));
  assert.ok(messages.includes("draft edits the conclusion"), messages.join(" | "));
  assert.ok(messages.some((m) => m.startsWith("Merge branch 'draft'")), messages.join(" | "));
});

test("merging the same branch twice is refused instead of replaying it", async () => {
  const vcs = await diverged();
  await vcs.commits.merge("draft");
  await rejectsWith(() => vcs.commits.merge("draft"), "already merged");
});

test("a merge result reconstructs from its own snapshot", async () => {
  const vcs = await diverged();
  const hash = await vcs.commits.merge("draft");
  assert.equal(await vcs.commits.reconstruct(hash), vcs.doc.ooxml);
});

test("a commit after a merge is a delta against the merge", async () => {
  const vcs = await diverged();
  await vcs.commits.merge("draft");
  vcs.doc.ooxml = makeOoxml(["One more paragraph"]);
  const hash = await vcs.commits.commit("after the merge");
  assert.equal(await vcs.commits.reconstruct(hash), vcs.doc.ooxml);
});

test("a fast-forward moves the pointer and checks the content out silently", async () => {
  const vcs = await started();
  await vcs.branches.createBranch("draft");
  const tip = await vcs.branches.switchBranch("draft");
  await vcs.commits.checkoutSilent(tip);
  vcs.doc.ooxml = makeOoxml(["Introduction", "Draft work", "Conclusion"]);
  const draftTip = await vcs.commits.commit("draft work");

  await vcs.branches.switchBranch("main");
  const result = await vcs.commits.merge("draft");

  assert.equal(result, draftTip);
  assert.equal(await vcs.branches.getTip("main"), draftTip);
  assert.deepEqual(paragraphs(vcs.doc.ooxml), ["Introduction", "Draft work", "Conclusion"]);
  assert.equal(vcs.doc.writes.at(-1).tracked, false);
});

test("merging a branch that does not exist is refused", async () => {
  const vcs = await started();
  await rejectsWith(() => vcs.commits.merge("nope"), `"nope" not found`);
});

test("merging a branch at the same commit is refused", async () => {
  const vcs = await started();
  await vcs.branches.createBranch("draft");
  await rejectsWith(() => vcs.commits.merge("draft"), "already at the same commit");
});

test("a conflicting merge leaves the document and both branches alone", async () => {
  const vcs = await started(makeOoxml(["Introduction", "The board approved the budget.", "End"]));
  await vcs.branches.createBranch("draft");

  vcs.doc.ooxml = makeOoxml(["Introduction", "The board approved the revised budget.", "End"]);
  const mainTip = await vcs.commits.commit("main rewrites the sentence");

  const tip = await vcs.branches.switchBranch("draft");
  await vcs.commits.checkoutSilent(tip);
  vcs.doc.ooxml = makeOoxml(["Introduction", "The committee rejected the budget.", "End"]);
  const draftTip = await vcs.commits.commit("draft rewrites the same sentence");

  await vcs.branches.switchBranch("main");
  await vcs.commits.checkoutSilent(mainTip);
  const before = vcs.doc.ooxml;
  const writes = vcs.doc.writes.length;

  await rejectsWith(() => vcs.commits.merge("draft"), "CONFLICT");

  assert.equal(vcs.doc.ooxml, before, "the document was not touched");
  assert.equal(vcs.doc.writes.length, writes, "nothing was written to the document");
  assert.equal(await vcs.branches.getTip("main"), mainTip);
  assert.equal(await vcs.branches.getTip("draft"), draftTip);
  assert.equal((await vcs.store.load()).objects.length, 3, "no merge commit was recorded");
});

// ── Damaged history ──────────────────────────────────────────────────────────

test("init over an unreadable repository refuses instead of replacing it", async () => {
  const vcs = await started();
  const xml = [...(await vcs.parts.list())][0].xml;
  const damaged = new FakeParts();
  damaged.seed(xml.slice(0, xml.length - 40));

  const broken = makeVcs({ parts: damaged, ooxml: V1 });
  const e = await rejectsWith(() => broken.commits.init(), "not a repository");
  assert.ok(e instanceof RepoFormat.RepoFormatError);
  assert.deepEqual(damaged.log, [], "nothing was written");
  assert.equal(damaged.parts.size, 1, "the damaged part is still there");
});

test("each commit records who made it, not just who initialised the repository", async () => {
  const vcs = makeVcs({ ooxml: V1 });
  vcs.doc.author = "Ada Lovelace";
  await vcs.commits.init();

  vcs.doc.author = "Grace Hopper";
  vcs.doc.ooxml = makeOoxml(["Introduction", "Grace's paragraph", "Conclusion"]);
  await vcs.commits.commit("Grace edits the body");

  const log = await vcs.commits.log();
  assert.deepEqual(log.map((c) => c.author), ["Grace Hopper", "Ada Lovelace"]);
  assert.equal((await vcs.store.load()).author, "Ada Lovelace", "the config author is unchanged");
});

test("a commit falls back to the stored author when the host will not say", async () => {
  const vcs = makeVcs({ ooxml: V1 });
  vcs.doc.author = "Ada Lovelace";
  await vcs.commits.init();

  vcs.doc.author = null;
  vcs.doc.ooxml = makeOoxml(["changed"]);
  await vcs.commits.commit("anonymous");
  assert.equal((await vcs.commits.log())[0].author, "Ada Lovelace");
});
