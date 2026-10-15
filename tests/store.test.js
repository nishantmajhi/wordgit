"use strict";

// ObjectStore: when it is safe to write to the document, and what happens when
// it is not. The rules here are the ones that decide whether a user's history
// survives a bad day.

const test = require("node:test");
const assert = require("node:assert/strict");

const { FakeParts, makeStore, RepoFormat, rejectsWith } = require("./harness");
const { repoToXml, emptyRepo, RepoFormatError } = RepoFormat;

function repoWith(hashes, author = "Ada") {
  return {
    head: "refs/heads/main",
    author,
    refs: [{ name: "main", tip: hashes[hashes.length - 1] ?? "" }],
    objects: hashes.map((hash, i) => ({
      hash,
      parent: i === 0 ? "00root" : hashes[i - 1],
      parent2: "",
      author,
      timestamp: `2026-01-0${i + 1}T09:00:00.000Z`,
      message: `commit ${hash}`,
      delta: "<pkg/>",
      encoding: "snapshot",
    })),
  };
}

test("a document with no parts loads as an empty repository", async () => {
  const store = makeStore(new FakeParts());
  assert.deepEqual(await store.load(), emptyRepo());
  assert.equal(store.damaged, false);
});

test("a saved repository reads back", async () => {
  const parts = new FakeParts();
  const store = makeStore(parts);
  await store.load();
  const repo = repoWith(["aaa", "bbb"]);
  await store.save(repo);

  const reader = makeStore(parts);
  assert.deepEqual(await reader.load(), repo);
});

test("saving twice replaces the part in place rather than adding another", async () => {
  const parts = new FakeParts();
  const store = makeStore(parts);
  await store.load();
  await store.save(repoWith(["aaa"]));
  await store.save(repoWith(["aaa", "bbb"]));
  assert.equal(parts.parts.size, 1);
  assert.deepEqual(parts.log, ["add:part1", "set:part1"]);
});

test("a host that cannot replace in place adds before it removes", async () => {
  // The order matters. The original deleted the existing part and then added
  // the new one, so anything that went wrong in between -- an error, the
  // document closing, Word stopping -- left the document with no history.
  const parts = new FakeParts({ supportsSetXml: false });
  const store = makeStore(parts);
  await store.load();
  await store.save(repoWith(["aaa"]));
  await store.save(repoWith(["aaa", "bbb"]));

  assert.deepEqual(parts.log, ["add:part1", "add:part2", "remove:part1"]);
  assert.equal(parts.parts.size, 1);
});

test("a failed add leaves the previous history in place", async () => {
  const parts = new FakeParts({ supportsSetXml: false });
  const store = makeStore(parts);
  await store.load();
  await store.save(repoWith(["aaa"]));

  parts.failNextAdd = true;
  await rejectsWith(() => store.save(repoWith(["aaa", "bbb"])), "refused the new part");

  const reader = makeStore(parts);
  assert.deepEqual((await reader.load()).objects.map((c) => c.hash), ["aaa"]);
});

// ── Damaged history ──────────────────────────────────────────────────────────

test("a damaged part throws rather than reading as an empty repository", async () => {
  const xml = repoToXml(repoWith(["aaa"]));
  const store = makeStore(FakeParts.withXml(xml.slice(0, xml.length - 40)));
  await rejectsWith(() => store.load(), "not a repository");
  assert.equal(store.damaged, true);
});

test("a store that failed to load refuses to save over the damage", async () => {
  const xml = repoToXml(repoWith(["aaa"]));
  const parts = FakeParts.withXml(xml.slice(0, xml.length - 40));
  const store = makeStore(parts);

  await rejectsWith(() => store.load(), "not a repository");
  const e = await rejectsWith(() => store.save(repoWith(["zzz"])), "Refusing to write");
  assert.ok(e instanceof RepoFormatError);

  // And nothing was written: the damaged part is still there, untouched.
  assert.equal(parts.parts.size, 1);
  assert.deepEqual(parts.log, []);
});

test("two parts: the one with the longer history wins", async () => {
  // Debris from a host with no in-place replacement, where the add succeeded
  // and the remove did not. The newer write is always the longer one.
  const parts = new FakeParts();
  parts.seed(repoToXml(repoWith(["aaa"])));
  parts.seed(repoToXml(repoWith(["aaa", "bbb", "ccc"])));
  const store = makeStore(parts);
  assert.deepEqual((await store.load()).objects.map((c) => c.hash), ["aaa", "bbb", "ccc"]);
});

test("two parts, one damaged: the readable one is used and kept", async () => {
  const parts = new FakeParts();
  parts.seed("<wordgit:repo not xml at all");
  parts.seed(repoToXml(repoWith(["aaa", "bbb"])));
  const store = makeStore(parts);

  const repo = await store.load();
  assert.deepEqual(repo.objects.map((c) => c.hash), ["aaa", "bbb"]);
  assert.equal(store.damaged, false);

  repo.objects.push({ ...repo.objects[1], hash: "ccc", parent: "bbb" });
  await store.save(repo);
  assert.deepEqual(parts.log, ["set:part2"], "the damaged part was left alone");
});

test("a store that loaded cleanly reports itself undamaged", async () => {
  const store = makeStore(FakeParts.withXml(repoToXml(repoWith(["aaa"]))));
  await store.load();
  assert.equal(store.damaged, false);
});

test("the author name survives a save and load", async () => {
  const parts = new FakeParts();
  const store = makeStore(parts);
  await store.load();
  await store.save(repoWith(["aaa"], "Grace Hopper"));
  assert.equal((await makeStore(parts).load()).author, "Grace Hopper");
});
