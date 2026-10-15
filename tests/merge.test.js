"use strict";

// The three-way merge, and what it refuses.
//
// The property under test is not "does diff-match-patch manage to apply both
// patch sets" -- it usually does, which was the original's undoing -- but
// "did the two sides edit regions that do not touch".

const test = require("node:test");
const assert = require("node:assert/strict");

const { load, DiffEngine, makeOoxml } = require("./harness");
const { MergeManager, MergeConflictError } = load("vcs/MergeManager.js");
const { normaliseOoxml, rangesOverlap } = load("vcs/DiffEngine.js");

const merger = () => new MergeManager(new DiffEngine());

function text(ooxml) {
  return [...ooxml.matchAll(/<w:t>([^<]*)<\/w:t>/g)].map((m) => m[1]);
}

function conflicts(base, ours, theirs) {
  let caught = null;
  try {
    merger().threeWayMerge(base, ours, theirs);
  } catch (e) {
    caught = e;
  }
  assert.ok(caught instanceof MergeConflictError, `expected a conflict, got ${caught}`);
  assert.ok(caught.message.startsWith("CONFLICT: "), caught.message);
  return caught;
}

const BASE = makeOoxml([
  "Introduction",
  "The quick brown fox jumps over the lazy dog.",
  "A second paragraph that neither side touches.",
  "Conclusion",
]);

// ── Trivial cases ────────────────────────────────────────────────────────────

test("a side that changed nothing yields the other side", () => {
  const ours = BASE.replace("Conclusion", "Conclusion, revised");
  assert.equal(merger().threeWayMerge(BASE, ours, BASE), normaliseOoxml(ours));
  assert.equal(merger().threeWayMerge(BASE, BASE, ours), normaliseOoxml(ours));
});

test("two sides that made the same change do not conflict with each other", () => {
  const both = BASE.replace("Conclusion", "Conclusion, revised");
  assert.equal(merger().threeWayMerge(BASE, both, both), normaliseOoxml(both));
});

// ── Disjoint edits ───────────────────────────────────────────────────────────

test("edits to different paragraphs both survive", () => {
  const ours = BASE.replace("Introduction", "Introduction, rewritten");
  const theirs = BASE.replace("Conclusion", "Conclusion, rewritten");
  const merged = merger().threeWayMerge(BASE, ours, theirs);
  assert.deepEqual(text(merged), [
    "Introduction, rewritten",
    "The quick brown fox jumps over the lazy dog.",
    "A second paragraph that neither side touches.",
    "Conclusion, rewritten",
  ]);
});

test("a merge of disjoint edits gives the same answer in either order", () => {
  const ours = BASE.replace("Introduction", "Introduction, rewritten");
  const theirs = BASE.replace("Conclusion", "Conclusion, rewritten");
  assert.equal(
    merger().threeWayMerge(BASE, ours, theirs),
    merger().threeWayMerge(BASE, theirs, ours)
  );
});

test("one side adding a paragraph and the other editing a distant one both survive", () => {
  const ours = makeOoxml([
    "Inserted at the top",
    "Introduction",
    "The quick brown fox jumps over the lazy dog.",
    "A second paragraph that neither side touches.",
    "Conclusion",
  ]);
  const theirs = BASE.replace("Conclusion", "Conclusion, rewritten");
  const merged = merger().threeWayMerge(BASE, ours, theirs);
  assert.deepEqual(text(merged), [
    "Inserted at the top",
    "Introduction",
    "The quick brown fox jumps over the lazy dog.",
    "A second paragraph that neither side touches.",
    "Conclusion, rewritten",
  ]);
});

// ── Overlapping edits ────────────────────────────────────────────────────────

test("two rewrites of the same sentence conflict rather than merging", () => {
  const ours = BASE.replace(
    "The quick brown fox jumps over the lazy dog.",
    "The quick brown fox leaps over the idle dog."
  );
  const theirs = BASE.replace(
    "The quick brown fox jumps over the lazy dog.",
    "A slow grey badger ambles past the lazy dog."
  );
  const e = conflicts(BASE, ours, theirs);
  assert.ok(e.message.includes("same part of the document"), e.message);
});

test("two deletions of the same paragraph conflict", () => {
  const drop = (s) => s.replace(/<w:p[^>]*>(?:(?!<\/w:p>).)*A second paragraph[\s\S]*?<\/w:p>/, "");
  const ours = drop(BASE);
  const theirs = drop(BASE).replace("Conclusion", "Conclusion, rewritten");
  // Ours removes the paragraph; theirs removes it and edits elsewhere. The
  // shared deletion is the same region, so this is a conflict, not a merge.
  conflicts(BASE, ours, theirs);
});

test("one side editing text the other deleted conflicts", () => {
  const ours = BASE.replace("A second paragraph that neither side touches.", "");
  const theirs = BASE.replace(
    "A second paragraph that neither side touches.",
    "A second paragraph, now revised."
  );
  conflicts(BASE, ours, theirs);
});

test("edits close enough to share diff context conflict", () => {
  // diff-match-patch keeps four characters of context around each hunk. Two
  // edits inside that window cannot be placed independently, and being
  // conservative here is the point.
  const ours = BASE.replace("lazy dog.", "lazy fox.");
  const theirs = BASE.replace("lazy dog.", "lazy cat.");
  conflicts(BASE, ours, theirs);
});

test("a conflict message says what to do about it", () => {
  const ours = BASE.replace("brown fox jumps", "brown fox leaps");
  const theirs = BASE.replace("brown fox jumps", "grey badger ambles");
  const e = conflicts(BASE, ours, theirs);
  assert.ok(/by hand/.test(e.message), e.message);
});

// ── Normalisation ────────────────────────────────────────────────────────────

test("revision-save identifiers do not make a merge look like a rewrite", () => {
  const ours = BASE.replace(/w:rsidR="[^"]*"/g, 'w:rsidR="FFFFFF"');
  const theirs = BASE.replace("Conclusion", "Conclusion, rewritten");
  const merged = merger().threeWayMerge(BASE, ours, theirs);
  assert.deepEqual(text(merged), [
    "Introduction",
    "The quick brown fox jumps over the lazy dog.",
    "A second paragraph that neither side touches.",
    "Conclusion, rewritten",
  ]);
});

test("paragraph identifiers are stripped whatever prefix they carry", () => {
  assert.equal(
    normaliseOoxml(`<w:p w14:paraId="12345678" w14:textId="77777777"><w:r/></w:p>`),
    `<w:p><w:r/></w:p>`
  );
  assert.equal(normaliseOoxml(`<w:p w:rsidR="00AB" w:rsidRDefault="00CD"/>`), `<w:p/>`);
});

test("bookmark, comment and revision ids survive normalisation", () => {
  const src =
    `<w:p><w:bookmarkStart w:id="4" w:name="intro"/>` +
    `<w:ins w:id="7" w:author="Ada"><w:r><w:t>Hello</w:t></w:r></w:ins>` +
    `<w:commentRangeEnd w:id="2"/><w:bookmarkEnd w:id="4"/></w:p>`;
  assert.equal(normaliseOoxml(src), src);
});

test("a merged document keeps its bookmark pairing", () => {
  const base =
    `<w:body><w:p><w:bookmarkStart w:id="4" w:name="intro"/><w:r w:rsidR="001"><w:t>Alpha</w:t>` +
    `</w:r><w:bookmarkEnd w:id="4"/></w:p><w:p><w:r><w:t>Omega</w:t></w:r></w:p></w:body>`;
  const ours = base.replace("Alpha", "Alpha, revised");
  const theirs = base.replace("Omega", "Omega, revised");
  const merged = merger().threeWayMerge(base, ours, theirs);

  const starts = (merged.match(/<w:bookmarkStart[^>]*w:id="4"/g) || []).length;
  const ends = (merged.match(/<w:bookmarkEnd[^>]*w:id="4"/g) || []).length;
  assert.equal(starts, 1);
  assert.equal(ends, 1);
});

// ── The overlap predicate itself ─────────────────────────────────────────────

test("rangesOverlap is true only for ranges that actually intersect", () => {
  assert.equal(rangesOverlap([{ start: 0, end: 10 }], [{ start: 10, end: 20 }]), false);
  assert.equal(rangesOverlap([{ start: 0, end: 11 }], [{ start: 10, end: 20 }]), true);
  assert.equal(rangesOverlap([{ start: 5, end: 6 }], [{ start: 0, end: 20 }]), true);
  assert.equal(rangesOverlap([], [{ start: 0, end: 20 }]), false);
  assert.equal(
    rangesOverlap(
      [
        { start: 0, end: 5 },
        { start: 30, end: 40 },
      ],
      [{ start: 35, end: 36 }]
    ),
    true
  );
});
