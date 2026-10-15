"use strict";

// Proof that each fix fixes something.
//
// Every assertion here runs the original code, transcribed in original.js, and
// asserts the wrong answer it actually gives. The matching assertion against
// the fixed code lives in the suite for whichever module owns it, and is named
// the same way, so a reader can put the two side by side.

const test = require("node:test");
const assert = require("node:assert/strict");
const { execFileSync } = require("node:child_process");
const path = require("node:path");

const original = require("./original");
const { parse, RepoFormat, DiffEngine, makeOoxml } = require("./harness");

const REPO_ROOT = path.join(__dirname, "..");

function baselineSource(file) {
  return execFileSync("git", ["show", `${original.BASELINE_COMMIT}:${file}`], {
    cwd: REPO_ROOT,
    encoding: "utf8",
    maxBuffer: 8 * 1024 * 1024,
  });
}

test("the transcription still matches the baseline commit", () => {
  for (const [file, lines] of Object.entries(original.BASELINE_LINES)) {
    const src = baselineSource(file);
    for (const line of lines) {
      assert.ok(
        src.includes(line),
        `${file} in ${original.BASELINE_COMMIT} no longer contains:\n${line}`
      );
    }
  }
});

// ── Damaged history read as an empty document ────────────────────────────────

test("original: a truncated repository reads as a document with no repository", () => {
  const repo = {
    head: "refs/heads/main",
    author: "Ada",
    refs: [{ name: "main", tip: "aaa" }],
    objects: [
      {
        hash: "aaa",
        parent: "00root",
        author: "Ada",
        timestamp: "2026-01-01T00:00:00.000Z",
        message: "Initial commit",
        delta: "<pkg/>",
        encoding: "snapshot",
      },
    ],
  };
  const xml = original.repoToXml(repo);

  // Truncate mid-document, as a converter or a failed write would.
  const damaged = xml.slice(0, xml.length - 40);

  const read = original.xmlToRepo(damaged);
  assert.equal(read.refs.length, 0, "the original reported no branches");
  assert.equal(read.objects.length, 0, "the original reported no commits");
  assert.equal(read.head, "refs/heads/main", "and fell back to the default HEAD");
  // refs.length === 0 is exactly what checkRepoExists() used to decide the
  // document had never been initialised, which enabled the Init button.
});

test("original: a Custom XML Part that is not a repository reads as an empty one", () => {
  const read = original.xmlToRepo(`<?xml version="1.0"?><properties><title>Q3</title></properties>`);
  assert.deepEqual(read, {
    head: "refs/heads/main",
    author: "Author",
    refs: [],
    objects: [],
  });
});

test("fixed: both of those are refused rather than reported as empty", () => {
  const repo = {
    head: "refs/heads/main",
    author: "Ada",
    refs: [{ name: "main", tip: "aaa" }],
    objects: [
      {
        hash: "aaa",
        parent: "00root",
        parent2: "",
        author: "Ada",
        timestamp: "2026-01-01T00:00:00.000Z",
        message: "Initial commit",
        delta: "<pkg/>",
        encoding: "snapshot",
      },
    ],
  };
  const xml = RepoFormat.repoToXml(repo);

  assert.throws(
    () => RepoFormat.parseRepoXml(xml.slice(0, xml.length - 40), parse),
    RepoFormat.RepoFormatError
  );
  assert.throws(
    () => RepoFormat.parseRepoXml(`<?xml version="1.0"?><properties/>`, parse),
    RepoFormat.RepoFormatError
  );
});

// ── Carriage returns in stored content ───────────────────────────────────────

test("original: a payload containing CRLF does not survive a save and load", () => {
  const payload = "<w:p>one</w:p>\r\n<w:p>two</w:p>";
  const repo = {
    head: "refs/heads/main",
    author: "Ada",
    refs: [],
    objects: [
      {
        hash: "aaa",
        parent: "00root",
        author: "Ada",
        timestamp: "t",
        message: "line one\r\nline two",
        delta: payload,
        encoding: "snapshot",
      },
    ],
  };
  const read = original.xmlToRepo(original.repoToXml(repo));
  assert.notEqual(read.objects[0].delta, payload);
  assert.equal(read.objects[0].delta, "<w:p>one</w:p>\n<w:p>two</w:p>");
  assert.equal(read.objects[0].message, "line one\nline two");
});

test("fixed: a payload containing CRLF survives a save and load", () => {
  const payload = "<w:p>one</w:p>\r\n<w:p>two</w:p>";
  const repo = {
    head: "refs/heads/main",
    author: "Ada",
    refs: [],
    objects: [
      {
        hash: "aaa",
        parent: "00root",
        parent2: "",
        author: "Ada",
        timestamp: "t",
        message: "line one\r\nline two",
        delta: payload,
        encoding: "snapshot",
      },
    ],
  };
  const read = RepoFormat.parseRepoXml(RepoFormat.repoToXml(repo), parse);
  assert.equal(read.objects[0].delta, payload);
  assert.equal(read.objects[0].message, "line one\r\nline two");
});

// ── Injection from document-carried strings ──────────────────────────────────

const PAYLOAD = `" onmouseover="alert(1)" x="`;

test("original: a commit message escapes out of the log card", () => {
  const html = original.logCardHtml(
    { hash: "abc123", parent: "00root", message: `<img src=x onerror="alert(1)">`, author: "Ada" },
    true,
    "1 Jan"
  );
  assert.ok(html.includes("<img src=x onerror="), "the tag was written through unescaped");
});

test("original: a branch name escapes out of its attribute", () => {
  const html = original.branchRowHtml({ name: PAYLOAD, tip: "x" }, false);
  assert.ok(html.includes(`data-switch="" onmouseover="alert(1)"`), html);
});

test("original: an error message escapes out of the toast", () => {
  const html = original.toastHtml("error", `<img src=x onerror="alert(1)">`);
  assert.ok(html.includes("<img src=x onerror="), html);
});

test("original: a crafted commit hash unbalances the graph SVG", () => {
  // The hash is the one graph field that was never escaped. It is sliced to
  // six characters before it is written, so it cannot carry a script -- but it
  // can still close the element it is inside, which corrupts the graph for
  // every commit below it. Six characters is exactly enough for "</text".
  const html = original.graphNodeHtml(
    { hash: "</text><script>alert(1)</script>", message: "hi" },
    10,
    10,
    "#000",
    "#666"
  );
  assert.ok(html.includes("></text</text>"), html);
  assert.ok(!html.includes("<script>"), "six characters is not enough for a payload");
});

// ── Structural OOXML attributes stripped by the merge ────────────────────────

const BOOKMARKED =
  `<w:p><w:bookmarkStart w:id="4" w:name="intro"/><w:r w:rsidR="00AB12">` +
  `<w:t>Hello</w:t></w:r><w:bookmarkEnd w:id="4"/></w:p>`;

test("original: normalising OOXML unpairs bookmarks and strips comment ids", () => {
  const out = original.normaliseOoxml(BOOKMARKED);
  assert.ok(!out.includes(`w:id="4"`), "the bookmark ids were removed");
  assert.ok(out.includes(`<w:bookmarkStart w:name="intro"/>`), out);
  assert.ok(out.includes(`<w:bookmarkEnd/>`), out);
});

test("original: the paragraph-id alternatives never matched real OOXML", () => {
  // Real documents write w14:paraId, in the 2010 extension namespace. The
  // original's pattern was anchored to the w: prefix, so it matched neither.
  const src = `<w:p w14:paraId="12345678" w14:textId="77777777"><w:r><w:t>x</w:t></w:r></w:p>`;
  assert.equal(original.normaliseOoxml(src), src, "nothing was stripped");
});

// ── Merges that should conflict and do not ───────────────────────────────────

test("original: two branches rewriting the same paragraph merge silently", () => {
  const diff = new DiffEngine();
  const base = makeOoxml(["Introduction", "The quick brown fox jumps over the lazy dog.", "End"]);
  const ours = base.replace(
    "The quick brown fox jumps over the lazy dog.",
    "The quick brown fox leaps over the idle dog."
  );
  const theirs = base.replace(
    "The quick brown fox jumps over the lazy dog.",
    "A slow grey badger ambles past the lazy dog."
  );

  const merged = original.threeWayMerge(diff, base, ours, theirs);

  // No conflict was raised. The sentence in the merged document is one that
  // neither branch contains and nobody wrote: their subject, our object.
  const text = [...merged.matchAll(/<w:t>([^<]*)<\/w:t>/g)].map((m) => m[1]);
  assert.deepEqual(text, [
    "Introduction",
    "A slow grey badger ambles past the idle dog.",
    "End",
  ]);
  assert.ok(!ours.includes(text[1]) && !theirs.includes(text[1]));
});

// ── Commit identifiers ───────────────────────────────────────────────────────

test("original: sha6 loses entropy whenever Math.random is short", () => {
  const real = Math.random;
  try {
    // 0.5 renders as "0.i" in base 36, so slice(2, 6) yields one character.
    Math.random = () => 0.5;
    const id = original.sha6();
    assert.ok(id.endsWith("i000"), `expected three padding zeroes, got ${id}`);
  } finally {
    Math.random = real;
  }
});

// ── The graph could not represent a merge ────────────────────────────────────

test("original: GraphCommit had no second parent, so no merge edge was possible", () => {
  const src = baselineSource("src/vcs/GraphBuilder.ts");
  const iface = src.slice(src.indexOf("export interface GraphCommit"));
  const body = iface.slice(0, iface.indexOf("}"));
  assert.ok(!body.includes("parent2"), "the baseline interface already had a second parent");
  assert.ok(
    !baselineSource("src/vcs/CommitManager.ts").includes("parent2"),
    "the baseline merge already recorded a second parent"
  );
});
