"use strict";

// The markup the task pane builds, and the one property that matters about it:
// content taken out of a document cannot change the shape of that markup.
//
// The shape check below is the whole test. It renders the same thing twice,
// once with a hostile string and once with an inert one of the same kind, and
// asserts the two outputs have the same elements with the same attributes in
// the same order. Content that cannot change the shape of the markup cannot
// inject anything, whatever it says, so there is no payload list to maintain.

const test = require("node:test");
const assert = require("node:assert/strict");

const { load } = require("./harness");
const { escapeHtml, escapeAttr, logCardHtml, branchRowHtml, toastHtml, formatTimestamp } =
  load("taskpane/render.js");

/**
 * Every element in a fragment as "tag[attr,attr]", in document order.
 *
 * The attribute scanner consumes quoted values rather than searching for
 * `name=`, because escaped content legitimately contains things like
 * `onmouseover=` *inside* a value, and a loose scanner fails markup that is
 * perfectly safe.
 */
function shape(html) {
  const out = [];
  const tagRe = /<\/?([a-zA-Z][a-zA-Z0-9-]*)((?:[^>"']|"[^"]*"|'[^']*')*)>/g;
  let m;
  while ((m = tagRe.exec(html)) !== null) {
    const attrs = [];
    const attrRe = /([a-zA-Z_:][-a-zA-Z0-9_:.]*)(\s*=\s*("[^"]*"|'[^']*'|[^\s>]*))?/g;
    let a;
    while ((a = attrRe.exec(m[2])) !== null) {
      if (a[0].trim()) attrs.push(a[1]);
    }
    out.push(`${m[0].startsWith("</") ? "/" : ""}${m[1]}[${attrs.join(",")}]`);
  }
  return out;
}

/** Payloads that between them reach every context a string can land in. */
const HOSTILE = [
  `<img src=x onerror="alert(1)">`,
  `" onmouseover="alert(1)" x="`,
  `' onmouseover='alert(1)' x='`,
  `</span><script>alert(1)</script><span>`,
  `</text><text onload="alert(1)">`,
  `&lt;not really a tag&gt;`,
  `a & b < c > d`,
];

const INERT = "harmless";

function assertSameShape(render) {
  const baseline = shape(render(INERT));
  for (const payload of HOSTILE) {
    assert.deepEqual(
      shape(render(payload)),
      baseline,
      `payload changed the markup shape: ${payload}`
    );
  }
}

// ── The shape check catches what it is for ───────────────────────────────────

test("the shape check fails against the original builders", () => {
  // Without this, a shape function that returned the same thing for every
  // input would make every test below pass and prove nothing.
  const original = require("./original");
  const failures = [];
  const check = (name, render) => {
    try {
      assertSameShape(render);
    } catch {
      failures.push(name);
    }
  };
  check("log card", (p) =>
    original.logCardHtml({ hash: "abc123", parent: "00root", message: p, author: "Ada" }, true, "t")
  );
  check("branch row", (p) => original.branchRowHtml({ name: p, tip: "x" }, false));
  check("toast", (p) => original.toastHtml("error", p));
  assert.deepEqual(failures, ["log card", "branch row", "toast"]);
});

// ── The escapers ─────────────────────────────────────────────────────────────

test("escapeHtml covers the five characters that matter", () => {
  assert.equal(escapeHtml(`<&>"'`), "&lt;&amp;&gt;&quot;&#39;");
});

test("escapeHtml escapes the ampersand first, so nothing is double-escaped wrongly", () => {
  assert.equal(escapeHtml("&lt;"), "&amp;lt;");
});

test("escapeAttr matches escapeHtml", () => {
  for (const s of HOSTILE) assert.equal(escapeAttr(s), escapeHtml(s));
});

test("escaping leaves ordinary text alone", () => {
  assert.equal(escapeHtml("Draft 3 — final (really)"), "Draft 3 — final (really)");
});

// ── Shape ────────────────────────────────────────────────────────────────────

const commit = (over = {}) => ({
  hash: "abc123def456",
  parent: "00root",
  message: "Initial commit",
  author: "Ada",
  timestamp: "2026-01-01T09:00:00.000Z",
  ...over,
});

test("a commit message cannot change the shape of a log card", () => {
  assertSameShape((payload) => logCardHtml(commit({ message: payload }), true));
});

test("a commit author cannot change the shape of a log card", () => {
  assertSameShape((payload) => logCardHtml(commit({ author: payload }), true));
});

test("a commit hash cannot change the shape of a log card", () => {
  assertSameShape((payload) => logCardHtml(commit({ hash: payload }), true));
});

test("a parent hash cannot change the shape of a log card", () => {
  assertSameShape((payload) => logCardHtml(commit({ parent: payload }), false));
});

test("a timestamp cannot change the shape of a log card", () => {
  assertSameShape((payload) => logCardHtml(commit({ timestamp: payload }), true));
});

test("a branch name cannot change the shape of a branch row", () => {
  assertSameShape((payload) => branchRowHtml(payload, false));
  assertSameShape((payload) => branchRowHtml(payload, true));
});

test("an error message cannot change the shape of a toast", () => {
  assertSameShape((payload) => toastHtml("error", payload));
});

// ── What the markup still has to contain ─────────────────────────────────────

test("a log card carries the buttons the task pane wires up", () => {
  const html = logCardHtml(commit(), false);
  for (const attr of ["data-ask-restore", "data-cancel", "data-restore", "data-diff", "data-parent"]) {
    assert.ok(html.includes(attr), `missing ${attr}`);
  }
  assert.ok(html.includes(`class="log-summary"`));
  assert.ok(html.includes(`class="card-confirm"`));
});

test("the root commit has no diff button, because it has nothing to diff against", () => {
  assert.ok(!logCardHtml(commit(), true).includes("data-diff"));
});

test("the active branch shows a badge and no actions", () => {
  const active = branchRowHtml("main", true);
  assert.ok(active.includes("branch-badge"));
  assert.ok(!active.includes("data-switch"));
  assert.ok(!active.includes("data-merge"));
});

test("an inactive branch offers switch and merge", () => {
  const row = branchRowHtml("draft", false);
  assert.ok(row.includes(`data-switch="draft"`));
  assert.ok(row.includes(`data-merge="draft"`));
});

test("a branch name with a quote is still readable back off the attribute", () => {
  // The task pane reads the value off the element rather than building a
  // selector from it, so this has to survive escaping and come back intact.
  const name = `say "hi" & <goodbye>`;
  const row = branchRowHtml(name, false);
  const value = /data-switch="([^"]*)"/.exec(row)[1];
  const decoded = value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&");
  assert.equal(decoded, name);
});

test("an empty commit message shows a placeholder", () => {
  assert.ok(logCardHtml(commit({ message: "" }), true).includes("(no message)"));
});

test("an empty author shows a placeholder", () => {
  assert.ok(logCardHtml(commit({ author: "" }), true).includes("unknown"));
});

test("a toast carries its icon and its dismiss button", () => {
  const html = toastHtml("warning", "careful");
  assert.ok(html.includes("⚠️"));
  assert.ok(html.includes(`class="toast-close"`));
  assert.ok(html.includes("careful"));
});

test("an unparseable timestamp is shown as it stands rather than as Invalid Date", () => {
  assert.equal(formatTimestamp("not a date"), "not a date");
  assert.ok(!formatTimestamp("2026-01-01T09:00:00.000Z").includes("Invalid"));
});
