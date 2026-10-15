"use strict";

// The commit graph: what it places where, and the SVG it turns into.

const test = require("node:test");
const assert = require("node:assert/strict");

const { load } = require("./harness");
const { buildGraph } = load("vcs/GraphBuilder.js");
const { renderGraphSvg, pillWidth, graphTimestamp, truncate } = load("taskpane/graphSvg.js");

function c(hash, parent, over = {}) {
  return {
    hash,
    parent,
    parent2: "",
    message: `message ${hash}`,
    timestamp: "2026-01-01T09:00:00.000Z",
    author: "Ada",
    ...over,
  };
}

const rows = (layout) => layout.nodes.map((n) => n.hash);
const rowOf = (layout, hash) => layout.nodes.find((n) => n.hash === hash).row;

test("an empty history lays out as nothing", () => {
  const layout = buildGraph([], []);
  assert.deepEqual(layout, { nodes: [], edges: [], cols: 1, rows: 0 });
});

test("a linear history is newest first", () => {
  const commits = [
    c("a", "00root", { timestamp: "2026-01-01T00:00:00Z" }),
    c("b", "a", { timestamp: "2026-01-02T00:00:00Z" }),
    c("d", "b", { timestamp: "2026-01-03T00:00:00Z" }),
  ];
  const layout = buildGraph(commits, [{ name: "main", tip: "d" }]);
  assert.deepEqual(rows(layout), ["d", "b", "a"]);
  assert.equal(layout.cols, 1);
  assert.equal(layout.rows, 3);
});

test("a parent is never placed above its own child", () => {
  // Two branches of different lengths, with a clock that disagrees.
  const commits = [
    c("a", "00root", { timestamp: "2026-05-01T00:00:00Z" }),
    c("b", "a", { timestamp: "2020-01-01T00:00:00Z" }),
    c("x", "a", { timestamp: "2026-06-01T00:00:00Z" }),
    c("y", "x", { timestamp: "2026-06-02T00:00:00Z" }),
    c("z", "y", { timestamp: "2026-06-03T00:00:00Z" }),
  ];
  const layout = buildGraph(commits, [
    { name: "main", tip: "b" },
    { name: "long", tip: "z" },
  ]);
  for (const node of layout.nodes) {
    const commit = commits.find((k) => k.hash === node.hash);
    if (commit.parent === "00root") continue;
    assert.ok(
      rowOf(layout, commit.parent) > node.row,
      `${commit.parent} was placed above its child ${node.hash}`
    );
  }
});

test("each branch tip gets its own lane", () => {
  const commits = [c("a", "00root"), c("b", "a"), c("x", "a")];
  const layout = buildGraph(commits, [
    { name: "main", tip: "b" },
    { name: "draft", tip: "x" },
  ]);
  assert.equal(layout.cols, 2);
  const lanes = new Map(layout.nodes.map((n) => [n.hash, n.col]));
  assert.notEqual(lanes.get("b"), lanes.get("x"));
});

test("a merge commit gets an edge to each parent", () => {
  const commits = [
    c("a", "00root", { timestamp: "2026-01-01T00:00:00Z" }),
    c("b", "a", { timestamp: "2026-01-02T00:00:00Z" }),
    c("x", "a", { timestamp: "2026-01-03T00:00:00Z" }),
    c("m", "b", { parent2: "x", timestamp: "2026-01-04T00:00:00Z" }),
  ];
  const layout = buildGraph(commits, [{ name: "main", tip: "m" }]);
  const from = layout.edges.filter((e) => e.fromHash === "m").map((e) => e.toHash);
  assert.deepEqual(from.sort(), ["b", "x"]);
  assert.equal(layout.nodes.find((n) => n.hash === "m").isMerge, true);
  assert.equal(layout.nodes.find((n) => n.hash === "b").isMerge, false);
});

test("commits not reachable from any tip are left out", () => {
  const commits = [c("a", "00root"), c("b", "a"), c("orphan", "00root")];
  const layout = buildGraph(commits, [{ name: "main", tip: "b" }]);
  assert.deepEqual(rows(layout).sort(), ["a", "b"]);
});

test("a ref pointing at a commit that is not stored is ignored", () => {
  const layout = buildGraph([c("a", "00root")], [
    { name: "main", tip: "a" },
    { name: "ghost", tip: "missing" },
  ]);
  assert.deepEqual(rows(layout), ["a"]);
  assert.deepEqual(layout.nodes[0].refs, ["main"]);
});

test("two branches at the same commit both label it", () => {
  const layout = buildGraph([c("a", "00root")], [
    { name: "main", tip: "a" },
    { name: "release", tip: "a" },
  ]);
  assert.deepEqual(layout.nodes[0].refs, ["main", "release"]);
  assert.equal(layout.cols, 1, "the second ref does not claim a second lane");
});

test("the root commit is marked as such", () => {
  const layout = buildGraph([c("a", "00root"), c("b", "a")], [{ name: "main", tip: "b" }]);
  assert.equal(layout.nodes.find((n) => n.hash === "a").isRoot, true);
  assert.equal(layout.nodes.find((n) => n.hash === "b").isRoot, false);
});

// ── SVG ──────────────────────────────────────────────────────────────────────

function svgFor(commits, refs, dark = false) {
  return renderGraphSvg(buildGraph(commits, refs), { dark });
}

test("the SVG has one circle per commit and one path or line per edge", () => {
  const commits = [c("a", "00root"), c("b", "a"), c("x", "a"), c("m", "b", { parent2: "x" })];
  const refs = [{ name: "main", tip: "m" }];
  const svg = svgFor(commits, refs);
  assert.equal((svg.match(/<circle /g) || []).length, 4);
  assert.equal((svg.match(/<line |<path /g) || []).length, 4);
});

test("the SVG is well-formed and has a single root", () => {
  const svg = svgFor([c("a", "00root")], [{ name: "main", tip: "a" }]);
  const { parse } = require("./harness");
  const doc = parse(svg);
  assert.equal(doc.documentElement.localName, "svg");
});

test("a branch name and a message are escaped into the SVG", () => {
  const svg = svgFor([c("a", "00root", { message: `<script>alert(1)</script>` })], [
    { name: `"><script>x</script>`, tip: "a" },
  ]);
  const { parse } = require("./harness");
  assert.equal(parse(svg).documentElement.localName, "svg", "still parses as one SVG");
  assert.ok(!svg.includes("<script>"), svg);
});

test("the dark and light palettes differ only in their colours", () => {
  const commits = [c("a", "00root")];
  const refs = [{ name: "main", tip: "a" }];
  const light = svgFor(commits, refs, false);
  const dark = svgFor(commits, refs, true);
  const withoutColours = (s) =>
    s.replace(/#[0-9a-f]{6}/gi, "#").replace(/rgba\([^)]*\)/g, "rgba()");
  assert.notEqual(light, dark);
  assert.equal(withoutColours(light), withoutColours(dark), "only the colour literals change");
});

test("lane width grows with the longest branch name", () => {
  const commits = [c("a", "00root"), c("b", "a"), c("x", "a")];
  const short = svgFor(commits, [
    { name: "m", tip: "b" },
    { name: "d", tip: "x" },
  ]);
  const long = svgFor(commits, [
    { name: "a-very-long-branch-name-indeed", tip: "b" },
    { name: "d", tip: "x" },
  ]);
  const widthOf = (s) => Number(/viewBox="0 0 (\d+(?:\.\d+)?)/.exec(s)[1]);
  assert.ok(widthOf(long) > widthOf(short));
});

test("pill width never drops below the minimum", () => {
  assert.equal(pillWidth(""), 20);
  assert.ok(pillWidth("main") >= 20);
  assert.ok(pillWidth("a-long-branch-name") > pillWidth("main"));
});

test("an unparseable timestamp renders as nothing rather than Invalid Date", () => {
  assert.equal(graphTimestamp("not a date"), "");
  assert.equal(graphTimestamp(""), "");
  assert.match(graphTimestamp("2026-03-04T05:06:07Z"), /^2026-03-04 \d\d:\d\d$/);
});

test("truncate keeps short strings and adds an ellipsis to long ones", () => {
  assert.equal(truncate("short", 22), "short");
  assert.equal(truncate("x".repeat(30), 10), "x".repeat(9) + "…");
  assert.equal(truncate("x".repeat(10), 10), "x".repeat(10));
});
