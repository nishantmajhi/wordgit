"use strict";

// The storage format: what goes into a document and what comes back out.
//
// This is the contract that cannot be changed carelessly, because the data
// lives inside users' .docx files. Every test here is about one of two
// properties: a repository written by this code reads back identical, and a
// repository this code cannot understand is refused rather than misread.

const test = require("node:test");
const assert = require("node:assert/strict");

const { parse, RepoFormat } = require("./harness");
const { parseRepoXml, repoToXml, RepoFormatError, emptyRepo, NS } = RepoFormat;

function commit(over = {}) {
  return {
    hash: "aaa111",
    parent: "00root",
    parent2: "",
    author: "Ada",
    timestamp: "2026-01-01T09:00:00.000Z",
    message: "Initial commit",
    delta: "<pkg/>",
    encoding: "snapshot",
    ...over,
  };
}

function repo(over = {}) {
  return {
    head: "refs/heads/main",
    author: "Ada",
    refs: [{ name: "main", tip: "aaa111" }],
    objects: [commit()],
    ...over,
  };
}

const roundTrip = (r) => parseRepoXml(repoToXml(r), parse);

// ── Round trips ──────────────────────────────────────────────────────────────

test("an empty repository round-trips", () => {
  assert.deepEqual(roundTrip(emptyRepo()), emptyRepo());
});

test("a repository with commits and branches round-trips", () => {
  const r = repo({
    refs: [
      { name: "main", tip: "bbb222" },
      { name: "draft-2", tip: "aaa111" },
    ],
    objects: [commit(), commit({ hash: "bbb222", parent: "aaa111", encoding: "dmp-lz" })],
  });
  assert.deepEqual(roundTrip(r), r);
});

test("a merge commit's second parent round-trips", () => {
  const r = repo({
    objects: [commit(), commit({ hash: "ccc333", parent: "aaa111", parent2: "bbb222" })],
    refs: [{ name: "main", tip: "ccc333" }],
  });
  assert.equal(roundTrip(r).objects[1].parent2, "bbb222");
});

test("a commit with no second parent does not gain an empty attribute", () => {
  assert.ok(!repoToXml(repo()).includes("parent2"));
});

// ── Escaping ─────────────────────────────────────────────────────────────────
//
// The review asked specifically for this one: a message containing <, & and a
// ]]> sequence, because the escaping is load-bearing.

const NASTY = `a < b & c ]]> d "quoted" 'single' > e`;

test("a hostile commit message round-trips exactly", () => {
  const r = repo({ objects: [commit({ message: NASTY })] });
  assert.equal(roundTrip(r).objects[0].message, NASTY);
});

test("a hostile payload round-trips exactly", () => {
  const r = repo({ objects: [commit({ delta: NASTY })] });
  assert.equal(roundTrip(r).objects[0].delta, NASTY);
});

test("a hostile branch name round-trips exactly", () => {
  const r = repo({ refs: [{ name: `"><evil attr='x'>`, tip: "aaa111" }] });
  assert.equal(roundTrip(r).refs[0].name, `"><evil attr='x'>`);
});

test("a ]]> sequence in the payload does not end the element early", () => {
  const xml = repoToXml(repo({ objects: [commit({ delta: "before ]]> after" })] }));
  assert.equal(parseRepoXml(xml, parse).objects[0].delta, "before ]]> after");
});

test("a repository written with CDATA payloads is still readable", () => {
  // What the original wrote. Reading it has to keep working, because the data
  // is inside documents that already exist.
  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<wordgit:repo xmlns:wordgit="${NS}">
  <HEAD>refs/heads/main</HEAD>
  <config><author>Ada</author></config>
  <refs>
    <branch name="main" tip="aaa111"/>
  </refs>
  <objects>
    <commit hash="aaa111" parent="00root">
      <author>Ada</author>
      <timestamp>2026-01-01T09:00:00.000Z</timestamp>
      <message>Initial commit</message>
      <delta encoding="snapshot"><![CDATA[<pkg>a ]]]]><![CDATA[> b</pkg>]]></delta>
    </commit>
  </objects>
</wordgit:repo>`;
  const read = parseRepoXml(xml, parse);
  assert.equal(read.objects[0].delta, "<pkg>a ]]> b</pkg>");
  assert.equal(read.objects[0].parent2, "", "an old commit has no second parent");
});

test("every ASCII character except the ones XML forbids round-trips in a payload", () => {
  const allowed = [];
  for (let i = 0; i < 128; i++) {
    if (i < 0x20 && i !== 0x09 && i !== 0x0a && i !== 0x0d) continue;
    allowed.push(String.fromCharCode(i));
  }
  const payload = allowed.join("");
  assert.equal(roundTrip(repo({ objects: [commit({ delta: payload })] })).objects[0].delta, payload);
});

// ── Refusing what cannot be understood ───────────────────────────────────────

/** Run `fn`, require it to throw a RepoFormatError, and return the error. */
function throwsFormatError(fn) {
  let caught = null;
  try {
    fn();
  } catch (e) {
    caught = e;
  }
  assert.ok(caught, "expected a RepoFormatError, nothing was thrown");
  assert.ok(caught instanceof RepoFormatError, `expected a RepoFormatError, got ${caught}`);
  return caught;
}

const rejects = (xml, needle) => {
  const e = throwsFormatError(() => parseRepoXml(xml, parse));
  if (needle) assert.ok(e.message.includes(needle), e.message);
  return e;
};

test("truncated XML is refused", () => {
  const xml = repoToXml(repo());
  rejects(xml.slice(0, xml.length - 40), "not a repository");
});

test("XML that parses but is not a repository is refused", () => {
  rejects(`<?xml version="1.0"?><properties><title>Q3</title></properties>`, "not a repository");
});

test("a commit with no hash is refused", () => {
  const xml = repoToXml(repo()).replace(`hash="aaa111"`, `hash=""`);
  rejects(xml, "no hash");
});

test("a commit with no parent pointer is refused", () => {
  const xml = repoToXml(repo()).replace(` parent="00root"`, "");
  rejects(xml, "no parent pointer");
});

test("a commit with no payload is refused", () => {
  const xml = repoToXml(repo()).replace(/<delta[\s\S]*?<\/delta>/, "");
  rejects(xml, "no payload");
});

test("a payload encoding this version does not know is refused by name", () => {
  const xml = repoToXml(repo()).replace(`encoding="snapshot"`, `encoding="zstd-2"`);
  rejects(xml, `"zstd-2"`);
});

test("two commits with the same hash are refused", () => {
  const r = repo({ objects: [commit(), commit()] });
  // repoToXml will happily write it; reading it back is where it is caught.
  rejects(repoToXml(r), "stored twice");
});

test("two branches with the same name are refused", () => {
  const r = repo({
    refs: [
      { name: "main", tip: "aaa111" },
      { name: "main", tip: "aaa111" },
    ],
  });
  rejects(repoToXml(r), "listed twice");
});

test("a branch with no name is refused", () => {
  const xml = repoToXml(repo()).replace(`name="main"`, `name=""`);
  rejects(xml, "no name");
});

test("a branch with no tip is allowed", () => {
  const r = repo({ refs: [{ name: "main", tip: "" }] });
  assert.equal(parseRepoXml(repoToXml(r), parse).refs[0].tip, "");
});

test("a missing HEAD falls back to main rather than failing", () => {
  const xml = repoToXml(repo()).replace(/<HEAD>[\s\S]*?<\/HEAD>/, "");
  assert.equal(parseRepoXml(xml, parse).head, "refs/heads/main");
});

// ── Refusing to write what cannot be read back ───────────────────────────────

test("a payload containing a character XML cannot store is refused at write time", () => {
  const bell = String.fromCharCode(7);
  const e = throwsFormatError(() =>
    repoToXml(repo({ objects: [commit({ delta: `a${bell}b` })] }))
  );
  assert.ok(e.message.includes("U+0007"), e.message);
  assert.ok(e.message.includes("aaa111"), "the message names the commit");
});

test("a commit message containing a NUL is refused at write time", () => {
  const nul = String.fromCharCode(0);
  const e = throwsFormatError(() =>
    repoToXml(repo({ objects: [commit({ message: `a${nul}b` })] }))
  );
  assert.ok(e.message.includes("U+0000"), e.message);
});

test("tab, newline and carriage return are all storable", () => {
  const payload = "a\tb\nc\rd";
  assert.equal(roundTrip(repo({ objects: [commit({ delta: payload })] })).objects[0].delta, payload);
});
