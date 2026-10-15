"use strict";

// Stripping WordGit's history out of a .docx.
//
// The fixtures are real zips built with the same JSZip the add-in uses, and
// every assertion reads the result back out of a regenerated zip rather than
// inspecting the JSZip object it was building. A package that is right in
// memory and wrong on disk is the failure worth catching.

const test = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");

const { load, RepoFormat } = require("./harness");
const { stripWordGitParts, isWordGitPart, cleanFileName } = load("vcs/DocxCleaner.js");

const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CUSTOM_XML_CT = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

function wordgitPart() {
  return RepoFormat.repoToXml({
    head: "refs/heads/main",
    author: "Ada",
    refs: [{ name: "main", tip: "aaa" }],
    objects: [
      {
        hash: "aaa",
        parent: "00root",
        parent2: "",
        author: "Ada",
        timestamp: "2026-01-01T09:00:00.000Z",
        message: "Initial commit",
        delta: "<pkg/>",
        encoding: "snapshot",
      },
    ],
  });
}

/**
 * A .docx with `items` custom XML parts, each described as
 * { n, xml, props?: boolean }. Mirrors what Word actually writes: an item, a
 * rels file pointing at its itemProps, and entries in [Content_Types].xml and
 * word/_rels/document.xml.rels.
 */
function makeDocx(items) {
  const zip = new JSZip();
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document/>`);

  const overrides = [
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`,
  ];
  const rels = [
    `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="word/document.xml"/>`,
  ];

  items.forEach((item, i) => {
    zip.file(`customXml/item${item.n}.xml`, item.xml);
    overrides.push(
      `<Override PartName="/customXml/item${item.n}.xml" ContentType="application/xml"/>`
    );
    rels.push(
      `<Relationship Id="rIdc${i}" Type="${REL_NS}/customXml" Target="../customXml/item${item.n}.xml"/>`
    );
    if (item.props !== false) {
      zip.file(
        `customXml/_rels/item${item.n}.xml.rels`,
        `<?xml version="1.0"?><Relationships xmlns="${REL_NS}">` +
          `<Relationship Id="rId1" Type="${REL_NS}/customXmlProps" Target="itemProps${item.n}.xml"/>` +
          `</Relationships>`
      );
      zip.file(`customXml/itemProps${item.n}.xml`, `<?xml version="1.0"?><ds:datastoreItem/>`);
      overrides.push(
        `<Override PartName="/customXml/itemProps${item.n}.xml" ContentType="${CUSTOM_XML_CT}"/>`
      );
    }
  });

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0"?><Types xmlns="${CT_NS}">${overrides.join("")}</Types>`
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0"?><Relationships xmlns="${REL_NS}">${rels.join("")}</Relationships>`
  );
  return zip;
}

/** Strip, regenerate the zip, and read the result back as a fresh package. */
async function cleanRoundTrip(zip) {
  const result = await stripWordGitParts(zip);
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { result, out: await JSZip.loadAsync(bytes) };
}

const paths = (zip) => Object.keys(zip.files).sort();
const textOf = (zip, path) => zip.file(path).async("text");

// ── Recognising our own parts ────────────────────────────────────────────────

test("a WordGit part is recognised by its root element", () => {
  assert.equal(isWordGitPart(wordgitPart()), true);
});

test("a part that merely mentions the namespace is not ours", () => {
  // A document about WordGit, with the namespace quoted in a property.
  const other =
    `<?xml version="1.0"?><properties><note>` +
    `See ${RepoFormat.NS} for the schema.</note></properties>`;
  assert.equal(isWordGitPart(other), false);
});

test("a part with our element name but another namespace is not ours", () => {
  assert.equal(isWordGitPart(`<?xml version="1.0"?><repo xmlns="http://example.invalid/"/>`), false);
});

test("an empty or non-XML part is not ours", () => {
  assert.equal(isWordGitPart(""), false);
  assert.equal(isWordGitPart("not xml at all"), false);
});

// ── Stripping ────────────────────────────────────────────────────────────────

test("a document with no WordGit data is left alone", async () => {
  const zip = makeDocx([{ n: 1, xml: `<?xml version="1.0"?><properties/>` }]);
  const before = paths(zip);
  const { result, out } = await cleanRoundTrip(zip);
  assert.deepEqual(result.removed, []);
  assert.deepEqual(paths(out), before);
});

test("the WordGit item, its rels and its itemProps all go", async () => {
  const zip = makeDocx([{ n: 3, xml: wordgitPart() }]);
  const { result, out } = await cleanRoundTrip(zip);
  assert.deepEqual(result.removed.sort(), [
    "customXml/_rels/item3.xml.rels",
    "customXml/item3.xml",
    "customXml/itemProps3.xml",
  ]);
  assert.deepEqual(paths(out), [
    "[Content_Types].xml",
    "customXml/",
    "customXml/_rels/",
    "word/",
    "word/_rels/",
    "word/_rels/document.xml.rels",
    "word/document.xml",
  ]);
});

test("the content-type overrides for the removed parts go too", async () => {
  const { out } = await cleanRoundTrip(makeDocx([{ n: 3, xml: wordgitPart() }]));
  const ct = await textOf(out, "[Content_Types].xml");
  assert.ok(!ct.includes("item3.xml"), ct);
  assert.ok(!ct.includes("itemProps3.xml"), ct);
  assert.ok(ct.includes("/word/document.xml"), "the document's own override is untouched");
});

test("the document relationship to the removed part goes too", async () => {
  const { out } = await cleanRoundTrip(makeDocx([{ n: 3, xml: wordgitPart() }]));
  const rels = await textOf(out, "word/_rels/document.xml.rels");
  assert.ok(!rels.includes("customXml/item3.xml"), rels);
  assert.ok(rels.includes("word/document.xml"), "the document relationship is untouched");
});

test("a user's own custom XML part is kept, with its bookkeeping", async () => {
  const zip = makeDocx([
    { n: 1, xml: `<?xml version="1.0"?><properties><title>Q3</title></properties>` },
    { n: 2, xml: wordgitPart() },
  ]);
  const { result, out } = await cleanRoundTrip(zip);

  assert.ok(result.removed.includes("customXml/item2.xml"));
  assert.ok(!result.removed.includes("customXml/item1.xml"));
  assert.ok(out.file("customXml/item1.xml"), "the user's part survived");
  assert.ok(out.file("customXml/itemProps1.xml"), "and so did its props");

  const ct = await textOf(out, "[Content_Types].xml");
  assert.ok(ct.includes("/customXml/item1.xml"));
  assert.ok(!ct.includes("/customXml/item2.xml"));

  const rels = await textOf(out, "word/_rels/document.xml.rels");
  assert.ok(rels.includes("customXml/item1.xml"));
  assert.ok(!rels.includes("customXml/item2.xml"));
});

test("two WordGit parts are both removed", async () => {
  const zip = makeDocx([
    { n: 1, xml: wordgitPart() },
    { n: 2, xml: wordgitPart() },
  ]);
  const { out } = await cleanRoundTrip(zip);
  assert.deepEqual(
    paths(out).filter((p) => p.startsWith("customXml/item")),
    []
  );
});

test("an item with no rels file is still removed", async () => {
  const zip = makeDocx([{ n: 1, xml: wordgitPart(), props: false }]);
  const { result, out } = await cleanRoundTrip(zip);
  assert.deepEqual(result.removed, ["customXml/item1.xml"]);
  assert.ok(!out.file("customXml/item1.xml"));
});

test("an Override written with a separate closing tag is still removed", async () => {
  // The original matched only the self-closing form.
  const zip = makeDocx([{ n: 1, xml: wordgitPart(), props: false }]);
  const ct = await textOf(zip, "[Content_Types].xml");
  zip.file(
    "[Content_Types].xml",
    ct.replace(
      `<Override PartName="/customXml/item1.xml" ContentType="application/xml"/>`,
      `<Override PartName="/customXml/item1.xml" ContentType="application/xml"></Override>`
    )
  );
  const { out } = await cleanRoundTrip(zip);
  assert.ok(!(await textOf(out, "[Content_Types].xml")).includes("/customXml/item1.xml"));
});

test("an Override for a part whose name merely contains ours is kept", async () => {
  // "/customXml/item1.xml" is a substring of "/customXml/item1.xml.bak", which
  // a pattern built around the value would have matched.
  const zip = makeDocx([{ n: 1, xml: wordgitPart(), props: false }]);
  const ct = await textOf(zip, "[Content_Types].xml");
  zip.file(
    "[Content_Types].xml",
    ct.replace(
      "</Types>",
      `<Override PartName="/customXml/item1.xml.bak" ContentType="application/xml"/></Types>`
    )
  );
  const { out } = await cleanRoundTrip(zip);
  const after = await textOf(out, "[Content_Types].xml");
  assert.ok(after.includes("/customXml/item1.xml.bak"), after);
  assert.ok(!after.includes(`PartName="/customXml/item1.xml"`), after);
});

test("a relationship target written without the ../ prefix is still matched", async () => {
  const zip = makeDocx([{ n: 1, xml: wordgitPart(), props: false }]);
  const rels = await textOf(zip, "word/_rels/document.xml.rels");
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(`Target="../customXml/item1.xml"`, `Target="/customXml/item1.xml"`)
  );
  const { out } = await cleanRoundTrip(zip);
  assert.ok(!(await textOf(out, "word/_rels/document.xml.rels")).includes("customXml/item1.xml"));
});

test("the cleaned package is still a readable zip with the document in it", async () => {
  const { out } = await cleanRoundTrip(makeDocx([{ n: 1, xml: wordgitPart() }]));
  assert.equal(await textOf(out, "word/document.xml"), `<?xml version="1.0"?><w:document/>`);
});

// ── Filename ─────────────────────────────────────────────────────────────────

test("a clean copy is named after the document", () => {
  assert.equal(cleanFileName("C:\\Users\\ada\\Report.docx"), "Report_clean.docx");
  assert.equal(cleanFileName("/home/ada/Report.docx"), "Report_clean.docx");
  assert.equal(cleanFileName("Report.docx"), "Report_clean.docx");
});

test("a query string or fragment does not end up in the filename", () => {
  assert.equal(
    cleanFileName("https://example.sharepoint.com/sites/x/Report.docx?web=1&d=abc"),
    "Report_clean.docx"
  );
  assert.equal(cleanFileName("https://example.invalid/Report.docx#page=2"), "Report_clean.docx");
});

test("a percent-encoded name is decoded", () => {
  assert.equal(cleanFileName("https://example.invalid/Q3%20Report.docx"), "Q3 Report_clean.docx");
});

test("a name that is not valid percent-encoding is used as it stands", () => {
  assert.equal(cleanFileName("https://example.invalid/100%25.docx"), "100%_clean.docx");
  assert.equal(cleanFileName("/tmp/50%.docx"), "50%_clean.docx");
});

test("a missing document URL gives a usable default", () => {
  assert.equal(cleanFileName(null), "document_clean.docx");
  assert.equal(cleanFileName(undefined), "document_clean.docx");
  assert.equal(cleanFileName(""), "document_clean.docx");
});

test("a name with no extension keeps its whole name", () => {
  assert.equal(cleanFileName("/tmp/Untitled"), "Untitled_clean.docx");
});
