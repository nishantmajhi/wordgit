"use strict";

// The manifest is what Word shows the user before they trust the add-in, and
// the scaffolding in it said the add-in was published by Contoso. These are
// the checks that stop that coming back.

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const { parse } = require("./harness");

const ROOT = path.join(__dirname, "..");
const manifest = fs.readFileSync(path.join(ROOT, "manifest.xml"), "utf8");
const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8"));
const doc = parse(manifest);

const text = (selector) => doc.querySelector(selector)?.textContent ?? null;
const attr = (selector, name) => doc.querySelector(selector)?.getAttribute(name) ?? null;

test("the manifest is well-formed XML with an OfficeApp root", () => {
  assert.equal(doc.documentElement.localName, "OfficeApp");
});

test("no scaffolding placeholder is left anywhere in the manifest", () => {
  for (const needle of [
    "Contoso",
    "contoso.com",
    "ButtonId1",
    "A template to get started",
    "sample add-in",
    "Show Task Pane",
    "Commands Group",
  ]) {
    assert.ok(!manifest.includes(needle), `manifest still contains "${needle}"`);
  }
});

test("the add-in names its own publisher", () => {
  const provider = text("ProviderName");
  assert.ok(provider && provider.trim(), "ProviderName is empty");
  assert.notEqual(provider, "Contoso");
});

test("the description says what the add-in does", () => {
  const description = attr("Description", "DefaultValue");
  assert.ok(description && description.length > 30, `description is ${description}`);
  assert.match(description, /version control/i);
});

test("no unrelated domain is declared as a trusted app domain", () => {
  // An AppDomain widens what the add-in's webview may navigate to. The only
  // location this add-in loads is its own SourceLocation, which is trusted
  // implicitly, so there should be no AppDomains element at all.
  assert.equal(doc.querySelector("AppDomains"), null);
});

test("every URL in the manifest is https and points at localhost or the project", () => {
  const urls = [...manifest.matchAll(/DefaultValue="(https?:[^"]*)"/g)].map((m) => m[1]);
  assert.ok(urls.length > 0);
  for (const url of urls) {
    assert.ok(url.startsWith("https://"), `${url} is not https`);
    assert.ok(
      url.startsWith("https://localhost:3000/") ||
        url.startsWith("https://github.com/nishantmajhi/wordgit"),
      `${url} points somewhere unexpected`
    );
  }
});

test("the taskpane id is not the generator's placeholder", () => {
  assert.equal(text("TaskpaneId"), "WordGitTaskpane");
});

test("the add-in is registered for Word only", () => {
  // The VersionOverrides block has its own <Hosts><Host> with an xsi:type
  // rather than a Name, so only the named ones are the host registrations.
  const hosts = doc
    .querySelectorAll("Hosts > Host")
    .map((el) => el.getAttribute("Name"))
    .filter((name) => name !== null);
  assert.deepEqual(hosts, ["Document"]);
});

test("it asks for the permission it needs and no more", () => {
  assert.equal(text("Permissions"), "ReadWriteDocument");
});

test("the manifest id is a well-formed GUID", () => {
  assert.match(text("Id"), /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i);
});

// ── package.json ─────────────────────────────────────────────────────────────

test("package.json is this project, not the template it was generated from", () => {
  assert.equal(pkg.name, "wordgit");
  assert.ok(!JSON.stringify(pkg).includes("Office-Addin-TaskPane"));
  assert.ok(!JSON.stringify(pkg).includes("office-addin-taskpane"));
});

test("package.json has a description and a repository", () => {
  assert.ok(pkg.description && pkg.description.length > 20);
  assert.match(pkg.repository.url, /nishantmajhi\/wordgit/);
});

test("npm test compiles the source before running the suite", () => {
  assert.match(pkg.scripts.pretest, /tsc -p tests\/tsconfig\.json/);
  assert.match(pkg.scripts.test, /node --test/);
});

test("a LICENSE is present and matches the licence package.json declares", () => {
  const licence = fs.readFileSync(path.join(ROOT, "LICENSE"), "utf8");
  assert.equal(pkg.license, "SEE LICENSE IN LICENSE");
  assert.match(licence, /^Personal Use License$/m);
  assert.match(licence, /^Copyright \(c\) [\d-]+ Manish Majhi\./m, "the licence names the copyright holder");
});

test("the manifest's provider is the copyright holder", () => {
  assert.equal(text("ProviderName"), "Manish Majhi");
});

// ── The pages the manifest points at ─────────────────────────────────────────

test("the source locations the manifest names exist in the source tree", () => {
  for (const [, url] of manifest.matchAll(/DefaultValue="https:\/\/localhost:3000\/([^"]*)"/g)) {
    const candidates = [
      path.join(ROOT, "src", "taskpane", url),
      path.join(ROOT, "src", "commands", url),
      path.join(ROOT, url),
    ];
    assert.ok(
      candidates.some((p) => fs.existsSync(p)),
      `${url} is served from nowhere in the source tree`
    );
  }
});
