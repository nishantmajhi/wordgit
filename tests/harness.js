"use strict";

// Shared scaffolding for the WordGit suite.
//
// `npm test` compiles src/ to tests/build/ as CommonJS first (see the pretest
// script and tests/tsconfig.json), so everything here loads the real add-in
// source rather than a copy of it.
//
// The two fakes below are the whole of what the VCS needs from Word. They are
// what makes a full init-commit-branch-merge sequence runnable in Node: the
// original reached for Office.js from inside the commit logic, so none of it
// could be exercised without the host.

const assert = require("node:assert/strict");
const path = require("node:path");
const { parse } = require("./xml");

const BUILD = path.join(__dirname, "build");

function load(modulePath) {
  return require(path.join(BUILD, modulePath));
}

const RepoFormat = load("vcs/RepoFormat.js");
const ObjectStoreModule = load("vcs/ObjectStore.js");
const HistoryModule = load("vcs/History.js");
const { DiffEngine } = load("vcs/DiffEngine.js");
const { CommitManager } = load("vcs/CommitManager.js");
const { BranchManager } = load("vcs/BranchManager.js");

/**
 * An in-memory stand-in for the document's Custom XML Parts.
 *
 * `supportsSetXml: false` reproduces a host with no in-place replacement, so
 * the add-before-remove path can be exercised too.
 */
class FakeParts {
  constructor(options = {}) {
    this.parts = new Map();
    this.nextId = 1;
    this.supportsSetXml = options.supportsSetXml !== false;
    this.log = [];
    this.failNextAdd = false;
  }

  static withXml(xml, options) {
    const parts = new FakeParts(options);
    parts.seed(xml);
    return parts;
  }

  seed(xml) {
    const id = `part${this.nextId++}`;
    this.parts.set(id, xml);
    return id;
  }

  async list() {
    return Array.from(this.parts, ([id, xml]) => ({ id, xml }));
  }

  async setXml(id, xml) {
    if (!this.supportsSetXml) return false;
    this.log.push(`set:${id}`);
    this.parts.set(id, xml);
    return true;
  }

  async add(xml) {
    if (this.failNextAdd) {
      this.failNextAdd = false;
      throw new Error("host refused the new part");
    }
    const id = `part${this.nextId++}`;
    this.log.push(`add:${id}`);
    this.parts.set(id, xml);
    return id;
  }

  async remove(id) {
    this.log.push(`remove:${id}`);
    this.parts.delete(id);
  }
}

/** An in-memory stand-in for the document body. */
class FakeDocument {
  constructor(ooxml = "<pkg>empty</pkg>") {
    this.ooxml = ooxml;
    this.writes = [];
    this.author = null;
  }

  async getOoxml() {
    return this.ooxml;
  }

  async replaceBody(ooxml, tracked) {
    this.ooxml = ooxml;
    this.writes.push({ ooxml, tracked });
  }

  async authorName() {
    return this.author;
  }
}

/** A store wired to the fakes, parsing with the XML shim rather than a browser. */
function makeStore(parts) {
  return new ObjectStoreModule.ObjectStore(parts, parse);
}

/** A complete VCS over in-memory fakes. */
function makeVcs(options = {}) {
  const parts = options.parts ?? new FakeParts(options);
  const doc = options.doc ?? new FakeDocument(options.ooxml);
  const store = makeStore(parts);
  const diff = new DiffEngine();
  return {
    parts,
    doc,
    store,
    diff,
    commits: new CommitManager(store, diff, doc),
    branches: new BranchManager(store),
  };
}

/**
 * A small OOXML-shaped document, long enough that diff-match-patch behaves as
 * it does on a real one rather than on a toy string.
 */
function makeOoxml(paragraphs) {
  const body = paragraphs
    .map(
      (text, i) =>
        `<w:p w:rsidR="00${i}A${i}B" w14:paraId="0000000${i}" w14:textId="7777000${i}">` +
        `<w:r><w:t>${text}</w:t></w:r></w:p>`
    )
    .join("");
  return (
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}</w:body></w:document>`
  );
}

/** Assert that `fn` throws an Error whose message contains `needle`. */
async function rejectsWith(fn, needle) {
  let caught = null;
  try {
    await fn();
  } catch (e) {
    caught = e;
  }
  assert.ok(caught, `expected a rejection mentioning ${JSON.stringify(needle)}, got none`);
  assert.ok(
    String(caught.message).includes(needle),
    `expected a rejection mentioning ${JSON.stringify(needle)}, got: ${caught.message}`
  );
  return caught;
}

module.exports = {
  BUILD,
  load,
  parse,
  RepoFormat,
  ObjectStoreModule,
  HistoryModule,
  DiffEngine,
  CommitManager,
  BranchManager,
  FakeParts,
  FakeDocument,
  makeStore,
  makeVcs,
  makeOoxml,
  rejectsWith,
};
