"use strict";

// Handing a file to the user.
//
// Two details decide whether the download happens at all, and both were wrong.
// Neither is visible by reading the code, which is why they get a test.

const test = require("node:test");
const assert = require("node:assert/strict");

const { load } = require("./harness");

/**
 * A DOM small enough to watch what saveBlob does, and nothing more.
 *
 * The globals have to stay in place until `run` has finished, deferred work
 * included -- saveBlob's cleanup happens on a later turn of the event loop,
 * which is the whole point of it.
 */
async function withFakeDom(run) {
  const saved = {
    document: globalThis.document,
    URL: globalThis.URL,
    Blob: globalThis.Blob,
  };
  const events = [];
  const anchor = {
    href: "",
    download: "",
    style: {},
    click() {
      events.push({ type: "click", attached: body.children.includes(anchor), href: anchor.href });
    },
    remove() {
      const i = body.children.indexOf(anchor);
      if (i >= 0) body.children.splice(i, 1);
      events.push({ type: "remove" });
    },
  };
  const body = {
    children: [],
    appendChild(el) {
      body.children.push(el);
      events.push({ type: "append" });
    },
  };

  globalThis.document = {
    body,
    createElement(tag) {
      assert.equal(tag, "a");
      return anchor;
    },
  };
  globalThis.URL = {
    createObjectURL() {
      events.push({ type: "create" });
      return "blob:fake-url";
    },
    revokeObjectURL(url) {
      events.push({ type: "revoke", url });
    },
  };

  try {
    return await run({ events, anchor, body });
  } finally {
    globalThis.document = saved.document;
    globalThis.URL = saved.URL;
    globalThis.Blob = saved.Blob;
  }
}

const { saveBlob } = load("taskpane/download.js");

const tick = () => new Promise((resolve) => setTimeout(resolve, 0));

test("the anchor is in the document before it is clicked", async () => {
  await withFakeDom(async ({ events }) => {
    saveBlob(new Blob(["x"]), "file.txt");
    const click = events.find((e) => e.type === "click");
    assert.ok(click, "nothing was clicked");
    assert.equal(click.attached, true, "a detached anchor does nothing in some engines");
    await tick();
  });
});

test("the object URL is still live when the click happens", async () => {
  await withFakeDom(async ({ events }) => {
    saveBlob(new Blob(["x"]), "file.txt");
    const order = events.map((e) => e.type);
    assert.deepEqual(order, ["create", "append", "click"]);
    assert.ok(!order.includes("revoke"), "revoking before the download starts loses the file");
    await tick();
  });
});

test("the object URL and the anchor are both cleaned up afterwards", async () => {
  await withFakeDom(async ({ events, body }) => {
    saveBlob(new Blob(["x"]), "file.txt");
    await tick();
    const revoke = events.find((e) => e.type === "revoke");
    assert.ok(revoke, "the object URL was never revoked");
    assert.equal(revoke.url, "blob:fake-url");
    assert.deepEqual(body.children, [], "the anchor was left in the document");
  });
});

test("the filename is put on the anchor", async () => {
  await withFakeDom(async ({ anchor }) => {
    saveBlob(new Blob(["x"]), "Report_clean.docx");
    assert.equal(anchor.download, "Report_clean.docx");
    assert.equal(anchor.href, "blob:fake-url");
    await tick();
  });
});
