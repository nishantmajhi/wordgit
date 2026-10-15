"use strict";

// The original code, transcribed.
//
// Three of this project's bugs are in code that cannot be run as it stands:
// the serialiser's helpers are module-private, the task pane's markup builders
// are inline template literals inside handlers that call Office.js, and the
// merge sits behind a manager that wanted a Word request context. So the parts
// that *decided* something are transcribed here with their surroundings lifted
// out, exactly as the original wrote them, and the assertions in
// original.test.js run against both these and the fixed code.
//
// A transcription that has quietly drifted from what it claims to reproduce
// asserts nothing, so original.test.js reads the baseline commit out of git and
// requires that every line quoted in BASELINE_LINES is still in it. If the
// original source changes, or this file is edited to make a test pass, that
// check fails first.
//
// Source: commit 0800c86 "Initial commit".

const { DOMParser } = require("./xml");

const NS = "http://schemas.wordgit.local/vcs/1";

// ── src/vcs/ObjectStore.ts ───────────────────────────────────────────────────

function esc(s) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeCdata(s) {
  return `<![CDATA[${s.replace(/\]\]>/g, "]]]]><![CDATA[>")}]]>`;
}

function repoToXml(repo) {
  const refs = repo.refs
    .map((r) => `<branch name="${esc(r.name)}" tip="${esc(r.tip)}"/>`)
    .join("\n    ");
  const objects = repo.objects
    .map(
      (c) => `<commit hash="${esc(c.hash)}" parent="${esc(c.parent)}">
      <author>${esc(c.author)}</author>
      <timestamp>${esc(c.timestamp)}</timestamp>
      <message>${esc(c.message)}</message>
      <delta encoding="${esc(c.encoding)}">${escapeCdata(c.delta)}</delta>
    </commit>`
    )
    .join("\n    ");
  return `<?xml version="1.0" encoding="UTF-8"?>
<wordgit:repo xmlns:wordgit="${NS}">
  <HEAD>${esc(repo.head)}</HEAD>
  <config><author>${esc(repo.author)}</author></config>
  <refs>
    ${refs}
  </refs>
  <objects>
    ${objects}
  </objects>
</wordgit:repo>`;
}

function xmlToRepo(xml) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");

  const head = doc.querySelector("HEAD")?.textContent ?? "refs/heads/main";
  const author = doc.querySelector("config > author")?.textContent ?? "Author";

  const refs = [];
  doc.querySelectorAll("refs > branch").forEach((el) => {
    refs.push({
      name: el.getAttribute("name") ?? "",
      tip: el.getAttribute("tip") ?? "",
    });
  });

  const objects = [];
  doc.querySelectorAll("objects > commit").forEach((el) => {
    objects.push({
      hash: el.getAttribute("hash") ?? "",
      parent: el.getAttribute("parent") ?? "",
      author: el.querySelector("author")?.textContent ?? "",
      timestamp: el.querySelector("timestamp")?.textContent ?? "",
      message: el.querySelector("message")?.textContent ?? "",
      delta: el.querySelector("delta")?.textContent ?? "",
      encoding: el.querySelector("delta")?.getAttribute("encoding") ?? "dmp",
    });
  });

  return { head, author, refs, objects };
}

function sha6() {
  return (Date.now().toString(36) + Math.random().toString(36).slice(2, 6))
    .slice(0, 12)
    .padEnd(12, "0");
}

// ── src/vcs/DiffEngine.ts ────────────────────────────────────────────────────

const VOLATILE_ATTR_RE = /\s+w:(?:rsid\w*|paraId|textId|id)="[^"]*"/g;

function normaliseOoxml(ooxml) {
  return ooxml.replace(VOLATILE_ATTR_RE, "");
}

// ── src/vcs/MergeManager.ts ──────────────────────────────────────────────────

function threeWayMerge(diff, baseRaw, oursRaw, theirsRaw) {
  const base = normaliseOoxml(baseRaw);
  const ours = normaliseOoxml(oursRaw);
  const theirs = normaliseOoxml(theirsRaw);

  if (ours === base) return theirs;
  if (theirs === base) return ours;

  const oursPatch = diff.makePatch(base, ours);
  const theirsPatch = diff.makePatch(base, theirs);

  const { result: afterOurs, applied: oursApplied } = diff.applyPatchWithResults(base, oursPatch);

  const oursFailures = oursApplied.filter((ok) => !ok).length;
  if (oursFailures > 0) {
    throw new Error(
      `CONFLICT: ours patch failed to apply (${oursFailures} hunk(s)). Manual resolution required.`
    );
  }

  const { result: merged, applied: theirsApplied } = diff.applyPatchWithResults(
    afterOurs,
    theirsPatch
  );

  const theirsFailures = theirsApplied.filter((ok) => !ok).length;
  if (theirsFailures > 0) {
    throw new Error(
      `CONFLICT: theirs patch failed to apply (${theirsFailures} hunk(s)). Manual resolution required.`
    );
  }

  return merged;
}

// ── src/taskpane/taskpane.ts ─────────────────────────────────────────────────

const TOAST_ICONS = {
  success: "✅",
  error: "❌",
  warning: "⚠️",
  info: "ℹ️",
};

function toastHtml(type, msg) {
  return `
    <span class="toast-icon">${TOAST_ICONS[type]}</span>
    <span class="toast-msg">${msg}</span>
    <button class="toast-close" aria-label="Dismiss">×</button>`;
}

function logCardHtml(c, isRoot, ts) {
  return `
        <div class="log-summary">
          <span class="log-chevron">▶</span>
          <span class="log-hash">${c.hash.slice(0, 6)}</span>
          <span class="log-msg">${c.message || "(no message)"}</span>
        </div>
        <div class="log-detail">
          <div class="log-detail-meta">
            <strong>${c.author || "unknown"}</strong> · ${ts}<br/>
            <code>${c.hash}</code>
          </div>
          <div class="log-actions">
            <button class="btn btn-primary" data-ask-restore="${c.hash}">↩ Restore</button>
            ${isRoot ? "" : `<button class="btn btn-danger" data-diff="${c.hash}" data-parent="${c.parent}">⟷ Show diff</button>`}
          </div>
          <div class="card-confirm" id="confirm-${c.hash}">
            <div class="card-confirm-msg">Restore to this commit?</div>
            <div class="card-confirm-btns">
              <button class="btn btn-primary" data-restore="${c.hash}">↩ Yes, restore</button>
              <button class="btn btn-secondary" data-cancel="${c.hash}">✕ Cancel</button>
            </div>
          </div>
        </div>`;
}

function branchRowHtml(branch, isActive) {
  return `
          <span class="branch-name">${branch.name}</span>
          ${
            isActive
              ? `<span class="branch-badge">current</span>`
              : `<button class="btn btn-primary" data-switch="${branch.name}">⇄ Switch</button>
               <button class="btn btn-danger"  data-merge="${branch.name}">⤢ Merge in</button>`
          }`;
}

function gcEsc(s) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function graphNodeHtml(n, tx, y, textFill, mutedFill) {
  return (
    `<text x="${tx}" y="${y - 9}" ` +
    `font-family="Cascadia Code,Consolas,monospace" font-size="9" fill="${mutedFill}">` +
    `${n.hash.slice(0, 6)}</text>` +
    `<text x="${tx}" y="${y + 2}" ` +
    `font-family="Segoe UI,system-ui,sans-serif" font-size="10" ` +
    `fill="${textFill}" font-weight="400">` +
    `${gcEsc(n.message)}</text>`
  );
}

/**
 * Lines that must still be present, verbatim, in the baseline commit. If a
 * transcription above is edited without the original changing, one of these
 * stops matching and original.test.js fails.
 */
const BASELINE_LINES = {
  "src/vcs/ObjectStore.ts": [
    `      <delta encoding="\${esc(c.encoding)}">\${escapeCdata(c.delta)}</delta>`,
    `  const head = doc.querySelector("HEAD")?.textContent ?? "refs/heads/main";`,
    `  return \`<![CDATA[\${s.replace(/\\]\\]>/g, "]]]]><![CDATA[>")}]]>\`;`,
    `    .replace(/"/g, "&quot;");`,
  ],
  "src/vcs/CommitManager.ts": [
    `  return (Date.now().toString(36) + Math.random().toString(36).slice(2, 6))`,
    `    .padEnd(12, "0");`,
  ],
  "src/vcs/DiffEngine.ts": [
    `const VOLATILE_ATTR_RE = /\\s+w:(?:rsid\\w*|paraId|textId|id)="[^"]*"/g;`,
  ],
  "src/vcs/MergeManager.ts": [
    `    const { result: afterOurs, applied: oursApplied } =`,
    `      this.diff.applyPatchWithResults(base, oursPatch);`,
    `    return merged;`,
  ],
  "src/taskpane/taskpane.ts": [
    `          <span class="log-msg">\${c.message || "(no message)"}</span>`,
    `    <span class="toast-msg">\${msg}</span>`,
    `          <span class="branch-name">\${branch.name}</span>`,
    `              : \`<button class="btn btn-primary" data-switch="\${branch.name}">⇄ Switch</button>`,
    `            \`\${n.hash.slice(0, 6)}</text>\``,
  ],
  "src/vcs/GraphBuilder.ts": [
    `export interface GraphCommit {`,
    `  hash: string;`,
    `  parent: string;`,
    `  message: string;`,
  ],
};

module.exports = {
  BASELINE_COMMIT: "0800c86",
  BASELINE_LINES,
  NS,
  esc,
  escapeCdata,
  repoToXml,
  xmlToRepo,
  sha6,
  normaliseOoxml,
  threeWayMerge,
  toastHtml,
  logCardHtml,
  branchRowHtml,
  graphNodeHtml,
};
