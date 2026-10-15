/* global Word, Office */

import JSZip from "jszip";
import { ObjectStore, WordCustomXmlParts } from "../vcs/ObjectStore";
import { RepoFormatError, ROOT_PARENT } from "../vcs/RepoFormat";
import { BranchManager } from "../vcs/BranchManager";
import { CommitManager, WordDocumentIO } from "../vcs/CommitManager";
import { DiffEngine } from "../vcs/DiffEngine";
import { buildGraph, GraphCommit } from "../vcs/GraphBuilder";
import { cleanFileName, stripWordGitParts } from "../vcs/DocxCleaner";
import { branchRowHtml, escapeHtml, logCardHtml, toastHtml, ToastType } from "./render";
import { renderGraphSvg } from "./graphSvg";
import { saveBlob } from "./download";

// ─── Wiring ───────────────────────────────────────────────────────────────────
// Every handler builds its own managers against a fresh Word request context.
// A context belongs to one Word.run, so they cannot be cached across handlers.

const io = new WordDocumentIO();

async function withVcs<T>(
  fn: (vcs: {
    store: ObjectStore;
    branches: BranchManager;
    commits: CommitManager;
    diff: DiffEngine;
  }) => Promise<T>
): Promise<T> {
  return Word.run(async (context) => {
    const store = new ObjectStore(new WordCustomXmlParts(context));
    const diff = new DiffEngine();
    return fn({
      store,
      branches: new BranchManager(store),
      commits: new CommitManager(store, diff, io),
      diff,
    });
  });
}

/**
 * Report a failure to the user.
 *
 * Damaged history is called out separately and for longer. It is the one
 * failure where what the user does next decides whether their work survives,
 * and the add-in will not write to the document until it is resolved.
 */
function reportError(prefix: string, e: unknown): void {
  const message = e instanceof Error ? e.message : String(e);
  if (e instanceof RepoFormatError) {
    showToast(message, "error", 12000);
    return;
  }
  showToast(`${prefix}: ${message}`, "error");
}

// ─── Init state ───────────────────────────────────────────────────────────────
// On load, read the XML store. If branch refs exist, the repo was previously
// initialised - enable commit immediately. Otherwise keep disabled until init.

function setRepoReady(ready: boolean): void {
  const input = document.getElementById("commit-msg") as HTMLInputElement;
  const commit = document.getElementById("btn-commit") as HTMLButtonElement;
  const init = document.getElementById("btn-init") as HTMLButtonElement;
  input.disabled = !ready;
  commit.disabled = !ready;
  init.disabled = ready;
}

/** Lock every action that writes, after the stored history turned out unreadable. */
function setRepoDamaged(): void {
  const input = document.getElementById("commit-msg") as HTMLInputElement;
  const commit = document.getElementById("btn-commit") as HTMLButtonElement;
  const init = document.getElementById("btn-init") as HTMLButtonElement;
  input.disabled = true;
  commit.disabled = true;
  init.disabled = true;
}

async function checkRepoExists(): Promise<void> {
  try {
    const refs = await withVcs(async ({ store }) => (await store.load()).refs.length);
    setRepoReady(refs > 0);
  } catch (e) {
    // A document whose WordGit data cannot be read is not a document without
    // one. Offering Init here is what destroyed the history in the original:
    // Init replaces the Custom XML Part, and the unreadable part was the only
    // copy of it.
    if (e instanceof RepoFormatError) {
      setRepoDamaged();
      showToast(e.message, "error", 12000);
      return;
    }
    setRepoReady(false);
  }
}

// ─── Toast ────────────────────────────────────────────────────────────────────

function showToast(msg: string, type: ToastType = "info", durationMs = 4000): void {
  const container = document.getElementById("toast-container")!;
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.innerHTML = toastHtml(type, msg);
  container.appendChild(el);
  const dismiss = () => {
    el.classList.add("toast-out");
    el.addEventListener("animationend", () => el.remove());
  };
  el.querySelector(".toast-close")!.addEventListener("click", dismiss);
  setTimeout(dismiss, durationMs);
}

// ─── Theme ────────────────────────────────────────────────────────────────────
// Detects dark mode by measuring luminance of the Office background colour.
// Toggles body.wg-dark and overrides CSS vars accordingly.

function isDark(hex: string): boolean {
  const c = hex.replace("#", "");
  const r = parseInt(c.substring(0, 2), 16) / 255;
  const g = parseInt(c.substring(2, 4), 16) / 255;
  const b = parseInt(c.substring(4, 6), 16) / 255;
  return 0.299 * r + 0.587 * g + 0.114 * b < 0.35;
}

function applyTheme(theme?: Office.OfficeTheme): void {
  if (!theme) return;
  const r = document.documentElement.style;
  const dark = isDark(theme.bodyBackgroundColor ?? "#f5f5f5");
  document.body.classList.toggle("wg-dark", dark);
  if (dark) {
    r.setProperty("--wg-bg", theme.bodyBackgroundColor ?? "#1e1e1e");
    r.setProperty("--wg-surface", theme.bodyBackgroundColor ?? "#252525");
    r.setProperty("--wg-surface-alt", "#2d2d2d");
    r.setProperty("--wg-surface-btn", "#2d2d2d");
    r.setProperty("--wg-border", "#3a3a3a");
    r.setProperty("--wg-border-input", "#4a4a4a");
    r.setProperty("--wg-text", theme.bodyForegroundColor ?? "#ececec");
    r.setProperty("--wg-text-muted", "#909090");
    r.setProperty("--wg-text-faint", "#606060");
    r.setProperty("--wg-accent-glow", "rgba(0,120,212,0.20)");
    r.setProperty("--wg-accent-text", "#4db6ff");
  } else {
    r.removeProperty("--wg-bg");
    r.removeProperty("--wg-surface");
    r.removeProperty("--wg-surface-alt");
    r.removeProperty("--wg-surface-btn");
    r.removeProperty("--wg-border");
    r.removeProperty("--wg-border-input");
    r.removeProperty("--wg-text");
    r.removeProperty("--wg-text-muted");
    r.removeProperty("--wg-text-faint");
    r.removeProperty("--wg-accent-glow");
    r.removeProperty("--wg-accent-text");
  }
}

// ─── Tabs ─────────────────────────────────────────────────────────────────────

function initTabs(): void {
  document.querySelectorAll<HTMLButtonElement>(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => switchToTab(btn.dataset.tab ?? ""));
  });
}

/** Programmatically switch to a tab by name (e.g. "log", "branches"). */
function switchToTab(tab: string): void {
  document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
  document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
  document.querySelector(`.tab-btn[data-tab="${tab}"]`)?.classList.add("active");
  document.getElementById(`panel-${tab}`)?.classList.add("active");
}

// ─── Init repo ────────────────────────────────────────────────────────────────

async function handleInit(): Promise<void> {
  try {
    await withVcs(({ commits }) => commits.init());
    setRepoReady(true);
    showToast("Repo initialised.", "success");
  } catch (e) {
    if (e instanceof RepoFormatError) setRepoDamaged();
    reportError("Init failed", e);
  }
}

// ─── Commit ───────────────────────────────────────────────────────────────────

async function handleCommit(): Promise<void> {
  const msgEl = document.getElementById("commit-msg") as HTMLInputElement;
  const message = msgEl.value.trim() || "Update";
  try {
    const hash = await withVcs(({ commits }) => commits.commit(message));
    msgEl.value = "";
    showToast(`Committed ${hash.slice(0, 6)} - "${message}"`, "success");
  } catch (e) {
    if (e instanceof RepoFormatError) setRepoDamaged();
    reportError("Commit failed", e);
  }
}

// ─── Commit with timestamp (Alt+C) ────────────────────────────────────────────
// Initialises the repo first if not yet initialised, then commits with the
// current timestamp as the commit message. Ignores any text in the commit
// message input - this is a quick "checkpoint" shortcut.

function timestampMessage(): string {
  const d = new Date();
  const p = (n: number) => n.toString().padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(
    d.getMinutes()
  )}:${p(d.getSeconds())}`;
}

async function handleCommitWithTimestamp(): Promise<void> {
  const message = timestampMessage();
  try {
    const hash = await withVcs(async ({ commits }) => {
      await commits.init(); // no-op when the repo already exists
      return commits.commit(message);
    });
    setRepoReady(true);
    showToast(`Committed ${hash.slice(0, 6)} - "${message}"`, "success");
  } catch (e) {
    if (e instanceof RepoFormatError) setRepoDamaged();
    reportError("Commit failed", e);
  }
}

async function handleExportJson(): Promise<void> {
  try {
    const commits = await withVcs(({ commits: mgr }) => mgr.log());
    const json = JSON.stringify(commits, null, 2);
    const date = new Date().toISOString().slice(0, 10);
    saveBlob(new Blob([json], { type: "application/json" }), `wordgit-history-${date}.json`);
    showToast("History exported.", "success");
  } catch (e) {
    reportError("Export failed", e);
  }
}

// ─── Download clean copy (Alt+S) ──────────────────────────────────────────────
// Fetches the raw .docx as bytes, strips all WordGit Custom XML Parts (and
// their relationship/content-type entries), and downloads the result as
// "<original-name>_clean.docx". See DocxCleaner.ts for what is removed.

async function getDocumentBytes(): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4 * 1024 * 1024 },
      (fileResult) => {
        if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(fileResult.error?.message || "getFileAsync failed"));
          return;
        }
        const file = fileResult.value;
        const sliceCount = file.sliceCount;
        const slices: Uint8Array[] = new Array(sliceCount);
        let received = 0;
        // One slice failing has to stop the others reporting, or the file is
        // closed twice and the promise settles more than once.
        let done = false;

        const finish = (fn: () => void) => {
          if (done) return;
          done = true;
          file.closeAsync(() => {
            /* ignore close result */
          });
          fn();
        };

        if (sliceCount === 0) {
          finish(() => resolve(new Uint8Array(0)));
          return;
        }

        for (let i = 0; i < sliceCount; i++) {
          file.getSliceAsync(i, (sliceResult) => {
            if (done) return;
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              finish(() => reject(new Error(sliceResult.error?.message || "getSliceAsync failed")));
              return;
            }
            slices[sliceResult.value.index] = new Uint8Array(sliceResult.value.data);
            received++;
            if (received !== sliceCount) return;
            finish(() => {
              const total = slices.reduce((sum, s) => sum + s.length, 0);
              const merged = new Uint8Array(total);
              let offset = 0;
              for (const s of slices) {
                merged.set(s, offset);
                offset += s.length;
              }
              resolve(merged);
            });
          });
        }
      }
    );
  });
}

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

async function handleDownloadClean(): Promise<void> {
  try {
    showToast("Preparing clean copy…", "info", 2000);

    const bytes = await getDocumentBytes();
    const filename = cleanFileName(Office.context.document.url);
    const zip = await JSZip.loadAsync(bytes);
    const { removed } = await stripWordGitParts(zip);

    if (removed.length === 0) {
      saveBlob(new Blob([bytes as BlobPart], { type: DOCX_MIME }), filename);
      showToast("No WordGit data found - nothing to clean.", "warning");
      return;
    }

    const outBytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    saveBlob(new Blob([outBytes as BlobPart], { type: DOCX_MIME }), filename);
    showToast("Clean copy downloaded.", "success");
  } catch (e) {
    reportError("Clean export failed", e);
  }
}

// ─── Diff vs parent ───────────────────────────────────────────────────────────

async function handleDiffVsParent(hash: string, parentHash: string): Promise<void> {
  try {
    await withVcs(async ({ commits }) => {
      const fromOoxml = await commits.reconstruct(parentHash);
      const toOoxml = await commits.reconstruct(hash);
      await io.replaceBody(fromOoxml, false);
      await io.replaceBody(toOoxml, true);
    });
    showToast("Diff loaded. Accept all tracked changes when done reviewing.", "success");
  } catch (e) {
    reportError("Error showing diff", e);
  }
}

// ─── Restore ─────────────────────────────────────────────────────────────────

async function handleRestore(hash: string): Promise<void> {
  try {
    await withVcs(({ commits }) => commits.checkout(hash));
    showToast(
      `Restored to ${hash.slice(0, 6)}. Review tracked changes and accept when ready.`,
      "success"
    );
  } catch (e) {
    reportError("Restore failed", e);
  }
}

// ─── Log ─────────────────────────────────────────────────────────────────────

/**
 * Wire up the buttons inside one rendered element.
 *
 * The handlers are found by attribute *name* and the value is read off the
 * element. Building a selector out of the value, which is what the original
 * did, means a commit or branch whose name contains a quote throws a
 * SyntaxError out of querySelector and leaves the whole list inert.
 */
function onEach(
  root: ParentNode,
  attribute: string,
  handler: (value: string, el: HTMLElement) => void
): void {
  root.querySelectorAll<HTMLElement>(`[${attribute}]`).forEach((el) => {
    const value = el.getAttribute(attribute) ?? "";
    el.addEventListener("click", () => handler(value, el));
  });
}

async function handleRefreshLog(): Promise<void> {
  const list = document.getElementById("log-list")!;
  list.innerHTML = "";
  try {
    const commits = await withVcs(({ commits: mgr }) => mgr.log());

    if (commits.length === 0) {
      list.innerHTML = `<li style="color:var(--wg-text-muted);padding:4px 0;font-size:12px;">No commits yet.</li>`;
      return;
    }

    commits.forEach((c) => {
      const li = document.createElement("li");
      li.className = "log-card";
      li.innerHTML = logCardHtml(c, c.parent === ROOT_PARENT);
      list.appendChild(li);

      li.querySelector(".log-summary")!.addEventListener("click", () => {
        const isExpanded = li.classList.contains("expanded");
        document
          .querySelectorAll(".log-card.expanded")
          .forEach((card) => card.classList.remove("expanded"));
        if (!isExpanded) li.classList.add("expanded");
      });

      const confirmBox = li.querySelector(".card-confirm")!;
      onEach(li, "data-ask-restore", () => confirmBox.classList.add("visible"));
      onEach(li, "data-cancel", () => confirmBox.classList.remove("visible"));
      onEach(li, "data-restore", (hash) => handleRestore(hash));
      onEach(li, "data-diff", (hash, el) =>
        handleDiffVsParent(hash, el.getAttribute("data-parent") ?? "")
      );
    });
  } catch (e) {
    reportError("Error loading log", e);
  }
}

// ─── Branch graph ─────────────────────────────────────────────────────────────

async function renderBranchGraph(): Promise<void> {
  const container = document.getElementById("branch-graph")!;
  container.classList.remove("has-graph");
  container.innerHTML = "";

  try {
    const svg = await withVcs(async ({ store }) => {
      const repo = await store.load();
      if (repo.refs.length === 0 || repo.objects.length === 0) return "";

      const commits: GraphCommit[] = repo.objects.map((c) => ({
        hash: c.hash,
        parent: c.parent,
        parent2: c.parent2,
        message: c.message,
        timestamp: c.timestamp,
        author: c.author,
      }));
      const layout = buildGraph(commits, repo.refs);
      if (layout.nodes.length === 0) return "";
      return renderGraphSvg(layout, { dark: document.body.classList.contains("wg-dark") });
    });

    if (!svg) {
      container.innerHTML = `<span class="wg-hint">No commits yet.</span>`;
      return;
    }
    container.classList.add("has-graph");
    container.innerHTML = svg;
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    container.innerHTML = `<span class="wg-hint" style="color:var(--wg-danger)">Graph error: ${escapeHtml(
      message
    )}</span>`;
  }
}

// ─── Branches ────────────────────────────────────────────────────────────────

async function handleRefreshBranches(): Promise<void> {
  const container = document.getElementById("branch-list")!;
  container.innerHTML = "";
  try {
    const { current, refs } = await withVcs(async ({ branches }) => ({
      current: await branches.currentBranch(),
      refs: await branches.listBranches(),
    }));

    refs.forEach((branch) => {
      const isActive = branch.name === current;
      const row = document.createElement("div");
      row.className = `branch-row${isActive ? " active" : ""}`;
      row.innerHTML = branchRowHtml(branch.name, isActive);
      container.appendChild(row);
      onEach(row, "data-switch", (name) => handleSwitchBranch(name));
      onEach(row, "data-merge", (name) => handleMerge(name));
    });

    await renderBranchGraph();
  } catch (e) {
    reportError("Error loading branches", e);
  }
}

async function handleSwitchBranch(name: string): Promise<void> {
  try {
    await withVcs(async ({ branches, commits }) => {
      const tip = await branches.switchBranch(name);
      await commits.checkoutSilent(tip);
    });
    showToast(`Switched to branch "${name}".`, "success");
    await handleRefreshBranches();
  } catch (e) {
    reportError("Switch failed", e);
  }
}

async function handleCreateBranch(): Promise<void> {
  const nameEl = document.getElementById("new-branch-name") as HTMLInputElement;
  const name = nameEl.value.trim();
  if (!name) {
    showToast("Enter a branch name.", "warning");
    return;
  }
  try {
    await withVcs(({ branches }) => branches.createBranch(name));
    nameEl.value = "";
    showToast(`Branch "${name}" created.`, "success");
    await handleRefreshBranches();
  } catch (e) {
    reportError("Create branch failed", e);
  }
}

async function handleMerge(theirBranch: string): Promise<void> {
  try {
    await withVcs(({ commits }) => commits.merge(theirBranch));
    showToast(`Merged "${theirBranch}". Review tracked changes and accept when ready.`, "success");
    await handleRefreshBranches();
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    if (message.includes("CONFLICT")) {
      showToast(`Merge conflict: ${message}`, "warning", 8000);
      return;
    }
    reportError("Merge failed", e);
  }
}

// ─── Office.onReady ───────────────────────────────────────────────────────────

Office.onReady(async () => {
  applyTheme(Office.context.officeTheme);

  // The enum member is OfficeThemeChanged, with a capital O. The original
  // asked for the camel-cased spelling, which is undefined, and silenced the
  // resulting type error with a cast -- so no handler was ever registered and
  // the task pane kept its startup colours when Office switched between light
  // and dark.
  if (Office.EventType?.OfficeThemeChanged) {
    Office.context.document.addHandlerAsync(
      Office.EventType.OfficeThemeChanged,
      (args: { officeTheme?: Office.OfficeTheme }) => applyTheme(args.officeTheme)
    );
  }

  initTabs();

  document.getElementById("btn-init")!.addEventListener("click", handleInit);
  document.getElementById("btn-commit")!.addEventListener("click", handleCommit);
  document.getElementById("commit-msg")!.addEventListener("keydown", (e) => {
    if ((e as KeyboardEvent).key === "Enter") handleCommit();
  });

  document.getElementById("btn-refresh-log")!.addEventListener("click", handleRefreshLog);
  document.getElementById("btn-export-json")!.addEventListener("click", handleExportJson);

  document.getElementById("btn-refresh-branches")!.addEventListener("click", handleRefreshBranches);
  document.getElementById("btn-create-branch")!.addEventListener("click", handleCreateBranch);
  document.getElementById("new-branch-name")!.addEventListener("keydown", (e) => {
    if ((e as KeyboardEvent).key === "Enter") handleCreateBranch();
  });

  // ── Global keyboard shortcuts ────────────────────────────────────────────────
  //   Alt+C - commit with timestamp message (init repo first if needed)
  //   Alt+L - switch to Log tab + refresh
  //   Alt+B - switch to Branches tab + refresh
  //   Alt+S - download clean copy (.docx with WordGit data stripped)
  document.addEventListener("keydown", (e: KeyboardEvent) => {
    if (!e.altKey || e.ctrlKey || e.metaKey) return;
    switch (e.key.toLowerCase()) {
      case "c":
        e.preventDefault();
        handleCommitWithTimestamp();
        break;
      case "l":
        e.preventDefault();
        switchToTab("log");
        handleRefreshLog();
        break;
      case "b":
        e.preventDefault();
        switchToTab("branches");
        handleRefreshBranches();
        break;
      case "s":
        e.preventDefault();
        handleDownloadClean();
        break;
    }
  });

  await checkRepoExists();
});
