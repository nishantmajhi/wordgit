/**
 * render.ts - the markup the task pane builds at runtime.
 *
 * Everything here is a pure string function, so the one thing that matters
 * about this markup can be tested: that content taken out of a document can
 * never change its shape.
 *
 * It matters because a WordGit repository travels inside the .docx. Commit
 * messages, branch names and the author line are written by whoever edited the
 * document before you, and the task pane is a webview with the document open
 * through Office.js. The original interpolated all three straight into
 * `innerHTML`, so opening a file someone else had committed to was enough to
 * run their script against your document.
 */

export type ToastType = "success" | "error" | "warning" | "info";

export const TOAST_ICONS: Record<ToastType, string> = {
  success: "✅",
  error: "❌",
  warning: "⚠️",
  info: "ℹ️",
};

/** Escape for HTML element content. */
export function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Escape for a double-quoted HTML attribute value.
 *
 * The same set as `escapeHtml`. It is a separate function because the two are
 * separate contexts and only one of them was ever the reason the quote is in
 * the list; keeping them apart is what stops someone trimming the quote out of
 * a function they think is only used for element content.
 */
export function escapeAttr(s: string): string {
  return escapeHtml(s);
}

export interface LogCardCommit {
  hash: string;
  parent: string;
  message: string;
  author: string;
  timestamp: string;
}

/** Format a commit timestamp for display, or return it unchanged if unparseable. */
export function formatTimestamp(iso: string): string {
  const d = new Date(iso);
  return isNaN(d.getTime()) ? iso : d.toLocaleString();
}

export function toastHtml(type: ToastType, msg: string): string {
  return `
    <span class="toast-icon">${TOAST_ICONS[type]}</span>
    <span class="toast-msg">${escapeHtml(msg)}</span>
    <button class="toast-close" aria-label="Dismiss">×</button>`;
}

export function logCardHtml(c: LogCardCommit, isRoot: boolean): string {
  const hash = escapeAttr(c.hash);
  return `
        <div class="log-summary">
          <span class="log-chevron">▶</span>
          <span class="log-hash">${escapeHtml(c.hash.slice(0, 6))}</span>
          <span class="log-msg">${escapeHtml(c.message || "(no message)")}</span>
        </div>
        <div class="log-detail">
          <div class="log-detail-meta">
            <strong>${escapeHtml(c.author || "unknown")}</strong> · ${escapeHtml(
              formatTimestamp(c.timestamp)
            )}<br/>
            <code>${escapeHtml(c.hash)}</code>
          </div>
          <div class="log-actions">
            <button class="btn btn-primary" data-ask-restore="${hash}">↩ Restore</button>
            ${
              isRoot
                ? ""
                : `<button class="btn btn-danger" data-diff="${hash}" data-parent="${escapeAttr(
                    c.parent
                  )}">⟷ Show diff</button>`
            }
          </div>
          <div class="card-confirm" id="confirm-${hash}">
            <div class="card-confirm-msg">Restore to this commit?</div>
            <div class="card-confirm-btns">
              <button class="btn btn-primary" data-restore="${hash}">↩ Yes, restore</button>
              <button class="btn btn-secondary" data-cancel="${hash}">✕ Cancel</button>
            </div>
          </div>
        </div>`;
}

export function branchRowHtml(name: string, isActive: boolean): string {
  const attr = escapeAttr(name);
  return `
          <span class="branch-name">${escapeHtml(name)}</span>
          ${
            isActive
              ? `<span class="branch-badge">current</span>`
              : `<button class="btn btn-primary" data-switch="${attr}">⇄ Switch</button>
               <button class="btn btn-danger"  data-merge="${attr}">⤢ Merge in</button>`
          }`;
}
