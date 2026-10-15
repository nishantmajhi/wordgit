/**
 * download.ts - handing a file to the user from inside the task pane.
 *
 * Small, but its own file because both places that download something got the
 * same two things wrong, and the fix is easy to undo by accident.
 */

/**
 * Save `blob` as `filename`.
 *
 * Two details matter and neither is obvious:
 *
 * The anchor is put in the document before it is clicked. A click on a
 * detached anchor starts a download in Chromium and does nothing at all in
 * Firefox, and an Office task pane is a webview whose engine is not something
 * this code gets to choose.
 *
 * The object URL is revoked on the next turn of the event loop, not on the
 * line after `click()`. `click()` only *starts* the download; revoking the URL
 * synchronously can invalidate it before the fetch behind the download has
 * read it, which fails silently and looks like a button that does nothing.
 */
export function saveBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    a.remove();
    URL.revokeObjectURL(url);
  }, 0);
}
