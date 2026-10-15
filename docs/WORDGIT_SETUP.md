# WordGit - reproduction guide

> **Superseded in places, 2026-09-19.** These are the design notes from the
> sessions that built WordGit, kept because the architecture, the UI decisions
> and the "why blame was dropped" reasoning are still the best account of why
> the project is shaped the way it is. They are not the current contract. The
> storage format has changed since -- merge commits now record a second parent,
> payloads are escaped text rather than CDATA, and unreadable data is refused
> rather than treated as absent -- so read `README.md` and `src/vcs/RepoFormat.ts`
> for what the code does now, and this for why.

> **Purpose**: Reproduce the entire dev environment from scratch. Follow every step in order.
>
> **Instructions for Claude**: After completing any phase or significant step:
>
> 1. Update `WORDGIT_SETUP.md` - add new pitfalls, update file structure, update source sections
> 2. Update `WORDGIT_PROJECT.md` - mark checklist items, update next session prompt
> 3. Present both files for download before ending the session

---

## System requirements

| Thing   | Details                                                           |
| ------- | ----------------------------------------------------------------- |
| OS      | Windows (any modern version)                                      |
| Word    | Microsoft 365 desktop - not one-time license, not Word on the web |
| Node.js | Any LTS - https://nodejs.org                                      |
| npm     | Comes with Node                                                   |
| VS Code | https://code.visualstudio.com                                     |

---

## Phase 0 - install global tools

```powershell
npm install -g yo generator-office
yo --version   # verify
```

**Pitfalls**

- `yo` not found after install → close and reopen PowerShell
- `EACCES` error → run PowerShell as Administrator

---

## Phase 1 - scaffold

```powershell
yo office
```

| Prompt        | Answer                          |
| ------------- | ------------------------------- |
| Project type  | Office Add-in Task Pane project |
| Script type   | TypeScript                      |
| Add-in name   | wordgit                         |
| Office client | Word                            |

```powershell
cd wordgit
code .
npm start
```

**Pitfalls**

- Do not choose the React variant
- Do not rename the folder after generation
- No ribbon button → close Word, `Ctrl+C`, `npm start` again
- SSL warning → `npx office-addin-dev-certs install`, then retry

---

## Phase 2 - install dependencies

```powershell
npm install diff-match-patch jszip lz-string
npm install --save-dev @types/diff-match-patch
```

Note: `jszip` and `lz-string` ship their own TypeScript types - no `@types/jszip` or `@types/lz-string` packages needed.

---

## Phase 3 - file structure

```
wordgit/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html   ← UI shell; button colours live in <style> block here
│   │   ├── taskpane.css    ← layout, components, tokens - no button colours
│   │   └── taskpane.ts     ← all event handlers + Office.onReady
│   └── vcs/
│       ├── ObjectStore.ts      ← Custom XML read/write
│       ├── DiffEngine.ts       ← diff-match-patch wrapper + normaliseOoxml()
│       ├── BranchManager.ts    ← branch CRUD + HEAD
│       ├── CommitManager.ts    ← commit/log/checkout/checkoutSilent/merge/reconstruct + LZ compression
│       ├── MergeManager.ts     ← pure 3-way merge, no Word API calls
│       └── GraphBuilder.ts     ← DAG layout engine for branch graph SVG (Phase 7)
├── manifest.xml
├── package.json
├── tsconfig.json
└── webpack.config.js
```

---

## CRITICAL pitfalls - read before touching any file

### Office JS globals

- **Never** `import * as Office from "@microsoft/office-js"` → causes `Office is not defined`
- **Always** add `/* global Word, Office */` at top of any TS file using Office globals

### office.js stylesheet injection - THE BUTTON COLOUR BUG

`office.js` injects its own Fabric/Fluent stylesheet into the page **after** all `<link>` stylesheets
have loaded. This resets button `background-color` and `color` to browser defaults, causing a
"flash of correct colour then revert to black/transparent" bug.

**Fix**: Button colours must live in a `<style>` block inside `<head>` in `taskpane.html`.
A `<style>` block is part of the document cascade and always wins over injected external stylesheets.
`taskpane.css` handles everything else (layout, spacing, components, tokens).
**Never move button colour rules into taskpane.css** - they will be overridden by office.js.

### Other pitfalls

- `window.confirm` is NOT supported in Office sandbox - use inline HTML confirm panels
- `insertOoxml` must only be called ONCE per operation - two calls duplicates every paragraph
- Import diff-match-patch as `import * as dmp from "diff-match-patch"` - named imports fail
- DevTools: task pane runs in an iframe - switch console context from "top" to the taskpane frame
- Webpack rebuild: look for `webpack compiled successfully` in terminal after saving
- **Branch graph (Phase 7)**: `renderBranchGraph()` is called at the end of `handleRefreshBranches()`. Layout model: `[pill LEFT of circle] [circle] [text RIGHT of circle]` - everything centred on the same Y. Lane widths are **dynamic** - computed per-lane from `max(pillWidth(branchName)) + PILL_CIRCLE_GAP + R + LANE_SAFETY_PAD`. Do NOT revert to a fixed `COL_W` constant. Dark mode detected at render time via `body.classList.contains("wg-dark")`.
- **LZ compression (Phase 7)**:
  - `ObjectStore.CommitXml.encoding` now has 4 values: `"snapshot"`, `"dmp"` (legacy, uncompressed) and `"snapshot-lz"`, `"dmp-lz"` (compressed via `LZString.compressToBase64`).
  - `CommitManager.encodeDelta()` compresses on write; `decodeDelta()` decompresses on read. All new commits (init, commit, merge) write `-lz` variants.
  - **Backward compatible**: old `"snapshot"`/`"dmp"` commits from before this change still `reconstruct()` correctly - `decodeDelta()` passes them through unchanged.
  - Base64 output never contains `]]>`, so the existing `escapeCdata()` in `ObjectStore` needs no changes.
  - `LZString.decompressFromBase64()` returns `null` or `""` on corrupt input - `decodeDelta()` throws a clear "repo may be corrupt" error in that case rather than silently producing garbage OOXML.
- **Per-paragraph blame - NOT IMPLEMENTED (decision, not oversight)**:
  - `w:rsid*` attributes are Word _session_ IDs, regenerated by Word on every `insertOoxml(..., InsertLocation.replace)`. WordGit calls this on every checkout, branch switch, and merge - so a commit-time rsid → commit mapping goes stale the next time the user switches branches, with no way to detect the staleness.
  - If implemented later: do NOT key blame on `w:rsid`. Instead diff each paragraph's text content between consecutive commits (via `CommitManager.reconstruct()` for each commit in `CommitManager.log()`), keyed on `w:paraId` with a positional fallback, and report "last changed in commit X by author Y" per paragraph. Build as a new `BlameEngine.ts`; render as a 4th tab or an overlay on the Log tab.
  - `Alt+C` - `handleCommitWithTimestamp()`: checks if repo is initialised (no refs → calls `mgr.init()` first, then `setRepoReady(true)`), then commits with `YYYY-MM-DD HH:MM:SS` as the message. Ignores the commit-message input box entirely.
  - `Alt+L` / `Alt+B` - `switchToTab()` + refresh. `switchToTab()` is a programmatic version of the tab click handler (doesn't fire a click event, just toggles `.active` classes directly).
  - `Alt+S` - `handleDownloadClean()`: pulls the raw .docx via `Office.context.document.getFileAsync(Office.FileType.Compressed, ...)`, assembles slices into one `Uint8Array`, opens it with JSZip, finds `customXml/itemN.xml` files containing the `http://schemas.wordgit.local/vcs/1` namespace string, removes those files + their `_rels` + their `[Content_Types].xml` overrides + their `word/_rels/document.xml.rels` relationship entries, re-zips, and downloads as `<original-basename>_clean.docx`.
  - All shortcuts check `e.altKey && !e.ctrlKey && !e.metaKey` and call `e.preventDefault()` to avoid triggering Word's own Alt-menu navigation inside the task pane.

---

## UI/UX system - DO NOT change without updating this doc

### Colour tokens (CSS custom properties in taskpane.css `:root`)

```
--wg-bg:           #f5f5f5   app shell background
--wg-surface:      #ffffff   cards, panels, inputs
--wg-surface-alt:  #fafafa   expanded card body
--wg-surface-btn:  #eeeeee   secondary button fill
--wg-border:       #e0e0e0   default borders
--wg-border-input: #c8c8c8   input borders (slightly darker)
--wg-text:         #1a1a1a   primary text
--wg-text-muted:   #6b6b6b   labels, meta text
--wg-text-faint:   #a0a0a0   placeholders, hints
--wg-accent:       #0078d4   Office blue
--wg-accent-dark:  #005fa3   hover/pressed accent
--wg-accent-glow:  rgba(0,120,212,0.12)   focus ring, hover tints
--wg-danger:       #c4262e   destructive actions
```

Dark mode: `applyTheme()` in taskpane.ts detects dark via luminance check on `bodyBackgroundColor`,
toggles `body.wg-dark`, and overrides the above vars with dark palette values.

### Button classes (in taskpane.html `<style>` block)

| Class           | bg        | text      | border    | Hover                 |
| --------------- | --------- | --------- | --------- | --------------------- |
| `btn-primary`   | `#0078d4` | white     | `#0078d4` | `#005fa3`             |
| `btn-refresh`   | white     | `#0078d4` | `#0078d4` | inverts to solid blue |
| `btn-secondary` | `#eeeeee` | `#1a1a1a` | `#d0d0d0` | `#e0e0e0`             |
| `btn-danger`    | `#c4262e` | white     | `#c4262e` | `#a31f26`             |

### Where each button class is used

- `btn-primary`: Commit, Create branch, ↩ Restore (inside log card), Yes restore (confirm)
- `btn-refresh`: Refresh (Log tab), Refresh (Branches tab)
- `btn-secondary`: Init repo, Export history as JSON, Cancel (confirm), Export
- `btn-danger`: ⟷ Show diff (log card), ⤢ Merge in (branch row)

### Tab styling

- Emojis in tab labels: 💾 Commit, 📜 Log, 🌿 Branches
- No emojis in buttons (except ⚙️ Init repo as a special one-time action)
- Active tab: blue underline + bold + `rgba(0,120,212,0.08)` tinted background

### Log tab behaviour

- Accordion: only one card open at a time - opening one collapses the previously open one
- Expanded card: NO blue outline or glow - open state is self-evident from visible detail + rotated chevron
- Export JSON button: always pinned to bottom via `panel-footer` flex layout

### Branch list

- Rendered as a bordered card (`border: 1px solid var(--wg-border); border-radius: 6px`)
- Active branch row: `rgba(0,120,212,0.10)` bg + "current" pill badge, no action buttons
- Inactive rows: `⇄ Switch` (btn-primary), `⤢ Merge in` (btn-danger)

### Init state

- On load: Commit input + Commit button disabled; Init repo enabled
- After init (or if repo already exists on open): Commit input + Commit button enabled; Init repo disabled
- `checkRepoExists()` runs in `Office.onReady` - checks XML store for branch refs

### Enter key support

- Commit message input → triggers commit
- Branch name input → triggers create branch

---

## VCS class notes

### ObjectStore.ts

- Loads/saves entire repo XML as one Custom XML Part
- On save: deletes old part, adds new one (no update API in Office JS)

### DiffEngine.ts

- `normaliseOoxml()` strips volatile OOXML attributes (`w:rsidR`, `w:paraId`, `w:textId`, etc.)
- Only used in merge path - stored deltas are NOT normalised

### CommitManager.ts

- `reconstruct(hash)` - walks parent chain, replays deltas, returns OOXML string
- `checkout(hash)` - tracking ON (`trackMineOnly`) - for Restore
- `checkoutSilent(hash)` - tracking OFF - for branch switch
- `merge(theirBranch)` - finds common ancestor, calls MergeManager, stores merge commit as snapshot
- Root commit: `parent = "00root"`, `encoding = "snapshot"` (full OOXML, not delta)
- Merge commits: always stored as snapshots to avoid ambiguous parent-chain replay

### MergeManager.ts

- Pure computation - NO Word API calls
- Normalises all three inputs before any diff/patch
- Throws `"CONFLICT: ..."` on failure - caller catches, shows warning toast, does NOT modify document

---

## Dev workflow

```powershell
cd wordgit
npm start          # starts webpack + opens Word with add-in sideloaded
```

- Save any file → webpack hot-reloads → task pane refreshes
- Watch for `webpack compiled successfully` in terminal
- Keep terminal open - closing it stops the dev server

## Resuming on a new machine

1. Install Node.js + VS Code
2. Copy the `wordgit/` folder
3. `npm install` then `npm start`

---

## Current source files - taskpane (Phase 6 final)

These are the exact current versions. Copy them verbatim - do not reconstruct from descriptions.

---

### `src/taskpane/taskpane.html`

```html
<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>WordGit</title>
    <script
      src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
      type="text/javascript"
    ></script>
    <link rel="stylesheet" href="taskpane.css" />
    <style>
      /*
       * Button colours live here, not in taskpane.css.
       *
       * Why: office.js injects its own Fabric stylesheet into the page after
       * all <link> stylesheets have loaded, resetting button colours. A <style>
       * block in <head> is part of the document cascade and always wins over
       * any externally injected stylesheet - no !important needed.
       *
       * Button roles:
       *   btn-primary  - solid blue  - ONE main action per panel
       *   btn-refresh  - blue outline, inverts on hover - frequent secondary action
       *   btn-secondary- gray        - low-priority actions
       *   btn-danger   - solid red   - destructive/document-replacing actions
       */

      .btn {
        display: block;
        width: 100%;
        padding: 8px 14px;
        margin-bottom: 8px;
        border-radius: 5px;
        font-family: "Segoe UI", system-ui, sans-serif;
        font-size: 13px;
        font-weight: 500;
        cursor: pointer;
        text-align: center;
        border: 1.5px solid transparent;
        transition:
          background-color 0.15s,
          color 0.15s,
          border-color 0.15s,
          box-shadow 0.15s;
      }
      .btn:disabled {
        opacity: 0.35;
        cursor: not-allowed;
        pointer-events: none;
      }

      .btn-primary {
        background-color: #0078d4;
        color: #ffffff;
        border-color: #0078d4;
      }
      .btn-primary:hover {
        background-color: #005fa3;
        border-color: #005fa3;
      }
      .btn-primary:active {
        background-color: #004f8f;
        border-color: #004f8f;
      }

      .btn-refresh {
        background-color: #ffffff;
        color: #0078d4;
        border-color: #0078d4;
      }
      .btn-refresh:hover {
        background-color: #0078d4;
        color: #ffffff;
        border-color: #0078d4;
      }
      .btn-refresh:active {
        background-color: #005fa3;
        color: #ffffff;
        border-color: #005fa3;
      }

      .btn-secondary {
        background-color: #eeeeee;
        color: #1a1a1a;
        border-color: #d0d0d0;
      }
      .btn-secondary:hover {
        background-color: #e0e0e0;
        border-color: #bdbdbd;
      }
      .btn-secondary:active {
        background-color: #d4d4d4;
      }

      .btn-danger {
        background-color: #c4262e;
        color: #ffffff;
        border-color: #c4262e;
      }
      .btn-danger:hover {
        background-color: #a31f26;
        border-color: #a31f26;
      }
      .btn-danger:active {
        background-color: #8c1a20;
        border-color: #8c1a20;
      }

      .tab-btn.active {
        background-color: rgba(0, 120, 212, 0.08);
      }

      body.wg-dark .btn-refresh {
        background-color: transparent;
        color: #4db6ff;
        border-color: #4db6ff;
      }
      body.wg-dark .btn-refresh:hover {
        background-color: #4db6ff;
        color: #000000;
      }
      body.wg-dark .btn-secondary {
        background-color: #2d2d2d;
        color: #ececec;
        border-color: #3a3a3a;
      }
      body.wg-dark .btn-secondary:hover {
        background-color: #383838;
      }
      body.wg-dark .btn-danger {
        background-color: #c4262e;
        color: #ffffff;
      }
      body.wg-dark .btn-danger:hover {
        background-color: #a31f26;
      }
    </style>
  </head>
  <body>
    <div id="toast-container"></div>
    <div id="app">
      <div id="tabs">
        <button class="tab-btn active" data-tab="commit">💾 Commit</button>
        <button class="tab-btn" data-tab="log">📜 Log</button>
        <button class="tab-btn" data-tab="branches">🌿 Branches</button>
      </div>

      <!-- COMMIT panel -->
      <div id="panel-commit" class="panel active">
        <div class="panel-body">
          <p class="wg-section">Commit</p>
          <label class="wg-label" for="commit-msg">Message (optional)</label>
          <input
            class="wg-input"
            type="text"
            id="commit-msg"
            placeholder="e.g. Draft intro section"
            disabled
          />
          <button class="btn btn-primary" id="btn-commit" disabled>Commit</button>
          <button class="btn btn-secondary" id="btn-init">Init repo</button>
        </div>
      </div>

      <!-- LOG panel -->
      <div id="panel-log" class="panel">
        <div class="panel-body">
          <p class="wg-section">History</p>
          <button class="btn btn-refresh" id="btn-refresh-log">Refresh</button>
          <ul id="log-list">
            <li class="wg-hint">No history loaded yet.</li>
          </ul>
        </div>
        <div class="panel-footer">
          <button class="btn btn-secondary" id="btn-export-json" style="margin:0;">
            Export history as JSON
          </button>
        </div>
      </div>

      <!-- BRANCHES panel -->
      <div id="panel-branches" class="panel">
        <div class="panel-body">
          <p class="wg-section">Branches</p>
          <button class="btn btn-refresh" id="btn-refresh-branches">Refresh</button>

          <!-- Branch graph (Phase 7) - rendered SVG DAG above branch list -->
          <div id="branch-graph" class="branch-graph-container">
            <span class="wg-hint">Click Refresh to load graph.</span>
          </div>

          <div id="branch-list"></div>
          <hr class="wg-divider" />
          <p class="wg-section">Create branch</p>
          <label class="wg-label" for="new-branch-name">Branch name</label>
          <input
            class="wg-input"
            type="text"
            id="new-branch-name"
            placeholder="e.g. draft-chapter-2"
          />
          <button class="btn btn-primary" id="btn-create-branch">Create branch</button>
        </div>
      </div>
    </div>
  </body>
</html>
```

---

### `src/taskpane/taskpane.css`

```css
/* ============================================================
   WordGit - taskpane.css
   Layout, components, design tokens only.
   Button colours are intentionally in taskpane.html <style>
   block - see CRITICAL pitfalls section for why.
   ============================================================ */

html,
body {
  width: 100%;
  height: 100%;
  margin: 0;
  padding: 0;
}
*,
*::before,
*::after {
  box-sizing: border-box;
}
ul {
  list-style: none;
  margin: 0;
  padding: 0;
}

:root {
  --wg-bg: #f5f5f5;
  --wg-surface: #ffffff;
  --wg-surface-alt: #fafafa;
  --wg-surface-btn: #eeeeee;
  --wg-border: #e0e0e0;
  --wg-border-input: #c8c8c8;
  --wg-text: #1a1a1a;
  --wg-text-muted: #6b6b6b;
  --wg-text-faint: #a0a0a0;
  --wg-accent: #0078d4;
  --wg-accent-dark: #005fa3;
  --wg-accent-glow: rgba(0, 120, 212, 0.12);
  --wg-accent-text: #004f8f;
  --wg-danger: #c4262e;
  --wg-warn-bg: #fff8e1;
  --wg-warn-border: #f0b400;
  --wg-warn-text: #5c3d00;
}

body {
  font-family:
    "Segoe UI",
    system-ui,
    -apple-system,
    sans-serif;
  font-size: 13px;
  line-height: 1.5;
  background: var(--wg-bg);
  color: var(--wg-text);
  -webkit-font-smoothing: antialiased;
}

#app {
  display: flex;
  flex-direction: column;
  height: 100vh;
}

.panel {
  display: none;
  flex-direction: column;
  flex: 1;
  overflow: hidden;
}
.panel.active {
  display: flex;
}
.panel-body {
  flex: 1;
  overflow-y: auto;
  padding: 16px;
}
.panel-footer {
  flex-shrink: 0;
  padding: 10px 16px;
  background: var(--wg-surface);
  border-top: 1px solid var(--wg-border);
}

#tabs {
  display: flex;
  flex-shrink: 0;
  background: var(--wg-surface);
  border-bottom: 2px solid var(--wg-accent);
}
.tab-btn {
  flex: 1;
  padding: 10px 4px;
  border: none;
  background: transparent;
  font-family: inherit;
  font-size: 12px;
  font-weight: 500;
  color: var(--wg-text-muted);
  cursor: pointer;
  border-bottom: 2px solid transparent;
  margin-bottom: -2px;
  transition: color 0.15s;
}
.tab-btn:hover {
  color: var(--wg-accent);
}
.tab-btn.active {
  color: var(--wg-accent);
  font-weight: 600;
  border-bottom-color: var(--wg-accent);
}

.wg-section {
  margin: 0 0 10px;
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--wg-text-muted);
}
.wg-divider {
  border: none;
  border-top: 1px solid var(--wg-border);
  margin: 14px 0;
}
.wg-hint {
  font-size: 11px;
  color: var(--wg-text-faint);
  margin-bottom: 8px;
}
.wg-label {
  display: block;
  margin-bottom: 5px;
  font-size: 12px;
  font-weight: 600;
  color: var(--wg-text-muted);
}

.wg-input {
  width: 100%;
  padding: 7px 10px;
  margin-bottom: 10px;
  font-family: inherit;
  font-size: 13px;
  background: var(--wg-surface);
  color: var(--wg-text);
  border: 1px solid var(--wg-border-input);
  border-radius: 4px;
  transition:
    border-color 0.15s,
    box-shadow 0.15s;
}
.wg-input::placeholder {
  color: var(--wg-text-faint);
}
.wg-input:hover:not(:disabled) {
  border-color: #a8a8a8;
}
.wg-input:focus {
  outline: none;
  border-color: var(--wg-accent);
  box-shadow: 0 0 0 3px var(--wg-accent-glow);
}
.wg-input:disabled {
  opacity: 0.38;
  cursor: not-allowed;
  background: var(--wg-bg);
}

#log-list {
  margin: 0;
  padding: 0;
}

.log-card {
  border: 1px solid var(--wg-border);
  border-radius: 6px;
  margin-bottom: 8px;
  background: var(--wg-surface);
  overflow: hidden;
  transition: border-color 0.15s;
}
.log-card.expanded {
  border-color: var(--wg-border);
}

.log-summary {
  padding: 10px 12px;
  cursor: pointer;
  display: flex;
  align-items: center;
  gap: 8px;
  user-select: none;
}
.log-summary:hover {
  background: var(--wg-accent-glow);
}

.log-chevron {
  font-size: 9px;
  color: var(--wg-text-faint);
  transition: transform 0.15s;
  flex-shrink: 0;
}
.log-card.expanded .log-chevron {
  transform: rotate(90deg);
  color: var(--wg-accent);
}

.log-hash {
  font-family: "Cascadia Code", "Consolas", monospace;
  font-size: 11px;
  color: var(--wg-text-muted);
  flex-shrink: 0;
  background: var(--wg-bg);
  padding: 1px 5px;
  border-radius: 3px;
  border: 1px solid var(--wg-border);
}
.log-msg {
  font-weight: 600;
  font-size: 12px;
  flex: 1;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.log-detail {
  display: none;
  padding: 10px 12px 12px;
  border-top: 1px solid var(--wg-border);
  background: var(--wg-surface-alt);
}
.log-card.expanded .log-detail {
  display: block;
}

.log-detail-meta {
  font-size: 11px;
  color: var(--wg-text-muted);
  margin-bottom: 10px;
  line-height: 1.6;
}
.log-detail-meta code {
  font-family: "Cascadia Code", "Consolas", monospace;
  font-size: 10px;
  background: var(--wg-bg);
  padding: 1px 4px;
  border-radius: 3px;
  border: 1px solid var(--wg-border);
}

.log-actions {
  display: flex;
  gap: 6px;
  margin-top: 2px;
}
.log-actions .btn {
  flex: 1;
  margin: 0;
  padding: 5px 8px;
  font-size: 11px;
}

.card-confirm {
  display: none;
  margin-top: 8px;
  padding: 10px 12px;
  background: var(--wg-warn-bg);
  border: 1px solid var(--wg-warn-border);
  border-radius: 5px;
}
.card-confirm.visible {
  display: block;
}
.card-confirm-msg {
  font-size: 12px;
  font-weight: 600;
  color: var(--wg-warn-text);
  margin-bottom: 8px;
}
.card-confirm-btns {
  display: flex;
  gap: 6px;
}
.card-confirm-btns .btn {
  flex: 1;
  margin: 0;
  padding: 5px 8px;
  font-size: 11px;
}

.branch-graph-container {
  margin-bottom: 12px;
  border: 1px solid var(--wg-border);
  border-radius: 6px;
  overflow: hidden;
  min-height: 32px;
  background: var(--wg-surface);
  padding: 8px 4px;
}
.branch-graph-container:empty {
  display: none;
}

.branch-row {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 12px;
  border-bottom: 1px solid var(--wg-border);
  font-size: 12px;
  background: var(--wg-surface);
  transition: background 0.12s;
}
.branch-row:last-child {
  border-bottom: none;
}
.branch-row:hover {
  background: var(--wg-accent-glow);
}
.branch-row.active {
  background: rgba(0, 120, 212, 0.1);
  font-weight: 600;
}

.branch-name {
  flex: 1;
}
.branch-badge {
  font-size: 10px;
  font-weight: 600;
  color: var(--wg-accent);
  background: var(--wg-accent-glow);
  padding: 1px 6px;
  border-radius: 10px;
  border: 1px solid rgba(0, 120, 212, 0.25);
  flex-shrink: 0;
}
.branch-row .btn {
  width: auto;
  flex-shrink: 0;
  margin: 0;
  padding: 4px 10px;
  font-size: 11px;
}

#toast-container {
  position: fixed;
  bottom: 12px;
  left: 10px;
  right: 10px;
  display: flex;
  flex-direction: column;
  gap: 6px;
  z-index: 9999;
  pointer-events: none;
}
.toast {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  padding: 10px 12px;
  border-radius: 6px;
  border-left: 4px solid;
  font-size: 12px;
  line-height: 1.4;
  pointer-events: all;
  animation: toast-in 0.2s ease;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
}
.toast-out {
  animation: toast-out 0.2s ease forwards;
}
@keyframes toast-in {
  from {
    opacity: 0;
    transform: translateY(6px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}
@keyframes toast-out {
  from {
    opacity: 1;
    transform: translateY(0);
  }
  to {
    opacity: 0;
    transform: translateY(6px);
  }
}

.toast-success {
  background: #f0faf0;
  border-color: #107c10;
  color: #0a4f0a;
}
.toast-error {
  background: #fdf0f0;
  border-color: #c4262e;
  color: #6b0000;
}
.toast-warning {
  background: var(--wg-warn-bg);
  border-color: var(--wg-warn-border);
  color: var(--wg-warn-text);
}
.toast-info {
  background: #f0f6ff;
  border-color: #0078d4;
  color: #003f73;
}

.toast-icon {
  flex-shrink: 0;
  margin-top: 1px;
}
.toast-msg {
  flex: 1;
}
.toast-close {
  background: none;
  border: none;
  cursor: pointer;
  font-size: 16px;
  line-height: 1;
  padding: 0 2px;
  color: inherit;
  opacity: 0.5;
  flex-shrink: 0;
}
.toast-close:hover {
  opacity: 1;
}
```

---

### `src/taskpane/taskpane.ts`

```typescript
/* global Word, Office */

import { ObjectStore } from "../vcs/ObjectStore";
import { BranchManager } from "../vcs/BranchManager";
import { CommitManager } from "../vcs/CommitManager";
import { DiffEngine } from "../vcs/DiffEngine";
import { buildGraph, GraphLayout, GraphCommit, GraphRef } from "../vcs/GraphBuilder";

// ─── Init state ───────────────────────────────────────────────────────────────
// On load, read the XML store. If branch refs exist, repo was previously
// initialised - enable commit immediately. Otherwise keep disabled until init.

function setRepoReady(ready: boolean): void {
  const input = document.getElementById("commit-msg") as HTMLInputElement;
  const commit = document.getElementById("btn-commit") as HTMLButtonElement;
  const init = document.getElementById("btn-init") as HTMLButtonElement;
  input.disabled = !ready;
  commit.disabled = !ready;
  init.disabled = ready;
}

async function checkRepoExists(): Promise<void> {
  try {
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const repo = await store.load();
      await context.sync();
      setRepoReady(repo.refs.length > 0);
    });
  } catch {
    setRepoReady(false);
  }
}

// ─── Toast ────────────────────────────────────────────────────────────────────

type ToastType = "success" | "error" | "warning" | "info";
const TOAST_ICONS: Record<ToastType, string> = {
  success: "✅",
  error: "❌",
  warning: "⚠️",
  info: "ℹ️",
};

function showToast(msg: string, type: ToastType = "info", durationMs = 4000): void {
  const container = document.getElementById("toast-container")!;
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.innerHTML = `
    <span class="toast-icon">${TOAST_ICONS[type]}</span>
    <span class="toast-msg">${msg}</span>
    <button class="toast-close" aria-label="Dismiss">×</button>`;
  container.appendChild(el);
  const dismiss = () => {
    el.classList.add("toast-out");
    el.addEventListener("animationend", () => el.remove());
  };
  el.querySelector(".toast-close")!.addEventListener("click", dismiss);
  setTimeout(dismiss, durationMs);
}

// ─── Theme ────────────────────────────────────────────────────────────────────
// Detects dark mode by measuring luminance of Office background colour.
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
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(`panel-${btn.dataset.tab}`)?.classList.add("active");
    });
  });
}

// ─── Init repo ────────────────────────────────────────────────────────────────

async function handleInit(): Promise<void> {
  try {
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const bMgr = new BranchManager(store);
      const mgr = new CommitManager(store, bMgr, new DiffEngine());
      await mgr.init();
      await context.sync();
    });
    setRepoReady(true);
    showToast("Repo initialised.", "success");
  } catch (e: any) {
    showToast("Init failed: " + e.message, "error");
  }
}

// ─── Commit ───────────────────────────────────────────────────────────────────

async function handleCommit(): Promise<void> {
  const msgEl = document.getElementById("commit-msg") as HTMLInputElement;
  const message = msgEl.value.trim() || "Update";
  try {
    let hash = "";
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const bMgr = new BranchManager(store);
      const mgr = new CommitManager(store, bMgr, new DiffEngine());
      hash = await mgr.commit(message);
      await context.sync();
    });
    msgEl.value = "";
    showToast(`Committed ${hash.slice(0, 6)} - "${message}"`, "success");
  } catch (e: any) {
    showToast("Commit failed: " + e.message, "error");
  }
}

// ─── Export JSON ──────────────────────────────────────────────────────────────

async function handleExportJson(): Promise<void> {
  try {
    let commits: object[] = [];
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const bMgr = new BranchManager(store);
      const mgr = new CommitManager(store, bMgr, new DiffEngine());
      commits = (await mgr.log()) as any;
      await context.sync();
    });
    const json = JSON.stringify(commits, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const date = new Date().toISOString().slice(0, 10);
    const a = document.createElement("a");
    a.href = url;
    a.download = `wordgit-history-${date}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast("History exported.", "success");
  } catch (e: any) {
    showToast("Export failed: " + e.message, "error");
  }
}

// ─── Diff vs parent ───────────────────────────────────────────────────────────

async function handleDiffVsParent(hash: string, parentHash: string): Promise<void> {
  try {
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const bMgr = new BranchManager(store);
      const mgr = new CommitManager(store, bMgr, new DiffEngine());
      const fromOoxml = await mgr.reconstruct(parentHash);
      const toOoxml = await mgr.reconstruct(hash);
      const body = context.document.body;
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      body.insertOoxml(fromOoxml, Word.InsertLocation.replace);
      await context.sync();
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;
      body.insertOoxml(toOoxml, Word.InsertLocation.replace);
      await context.sync();
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    });
    showToast("Diff loaded. Accept all tracked changes when done reviewing.", "success");
  } catch (e: any) {
    showToast("Error showing diff: " + e.message, "error");
  }
}

// ─── Restore ─────────────────────────────────────────────────────────────────

async function handleRestore(hash: string): Promise<void> {
  try {
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const bMgr = new BranchManager(store);
      const mgr = new CommitManager(store, bMgr, new DiffEngine());
      await mgr.checkout(hash);
      await context.sync();
    });
    showToast(
      `Restored to ${hash.slice(0, 6)}. Review tracked changes and accept when ready.`,
      "success"
    );
  } catch (e: any) {
    showToast("Restore failed: " + e.message, "error");
  }
}

// ─── Log ─────────────────────────────────────────────────────────────────────

async function handleRefreshLog(): Promise<void> {
  const list = document.getElementById("log-list")!;
  list.innerHTML = "";
  try {
    let commits: Array<{
      hash: string;
      parent: string;
      message: string;
      timestamp: string;
      author: string;
    }> = [];
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const bMgr = new BranchManager(store);
      const mgr = new CommitManager(store, bMgr, new DiffEngine());
      commits = (await mgr.log()) as any;
      await context.sync();
    });

    if (commits.length === 0) {
      list.innerHTML = `<li style="color:var(--wg-text-muted);padding:4px 0;font-size:12px;">No commits yet.</li>`;
      return;
    }

    commits.forEach((c) => {
      const li = document.createElement("li");
      const ts = new Date(c.timestamp).toLocaleString();
      const isRoot = c.parent === "00root";
      li.className = "log-card";
      li.innerHTML = `
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
      list.appendChild(li);

      li.querySelector(".log-summary")!.addEventListener("click", () => {
        const isExpanded = li.classList.contains("expanded");
        document
          .querySelectorAll(".log-card.expanded")
          .forEach((c) => c.classList.remove("expanded"));
        if (!isExpanded) li.classList.add("expanded");
      });
      li.querySelector(`[data-ask-restore="${c.hash}"]`)!.addEventListener("click", () =>
        document.getElementById(`confirm-${c.hash}`)!.classList.add("visible")
      );
      li.querySelector(`[data-cancel="${c.hash}"]`)!.addEventListener("click", () =>
        document.getElementById(`confirm-${c.hash}`)!.classList.remove("visible")
      );
      li.querySelector(`[data-restore="${c.hash}"]`)!.addEventListener("click", () =>
        handleRestore(c.hash)
      );
      li.querySelector(`[data-diff="${c.hash}"]`)?.addEventListener("click", () =>
        handleDiffVsParent(c.hash, c.parent)
      );
    });
  } catch (e: any) {
    showToast("Error loading log: " + e.message, "error");
  }
}

// ─── Branch graph (Phase 7) ───────────────────────────────────────────────────
// Renders an SVG commit DAG inside #branch-graph in the Branches panel.
// Called from handleRefreshBranches() after the branch list is rendered.

const LANE_COLOURS = ["#0078d4", "#107c10", "#8764b8", "#d04a02", "#00b7c3", "#c4262e"];
const COL_W = 28,
  ROW_H = 36,
  R = 6,
  PAD_L = 14,
  PAD_T = 18;

function gcColour(col: number): string {
  return LANE_COLOURS[col % LANE_COLOURS.length];
}
function gcX(col: number): number {
  return PAD_L + col * COL_W;
}
function gcY(row: number): number {
  return PAD_T + row * ROW_H;
}
function gcTrunc(s: string, max: number): string {
  return s.length > max ? s.slice(0, max - 1) + "…" : s;
}
function gcTs(iso: string): string {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  const p = (n: number) => n.toString().padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}
function gcEsc(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

async function renderBranchGraph(): Promise<void> {
  const container = document.getElementById("branch-graph")!;
  container.innerHTML = `<span class="wg-hint">Loading…</span>`;
  try {
    let svg = "";
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const repo = await store.load();
      await context.sync();

      // Collect all commits reachable from any ref
      const reachable = new Map<string, GraphCommit>();
      for (const ref of repo.refs) {
        let hash = ref.tip;
        while (hash && hash !== "00root") {
          if (reachable.has(hash)) break;
          const c = repo.objects.find((o) => o.hash === hash);
          if (!c) break;
          reachable.set(hash, {
            hash: c.hash,
            parent: c.parent,
            message: c.message,
            timestamp: c.timestamp,
            author: c.author,
          });
          hash = c.parent;
        }
      }

      const layout: GraphLayout = buildGraph(
        Array.from(reachable.values()),
        repo.refs as GraphRef[]
      );
      const dark = document.body.classList.contains("wg-dark");
      const textFill = dark ? "#d0d0d0" : "#1a1a1a";
      const mutedFill = dark ? "#707070" : "#6b6b6b";
      const bgFill = dark ? "#252525" : "#ffffff";
      const pillBg = dark ? "rgba(0,120,212,0.22)" : "rgba(0,120,212,0.10)";
      const W = Math.max(PAD_L * 2 + layout.cols * COL_W, 200);
      const H = PAD_T * 2 + layout.rows * ROW_H;

      const lines: string[] = [
        `<svg xmlns="http://www.w3.org/2000/svg" width="100%" viewBox="0 0 ${W} ${H}" style="display:block;overflow:visible">`,
        `<rect width="${W}" height="${H}" fill="${bgFill}" rx="6"/>`,
      ];

      // Edges
      for (const e of layout.edges) {
        const x1 = gcX(e.fromCol),
          y1 = gcY(e.fromRow);
        const x2 = gcX(e.toCol),
          y2 = gcY(e.toRow);
        const col = gcColour(e.toCol);
        if (e.fromCol === e.toCol) {
          lines.push(
            `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${col}" stroke-width="2" opacity="0.7"/>`
          );
        } else {
          const midY = (y1 + y2) / 2;
          lines.push(
            `<path d="M ${x1} ${y1} C ${x1} ${midY}, ${x2} ${midY}, ${x2} ${y2}" fill="none" stroke="${col}" stroke-width="2" opacity="0.65"/>`
          );
        }
      }

      // Nodes
      for (const n of layout.nodes) {
        const x = gcX(n.col),
          y = gcY(n.row);
        const col = gcColour(n.col);
        const isTip = n.refs.length > 0;
        lines.push(
          `<circle cx="${x}" cy="${y}" r="${R}" fill="${isTip ? col : bgFill}" stroke="${col}" stroke-width="2"/>`
        );
        const lx = x + R + 8;
        lines.push(
          `<text x="${lx}" y="${y - 4}" font-family="Cascadia Code,Consolas,monospace" font-size="9" fill="${mutedFill}">${n.hash.slice(0, 6)}</text>`
        );
        lines.push(
          `<text x="${lx}" y="${y + 8}" font-family="Segoe UI,system-ui,sans-serif" font-size="10" fill="${textFill}" font-weight="${isTip ? 600 : 400}">${gcEsc(gcTrunc(n.message || "(no message)", 28))}</text>`
        );
        lines.push(
          `<text x="${lx}" y="${y + 19}" font-family="Segoe UI,system-ui,sans-serif" font-size="9" fill="${mutedFill}">${gcTs(n.ts)}</text>`
        );
        let pillX = lx;
        for (const refName of n.refs) {
          const pillW = refName.length * 5.5 + 10;
          lines.push(
            `<rect x="${pillX}" y="${y - ROW_H / 2 + 2}" width="${pillW}" height="13" rx="6" fill="${pillBg}" stroke="rgba(0,120,212,0.35)" stroke-width="0.8"/>`
          );
          lines.push(
            `<text x="${pillX + pillW / 2}" y="${y - ROW_H / 2 + 12}" text-anchor="middle" font-family="Segoe UI,system-ui,sans-serif" font-size="9" font-weight="600" fill="${col}">${gcEsc(refName)}</text>`
          );
          pillX += pillW + 4;
        }
      }

      lines.push("</svg>");
      svg = lines.join("\n");
    });
    container.innerHTML = svg || `<span class="wg-hint">No commits yet.</span>`;
  } catch (e: any) {
    container.innerHTML = `<span class="wg-hint" style="color:var(--wg-danger)">Graph error: ${e.message}</span>`;
  }
}

// ─── Branches ────────────────────────────────────────────────────────────────

async function handleRefreshBranches(): Promise<void> {
  const container = document.getElementById("branch-list")!;
  container.innerHTML = "";
  try {
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const bMgr = new BranchManager(store);
      const current = await bMgr.currentBranch();
      const refs = await bMgr.listBranches();
      await context.sync();

      if (refs.length === 0) {
        container.innerHTML = `<span class="wg-hint">No branches yet. Init the repo first.</span>`;
        return;
      }

      refs.forEach((branch) => {
        const isActive = branch.name === current;
        const row = document.createElement("div");
        row.className = `branch-row${isActive ? " active" : ""}`;
        row.innerHTML = `
          <span class="branch-name">${branch.name}</span>
          ${
            isActive
              ? `<span class="branch-badge">current</span>`
              : `<button class="btn btn-primary" data-switch="${branch.name}">⇄ Switch</button>
               <button class="btn btn-danger"  data-merge="${branch.name}">⤢ Merge in</button>`
          }`;
        container.appendChild(row);
        row
          .querySelector(`[data-switch="${branch.name}"]`)
          ?.addEventListener("click", () => handleSwitchBranch(branch.name));
        row
          .querySelector(`[data-merge="${branch.name}"]`)
          ?.addEventListener("click", () => handleMerge(branch.name));
      });
    });
    // Render DAG graph after branch list is ready (Phase 7)
    await renderBranchGraph();
  } catch (e: any) {
    showToast("Error loading branches: " + e.message, "error");
  }
}

async function handleSwitchBranch(name: string): Promise<void> {
  try {
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const bMgr = new BranchManager(store);
      const mgr = new CommitManager(store, bMgr, new DiffEngine());
      await bMgr.switchBranch(name);
      const tip = await bMgr.getTip(name);
      await mgr.checkoutSilent(tip);
      await context.sync();
    });
    showToast(`Switched to branch "${name}".`, "success");
    await handleRefreshBranches();
  } catch (e: any) {
    showToast("Switch failed: " + e.message, "error");
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
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const bMgr = new BranchManager(store);
      await bMgr.createBranch(name);
      await context.sync();
    });
    nameEl.value = "";
    showToast(`Branch "${name}" created.`, "success");
    await handleRefreshBranches();
  } catch (e: any) {
    showToast("Create branch failed: " + e.message, "error");
  }
}

async function handleMerge(theirBranch: string): Promise<void> {
  try {
    await Word.run(async (context) => {
      const store = new ObjectStore(context);
      const bMgr = new BranchManager(store);
      const mgr = new CommitManager(store, bMgr, new DiffEngine());
      await mgr.merge(theirBranch);
      await context.sync();
    });
    showToast(`Merged "${theirBranch}". Review tracked changes and accept when ready.`, "success");
  } catch (e: any) {
    if (e.message?.includes("CONFLICT")) {
      showToast(`Merge conflict: ${e.message}`, "warning", 8000);
    } else {
      showToast("Merge failed: " + e.message, "error");
    }
  }
}

// ─── Office.onReady ───────────────────────────────────────────────────────────

Office.onReady(async () => {
  applyTheme(Office.context.officeTheme);

  if (Office.EventType) {
    Office.context.document.addHandlerAsync(
      Office.EventType.officeThemeChanged as any,
      (args: any) => applyTheme(args.officeTheme)
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

  await checkRepoExists();
});
```

---

## Current source files - VCS engine (Phase 6 final)

These are the exact current versions. Copy them verbatim into `src/vcs/`.

---

### `src/vcs/ObjectStore.ts`

```typescript
/* global Word */

const NS = "http://schemas.wordgit.local/vcs/1";

export interface RepoXml {
  head: string;
  author: string;
  refs: Array<{ name: string; tip: string }>;
  objects: Array<CommitXml>;
}

export interface CommitXml {
  hash: string;
  parent: string;
  author: string;
  timestamp: string;
  message: string;
  /**
   * Payload string. Meaning depends on `encoding`:
   *   - "snapshot"    : raw OOXML (uncompressed)
   *   - "dmp"         : raw diff-match-patch patch text (uncompressed)
   *   - "snapshot-lz" : OOXML compressed with LZString.compressToBase64()
   *   - "dmp-lz"      : dmp patch text compressed with LZString.compressToBase64()
   */
  delta: string;
  encoding: "dmp" | "snapshot" | "dmp-lz" | "snapshot-lz";
}

const EMPTY_REPO: RepoXml = {
  head: "refs/heads/main",
  author: "Author",
  refs: [],
  objects: [],
};

function repoToXml(repo: RepoXml): string {
  const refs = repo.refs
    .map(r => `<branch name="${esc(r.name)}" tip="${esc(r.tip)}"/>`)
    .join("
    ");
  const objects = repo.objects
    .map(
      c => `<commit hash="${esc(c.hash)}" parent="${esc(c.parent)}">
      <author>${esc(c.author)}</author>
      <timestamp>${esc(c.timestamp)}</timestamp>
      <message>${esc(c.message)}</message>
      <delta encoding="${esc(c.encoding)}">${escapeCdata(c.delta)}</delta>
    </commit>`
    )
    .join("
    ");
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

function xmlToRepo(xml: string): RepoXml {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");

  const head = doc.querySelector("HEAD")?.textContent ?? "refs/heads/main";
  const author = doc.querySelector("config > author")?.textContent ?? "Author";

  const refs: RepoXml["refs"] = [];
  doc.querySelectorAll("refs > branch").forEach(el => {
    refs.push({
      name: el.getAttribute("name") ?? "",
      tip: el.getAttribute("tip") ?? "",
    });
  });

  const objects: CommitXml[] = [];
  doc.querySelectorAll("objects > commit").forEach(el => {
    objects.push({
      hash: el.getAttribute("hash") ?? "",
      parent: el.getAttribute("parent") ?? "",
      author: el.querySelector("author")?.textContent ?? "",
      timestamp: el.querySelector("timestamp")?.textContent ?? "",
      message: el.querySelector("message")?.textContent ?? "",
      delta: el.querySelector("delta")?.textContent ?? "",
      encoding: (el.querySelector("delta")?.getAttribute("encoding") ?? "dmp") as CommitXml["encoding"],
    });
  });

  return { head, author, refs, objects };
}

function esc(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeCdata(s: string): string {
  return `<![CDATA[${s.replace(/\]\]>/g, "]]]]><![CDATA[>")}]]>`;
}

export class ObjectStore {
  private context: Word.RequestContext;
  private partId: string | null = null;

  constructor(context: Word.RequestContext) {
    this.context = context;
  }

  /** Load repo XML from Custom XML Part, or return empty repo if none exists. */
  async load(): Promise<RepoXml> {
    const parts = this.context.document.customXmlParts;
    const items = parts.getByNamespace(NS);
    items.load("items");
    await this.context.sync();

    if (items.items.length === 0) {
      return { ...EMPTY_REPO, refs: [], objects: [] };
    }

    const part = items.items[0];
    this.partId = part.id;
    const xmlResult = part.getXml();
    await this.context.sync();
    return xmlToRepo(xmlResult.value);
  }

  /** Save repo XML back to Custom XML Part (create or replace). */
  async save(repo: RepoXml): Promise<void> {
    const xml = repoToXml(repo);
    const parts = this.context.document.customXmlParts;

    if (this.partId) {
      // Delete old part and re-add (Office JS has no "update XML" API)
      const old = parts.getItem(this.partId);
      old.delete();
      await this.context.sync();
      this.partId = null;
    }

    const newPart = parts.add(xml);
    newPart.load("id");
    await this.context.sync();
    this.partId = newPart.id;
  }
}
```

---

### `src/vcs/DiffEngine.ts`

```typescript
import * as dmp from "diff-match-patch";

const DMP = new dmp.diff_match_patch();

/**
 * Strip volatile OOXML attributes that Word regenerates on every save/checkout.
 * Only used in the merge path - stored deltas are kept as-is.
 */
const VOLATILE_ATTR_RE = /\s+w:(?:rsid\w*|paraId|textId|id)="[^"]*"/g;

export function normaliseOoxml(ooxml: string): string {
  return ooxml.replace(VOLATILE_ATTR_RE, "");
}

export class DiffEngine {
  /** Create a dmp patch string from oldText → newText. */
  makePatch(oldText: string, newText: string): string {
    return DMP.patch_toText(DMP.patch_make(oldText, newText));
  }

  /** Apply a dmp patch string to baseText. Throws if any hunk fails. */
  applyPatch(baseText: string, patchText: string): string {
    const { result, applied } = this.applyPatchWithResults(baseText, patchText);
    const failed = applied.filter((ok) => !ok).length;
    if (failed > 0) {
      throw new Error(`patch_apply: ${failed} of ${applied.length} hunks failed.`);
    }
    return result;
  }

  /** Apply a dmp patch string to baseText, returning result and per-hunk success flags. */
  applyPatchWithResults(
    baseText: string,
    patchText: string
  ): { result: string; applied: boolean[] } {
    const patches = DMP.patch_fromText(patchText);
    const [result, applied] = DMP.patch_apply(patches, baseText);
    return { result, applied: applied as boolean[] };
  }

  /** Sanity check: round-trip a patch and verify it reproduces newText. */
  verify(oldText: string, newText: string): boolean {
    try {
      return this.applyPatch(oldText, this.makePatch(oldText, newText)) === newText;
    } catch {
      return false;
    }
  }
}
```

---

### `src/vcs/BranchManager.ts`

```typescript
import { ObjectStore } from "./ObjectStore";

export class BranchManager {
  private store: ObjectStore;

  constructor(store: ObjectStore) {
    this.store = store;
  }

  /** Return the current branch name (e.g. "main"). */
  async currentBranch(): Promise<string> {
    const repo = await this.store.load();
    return repo.head.replace("refs/heads/", "");
  }

  /** Return all branch ref objects { name, tip }. */
  async listBranches(): Promise<Array<{ name: string; tip: string }>> {
    const repo = await this.store.load();
    return repo.refs;
  }

  /** Get the tip commit hash for a named branch. Throws if not found. */
  async getTip(branchName: string): Promise<string> {
    const repo = await this.store.load();
    const ref = repo.refs.find((r) => r.name === branchName);
    if (!ref) throw new Error(`Branch "${branchName}" not found.`);
    return ref.tip;
  }

  /** Advance the tip of a named branch to a new commit hash. */
  async advanceTip(branchName: string, hash: string): Promise<void> {
    const repo = await this.store.load();
    const ref = repo.refs.find((r) => r.name === branchName);
    if (!ref) throw new Error(`Branch "${branchName}" not found.`);
    ref.tip = hash;
    await this.store.save(repo);
  }

  /** Create a new branch pointing at the current branch's tip. */
  async createBranch(name: string): Promise<void> {
    const repo = await this.store.load();
    if (repo.refs.find((r) => r.name === name)) {
      throw new Error(`Branch "${name}" already exists.`);
    }
    const currentName = repo.head.replace("refs/heads/", "");
    const current = repo.refs.find((r) => r.name === currentName);
    const tip = current?.tip ?? "";
    repo.refs.push({ name, tip });
    await this.store.save(repo);
  }

  /** Switch HEAD to a different branch (does NOT modify document content). */
  async switchBranch(name: string): Promise<void> {
    const repo = await this.store.load();
    if (!repo.refs.find((r) => r.name === name)) {
      throw new Error(`Branch "${name}" not found.`);
    }
    repo.head = `refs/heads/${name}`;
    await this.store.save(repo);
  }
}
```

---

### `src/vcs/MergeManager.ts`

```typescript
import { DiffEngine, normaliseOoxml } from "./DiffEngine";

export class MergeManager {
  private diff: DiffEngine;

  constructor(diff: DiffEngine) {
    this.diff = diff;
  }

  /**
   * 3-way merge of OOXML strings.
   *
   * All three inputs are normalised before any diff/patch operation to prevent
   * hunk context mismatches from volatile OOXML attributes (w:rsidR, w:paraId etc.)
   *
   * Throws "CONFLICT: ..." if merge cannot be resolved automatically.
   * Caller must catch this and NOT modify the document on conflict.
   */
  threeWayMerge(baseRaw: string, oursRaw: string, theirsRaw: string): string {
    const base = normaliseOoxml(baseRaw);
    const ours = normaliseOoxml(oursRaw);
    const theirs = normaliseOoxml(theirsRaw);

    if (ours === base) return theirs;
    if (theirs === base) return ours;

    const oursPatch = this.diff.makePatch(base, ours);
    const theirsPatch = this.diff.makePatch(base, theirs);

    const { result: afterOurs, applied: oursApplied } = this.diff.applyPatchWithResults(
      base,
      oursPatch
    );

    const oursFailures = oursApplied.filter((ok) => !ok).length;
    if (oursFailures > 0) {
      throw new Error(
        `CONFLICT: ours patch failed to apply (${oursFailures} hunk(s)). Manual resolution required.`
      );
    }

    const { result: merged, applied: theirsApplied } = this.diff.applyPatchWithResults(
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
}
```

---

### `src/vcs/CommitManager.ts`

```typescript
/* global Word */

import LZString from "lz-string";
import { ObjectStore, CommitXml } from "./ObjectStore";
import { BranchManager } from "./BranchManager";
import { DiffEngine } from "./DiffEngine";
import { MergeManager } from "./MergeManager";

function sha6(): string {
  return (Date.now().toString(36) + Math.random().toString(36).slice(2, 6))
    .slice(0, 12)
    .padEnd(12, "0");
}

export interface CommitInfo {
  hash: string;
  parent: string;
  author: string;
  timestamp: string;
  message: string;
  encoding: CommitXml["encoding"];
}

/** Decompress a delta payload according to its encoding. Returns raw OOXML/patch text. */
function decodeDelta(delta: string, encoding: CommitXml["encoding"]): string {
  if (encoding === "snapshot-lz" || encoding === "dmp-lz") {
    const decompressed = LZString.decompressFromBase64(delta);
    if (decompressed === null || decompressed === "") {
      throw new Error("Failed to decompress commit data - repo may be corrupt.");
    }
    return decompressed;
  }
  return delta; // "snapshot" | "dmp" - already plain text
}

/** Compress a payload for storage. Returns { delta, encoding }. */
function encodeDelta(
  raw: string,
  kind: "snapshot" | "dmp"
): { delta: string; encoding: CommitXml["encoding"] } {
  const compressed = LZString.compressToBase64(raw);
  // Guard against pathological cases where compression doesn't help on tiny payloads -
  // still prefer compressed form for consistency (decoding is cheap either way),
  // but fall back to plain if compression somehow produced empty output.
  if (!compressed) {
    return { delta: raw, encoding: kind };
  }
  return { delta: compressed, encoding: kind === "snapshot" ? "snapshot-lz" : "dmp-lz" };
}

export class CommitManager {
  private store: ObjectStore;
  private bMgr: BranchManager;
  private diff: DiffEngine;

  constructor(store: ObjectStore, bMgr: BranchManager, diff: DiffEngine) {
    this.store = store;
    this.bMgr = bMgr;
    this.diff = diff;
  }

  /**
   * Initialise a new repo: create main branch + root commit (full snapshot).
   * Idempotent - safe to call again if already initialised.
   */
  async init(): Promise<void> {
    const repo = await this.store.load();
    if (repo.refs.length > 0) return;

    await Word.run(async (context) => {
      const body = context.document.body;
      const ooxmlResult = body.getOoxml();
      await context.sync();

      const hash: string = sha6();
      const { delta, encoding } = encodeDelta(ooxmlResult.value, "snapshot");
      const commit: CommitXml = {
        hash,
        parent: "00root",
        author: repo.author,
        timestamp: new Date().toISOString(),
        message: "Initial commit",
        delta,
        encoding,
      };

      repo.refs.push({ name: "main", tip: hash });
      repo.head = "refs/heads/main";
      repo.objects.push(commit);
      await this.store.save(repo);
    });
  }

  /**
   * Commit the current document state to the current branch.
   * Stores a dmp delta against the previous commit's OOXML.
   * Returns the new commit hash.
   */
  async commit(message: string): Promise<string> {
    const repo = await this.store.load();
    const branchName = repo.head.replace("refs/heads/", "");
    const ref = repo.refs.find((r) => r.name === branchName);
    if (!ref) throw new Error("No branch found. Did you Init?");

    return Word.run(async (context) => {
      const body = context.document.body;
      const ooxmlResult = body.getOoxml();
      await context.sync();

      const currentOoxml = ooxmlResult.value;
      const parentHash = ref.tip;
      const parentCommit = repo.objects.find((c) => c.hash === parentHash);

      let encoded: { delta: string; encoding: CommitXml["encoding"] };

      if (!parentCommit) {
        encoded = encodeDelta(currentOoxml, "snapshot");
      } else {
        const parentOoxml = await this.reconstruct(parentHash);
        const patch = this.diff.makePatch(parentOoxml, currentOoxml);
        encoded = encodeDelta(patch, "dmp");
      }

      const hash: string = sha6();
      const commit: CommitXml = {
        hash,
        parent: parentHash,
        author: repo.author,
        timestamp: new Date().toISOString(),
        message,
        delta: encoded.delta,
        encoding: encoded.encoding,
      };

      repo.objects.push(commit);
      ref.tip = hash;
      await this.store.save(repo);
      return hash;
    });
  }

  /**
   * Walk the parent chain and replay all deltas to reconstruct
   * the OOXML at a given commit.
   */
  async reconstruct(hash: string): Promise<string> {
    const repo = await this.store.load();

    const chain: CommitXml[] = [];
    let current: CommitXml | undefined = repo.objects.find((c) => c.hash === hash);
    while (current) {
      chain.unshift(current);
      if (current.parent === "00root") break;
      current = repo.objects.find((c) => c.hash === current!.parent);
    }

    if (chain.length === 0) throw new Error(`Commit ${hash} not found.`);
    if (chain[0].encoding !== "snapshot" && chain[0].encoding !== "snapshot-lz") {
      throw new Error("Root commit is not a snapshot - history may be corrupt.");
    }

    let ooxml = decodeDelta(chain[0].delta, chain[0].encoding);
    for (let i = 1; i < chain.length; i++) {
      const c = chain[i];
      const payload = decodeDelta(c.delta, c.encoding);
      if (c.encoding === "snapshot" || c.encoding === "snapshot-lz") {
        ooxml = payload;
      } else {
        ooxml = this.diff.applyPatch(ooxml, payload);
      }
    }
    return ooxml;
  }

  /**
   * Restore the document to a past commit with tracking ON.
   * Used for "Restore this version" - user reviews tracked changes.
   */
  async checkout(hash: string): Promise<void> {
    const ooxml = await this.reconstruct(hash);
    await Word.run(async (context) => {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;
      context.document.body.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    });
  }

  /**
   * Silently swap document content - tracking OFF.
   * Used for branch switching.
   */
  async checkoutSilent(hash: string): Promise<void> {
    const ooxml = await this.reconstruct(hash);
    await Word.run(async (context) => {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      context.document.body.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();
    });
  }

  /**
   * Return the commit log for the current branch, newest first.
   */
  async log(): Promise<CommitInfo[]> {
    const repo = await this.store.load();
    const branchName = repo.head.replace("refs/heads/", "");
    const ref = repo.refs.find((r) => r.name === branchName);
    if (!ref || !ref.tip) return [];

    const result: CommitInfo[] = [];
    let current: CommitXml | undefined = repo.objects.find((c) => c.hash === ref.tip);
    while (current) {
      result.push({
        hash: current.hash,
        parent: current.parent,
        author: current.author,
        timestamp: current.timestamp,
        message: current.message,
        encoding: current.encoding,
      });
      if (current.parent === "00root") break;
      current = repo.objects.find((c) => c.hash === current!.parent);
    }
    return result;
  }

  /**
   * Merge another branch into the current branch.
   * Common ancestor found via ancestor set walk.
   * Merge commit stored as snapshot to avoid ambiguous replay.
   * Throws "CONFLICT: ..." on failure - caller catches, shows toast, does NOT modify document.
   */
  async merge(theirBranch: string): Promise<string> {
    const repo = await this.store.load();
    const ourBranchName = repo.head.replace("refs/heads/", "");
    const ourRef = repo.refs.find((r) => r.name === ourBranchName);
    const theirRef = repo.refs.find((r) => r.name === theirBranch);

    if (!ourRef) throw new Error(`Current branch "${ourBranchName}" not found.`);
    if (!theirRef) throw new Error(`Branch "${theirBranch}" not found.`);

    const ourTip = ourRef.tip;
    const theirTip = theirRef.tip;

    if (ourTip === theirTip) throw new Error("Branches are already at the same commit.");

    const ourAncestors = new Set<string>();
    let c: CommitXml | undefined = repo.objects.find((x) => x.hash === ourTip);
    while (c) {
      ourAncestors.add(c.hash);
      if (c.parent === "00root") break;
      c = repo.objects.find((x) => x.hash === c!.parent);
    }

    let baseHash: string | null = null;
    let t: CommitXml | undefined = repo.objects.find((x) => x.hash === theirTip);
    while (t) {
      if (ourAncestors.has(t.hash)) {
        baseHash = t.hash;
        break;
      }
      if (t.parent === "00root") break;
      t = repo.objects.find((x) => x.hash === t!.parent);
    }

    if (!baseHash) throw new Error("No common ancestor found - cannot merge.");

    // Fast-forward
    if (baseHash === ourTip) {
      ourRef.tip = theirTip;
      await this.store.save(repo);
      await this.checkoutSilent(theirTip);
      return theirTip;
    }

    const baseOoxml = await this.reconstruct(baseHash);
    const oursOoxml = await this.reconstruct(ourTip);
    const theirsOoxml = await this.reconstruct(theirTip);

    const merger = new MergeManager(this.diff);
    const mergedOoxml = merger.threeWayMerge(baseOoxml, oursOoxml, theirsOoxml);

    const hash: string = sha6();
    const mergedEncoded = encodeDelta(mergedOoxml, "snapshot");
    const mergeCommit: CommitXml = {
      hash,
      parent: ourTip,
      author: repo.author,
      timestamp: new Date().toISOString(),
      message: `Merge branch '${theirBranch}' into ${ourBranchName}`,
      delta: mergedEncoded.delta,
      encoding: mergedEncoded.encoding,
    };

    repo.objects.push(mergeCommit);
    ourRef.tip = hash;
    await this.store.save(repo);

    await Word.run(async (context) => {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;
      context.document.body.insertOoxml(mergedOoxml, Word.InsertLocation.replace);
      await context.sync();
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    });

    return hash;
  }
}
```

---

### `src/vcs/GraphBuilder.ts` (Phase 7 - NEW)

```typescript
/**
 * GraphBuilder.ts - Commit DAG layout engine for WordGit branch graph.
 *
 * Given the raw repo data (commits + refs), computes x/y positions for
 * every commit node and edge list, ready to render as SVG.
 *
 * Layout algorithm:
 *   1. BFS from all branch tips following parent pointers (newest → oldest)
 *   2. Assign each tip a lane (col). Propagate lane to parents.
 *   3. row = BFS visit order index (0 = newest)
 *   4. Emit nodes and edges.
 */

export interface GraphCommit {
  hash: string;
  parent: string;
  message: string;
  timestamp: string;
  author: string;
}

export interface GraphRef {
  name: string;
  tip: string;
}

export interface GraphNode {
  hash: string;
  message: string;
  ts: string;
  author: string;
  col: number; // 0-indexed lane
  row: number; // 0 = newest
  refs: string[]; // branch names pointing here
  isRoot: boolean;
}

export interface GraphEdge {
  fromHash: string;
  toHash: string;
  fromCol: number;
  fromRow: number;
  toCol: number;
  toRow: number;
}

export interface GraphLayout {
  nodes: GraphNode[];
  edges: GraphEdge[];
  cols: number;
  rows: number;
}

export function buildGraph(commits: GraphCommit[], refs: GraphRef[]): GraphLayout {
  if (commits.length === 0) return { nodes: [], edges: [], cols: 1, rows: 0 };

  const byHash = new Map<string, GraphCommit>();
  for (const c of commits) byHash.set(c.hash, c);

  // Build children map
  const children = new Map<string, string[]>();
  for (const c of commits) {
    if (c.parent !== "00root") {
      const list = children.get(c.parent) ?? [];
      list.push(c.hash);
      children.set(c.parent, list);
    }
  }

  // BFS from all tips, newest-first
  const visited = new Set<string>();
  const topoOrder: string[] = [];
  const queue: string[] = refs.map((r) => r.tip).filter((h) => byHash.has(h));
  const enqueued = new Set<string>(queue);

  while (queue.length > 0) {
    const hash = queue.shift()!;
    if (visited.has(hash)) continue;
    visited.add(hash);
    topoOrder.push(hash);
    const c = byHash.get(hash);
    if (c && c.parent !== "00root" && !enqueued.has(c.parent) && byHash.has(c.parent)) {
      enqueued.add(c.parent);
      queue.push(c.parent);
    }
  }

  // Lane assignment: each ref tip gets its own lane
  const laneNames: string[] = [];
  const commitLane = new Map<string, number>();

  for (const ref of refs) {
    if (!byHash.has(ref.tip)) continue;
    if (!commitLane.has(ref.tip)) {
      commitLane.set(ref.tip, laneNames.length);
      laneNames.push(ref.name);
    }
  }

  // Propagate lanes down the parent chain
  for (const hash of topoOrder) {
    if (!commitLane.has(hash)) {
      const kids = children.get(hash) ?? [];
      const inheritFrom = kids.find((k) => commitLane.has(k));
      commitLane.set(hash, inheritFrom !== undefined ? commitLane.get(inheritFrom)! : 0);
    }
    const c = byHash.get(hash);
    if (c && c.parent !== "00root" && !commitLane.has(c.parent)) {
      commitLane.set(c.parent, commitLane.get(hash)!);
    }
  }

  // Map hash → ref names
  const hashRefs = new Map<string, string[]>();
  for (const ref of refs) {
    const list = hashRefs.get(ref.tip) ?? [];
    list.push(ref.name);
    hashRefs.set(ref.tip, list);
  }

  const nodes: GraphNode[] = topoOrder.map((hash, row) => {
    const c = byHash.get(hash)!;
    return {
      hash,
      message: c.message,
      ts: c.timestamp,
      author: c.author,
      col: commitLane.get(hash) ?? 0,
      row,
      refs: hashRefs.get(hash) ?? [],
      isRoot: c.parent === "00root",
    };
  });

  const nodeByHash = new Map<string, GraphNode>();
  for (const n of nodes) nodeByHash.set(n.hash, n);

  const edges: GraphEdge[] = [];
  for (const hash of topoOrder) {
    const c = byHash.get(hash)!;
    const from = nodeByHash.get(hash)!;
    if (c.parent === "00root") continue;
    const to = nodeByHash.get(c.parent);
    if (!to) continue;
    edges.push({
      fromHash: hash,
      toHash: c.parent,
      fromCol: from.col,
      fromRow: from.row,
      toCol: to.col,
      toRow: to.row,
    });
  }

  return { nodes, edges, cols: laneNames.length || 1, rows: nodes.length };
}
```
