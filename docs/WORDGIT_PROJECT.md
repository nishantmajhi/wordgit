# WordGit - project tracker

> **Superseded in places, 2026-09-19.** These are the design notes from the
> sessions that built WordGit, kept because the architecture, the UI decisions
> and the "why blame was dropped" reasoning are still the best account of why
> the project is shaped the way it is. They are not the current contract. The
> storage format has changed since -- merge commits now record a second parent,
> payloads are escaped text rather than CDATA, and unreadable data is refused
> rather than treated as absent -- so read `README.md` and `src/vcs/RepoFormat.ts`
> for what the code does now, and this for why.

> **PROJECT STATUS: COMPLETE.** Phases 0–7 are all done. WordGit is feature-complete.
>
> **How to use this file**: If you want to extend or modify WordGit later, upload this markdown + the diagrams to a new Claude session and start with: _"Resume the WordGit project. Here is the tracker and diagrams."_ Read the UI decisions and "Why blame was dropped" sections before changing anything - several design choices look arbitrary but are intentional and tested.
>
> **Instructions for Claude (if resumed)**: After completing any further work, update both `WORDGIT_PROJECT.md` and `WORDGIT_SETUP.md` (checklist items, UI decisions, pitfalls) and present both files for download before ending the session.

---

## What is WordGit?

A Microsoft Word Add-in that embeds a full Git-like version control system **inside the .docx file itself** - no external server, no separate repo. History, branches, commits, and diffs all live in Custom XML Parts within the docx ZIP. Word's own TrackedChanges API renders diffs natively.

---

## Architecture

### Three layers

| Layer      | What it is                      | Tech                  |
| ---------- | ------------------------------- | --------------------- |
| Word host  | Renders task pane, exposes APIs | Office JS sandbox     |
| VCS engine | Pure TypeScript VCS logic       | Runs inside task pane |
| Storage    | Custom XML Parts in .docx       | ZIP embedded XML      |

### Git primitives → Word mapping

| Git concept | WordGit equivalent                                              |
| ----------- | --------------------------------------------------------------- |
| Commit      | XML snapshot/delta stored as Custom XML                         |
| Diff        | diff-match-patch on raw OOXML string                            |
| Branch      | Named `<branch>` node with tip hash pointer                     |
| Checkout    | Apply delta chain → push as TrackedChanges                      |
| Merge       | 3-way dmp merge on normalised OOXML; conflicts surface as toast |
| `git log`   | Walk `parent` chain in Custom XML objects                       |
| `git blame` | Per-paragraph `w:rsid` tag + commit hash metadata (planned)     |

### Custom XML structure (inside .docx)

```
<wordgit:repo>
  <HEAD>refs/heads/main</HEAD>
  <config><author>Alice</author></config>
  <refs>
    <branch name="main"    tip="a3f9c1"/>
    <branch name="draft-2" tip="b7e2d4"/>
  </refs>
  <objects>
    <commit hash="a3f9c1" parent="00root">
      <author>Alice</author>
      <timestamp>2025-05-27T10:42:00Z</timestamp>
      <message>Draft intro section</message>
      <delta encoding="dmp">…diff-match-patch patch string…</delta>
    </commit>
  </objects>
</wordgit:repo>
```

---

## Tech stack (settled)

- **Add-in type**: Office Task Pane Add-in (Word)
- **UI**: Vanilla TypeScript + HTML/CSS inside the task pane
- **Styling**: `taskpane.css` for layout/components; button colours in `<style>` block in HTML (see UI pitfalls)
- **Diff algorithm**: `diff-match-patch` on raw `body.getOoxml()` strings
- **Storage**: Custom XML Parts (`CustomXmlParts` API)
- **Diff rendering**: `body.insertOoxml()` + `TrackedChanges` API
- **No server** - everything self-contained in the .docx
- **Scaffold**: Yeoman Office Add-in generator (`yo office`)
- **Clean export (Phase 7)**: `jszip` to strip Custom XML Parts from a raw .docx for "Alt+S download clean copy"
- **Storage compression (Phase 7)**: `lz-string` (`compressToBase64`/`decompressFromBase64`) - all commit deltas/snapshots compressed before being written to Custom XML

---

## UI spec (task pane - 3 tabs)

| Tab      | Emoji | What the user does                                                                          |
| -------- | ----- | ------------------------------------------------------------------------------------------- |
| Commit   | 💾    | Type optional message → Commit or Enter; Init repo when not yet done                        |
| Log      | 📜    | Refresh → expandable commit cards (accordion) → Restore or Show diff; Export JSON at bottom |
| Branches | 🌿    | Refresh → list of branches → Switch / Merge in; Create branch with Enter support            |

### UI decisions (DO NOT change without good reason)

#### Tabs

- Emojis in tab labels, no emojis in buttons (except Init repo ⚙️ which is a special action)
- Active tab: blue underline + bold text + `rgba(0,120,212,0.08)` tinted background
- Tab font size: 12px

#### Button hierarchy - three roles, strict visual weight

| Class           | Appearance                                                            | Used for                                                   |
| --------------- | --------------------------------------------------------------------- | ---------------------------------------------------------- |
| `btn-primary`   | Solid blue `#0078d4`, white text                                      | ONE main action per panel (Commit, Create branch, Restore) |
| `btn-refresh`   | White bg, blue `#0078d4` outline+text; inverts to solid blue on hover | Refresh buttons - frequent but not the main action         |
| `btn-secondary` | Gray `#eeeeee` bg, dark text, gray border                             | Low-priority actions (Init repo, Export JSON, Cancel)      |
| `btn-danger`    | Solid red `#c4262e`, white text                                       | Destructive/irreversible actions (Show diff, Merge in)     |

**Why Show diff and Merge are red**: both replace the document body - Show diff replaces content with a diff view, Merge modifies the document. Red signals "this changes your document."

#### Log tab

- Commit cards: accordion - only one card expanded at a time; clicking open card closes it
- Expanded card: no blue outline/glow - the open state itself (detail visible, chevron rotated) is sufficient affordance
- Card actions: `↩ Restore` (blue primary), `⟷ Show diff` (red danger)
- Confirm banner appears inside the card, not as a modal
- Export history as JSON: pinned to bottom of Log panel via `panel-footer` (always visible, even when few commits)

#### Branches tab

- Branch list rendered as a bordered card (`#branch-list`) with rows
- Active branch row: light blue background `rgba(0,120,212,0.10)` + `"current"` pill badge
- Inactive branch rows: `⇄ Switch` (blue primary), `⤢ Merge in` (red danger)
- No empty-state hint text - silent when no branches/no repo
- Panel order (top → bottom): Create branch section → divider → branch list → commit graph

#### Branch graph (Phase 7)

- Rendered inline as SVG inside `#branch-graph` div below the branch list, after Refresh
- `#branch-graph` has no border/bg by default; JS adds `.has-graph` class once SVG is injected
- "No commits yet." shown as plain `.wg-hint` text with no container styling when repo is empty
- Layout model: each node is a horizontal row - **[pill] [circle] [hash / message / timestamp]** - all vertically centred on the same Y coordinate
- Pills sit to the LEFT of the circle, right-aligned to it, vertically centred
- Text (hash, message, timestamp) sits to the RIGHT of the circle, 3 lines centred on Y
- Lane width is **dynamic**: computed per-lane from the widest branch-name pill in that lane + `PILL_CIRCLE_GAP + R + LANE_SAFETY_PAD(40px)` - guarantees no text bleeds into adjacent lane's circle
- Circle filled for branch tip nodes, outlined for history nodes
- Up to 6 lane colours cycling: blue, green, purple, orange, teal, red
- Dark mode handled at render time via `body.classList.contains("wg-dark")`

#### Keyboard shortcuts (Phase 7)

- `Alt+C` - Commit checkpoint: if repo isn't initialised, inits it first (silently, no toast for the init step), then commits with `YYYY-MM-DD HH:MM:SS` as the message. Does NOT read the commit-message input box.
- `Alt+L` - Jump to Log tab and refresh the commit list
- `Alt+B` - Jump to Branches tab and refresh the branch list + graph
- `Alt+S` - Download a clean copy of the document as `<original-basename>_clean.docx` with all WordGit Custom XML Parts (history, branches, refs) stripped out - useful for sharing a "final" doc without exposing version history
- All shortcuts use `e.preventDefault()` and only fire on `Alt` alone (not `Ctrl+Alt` or `Alt+Shift` combos clashing with OS/Word shortcuts)
- Listener is attached to `document`, so it works regardless of which input/tab has focus

#### Commit tab

- On load: text input + Commit button **disabled** until repo is initialised
- After Init repo: input + Commit enabled; Init repo button disabled
- `checkRepoExists()` runs on `Office.onReady` to handle reopening an already-initialised doc

#### Inputs

- Slightly off-white bg on rest, white on focus
- Blue border + glow ring on focus
- Dimmed + `not-allowed` cursor when disabled
- Enter key submits: commit message → commit, branch name → create branch

#### Tracked changes UX

| Operation           | changeTrackingMode          | Rationale                                    |
| ------------------- | --------------------------- | -------------------------------------------- |
| Branch switch       | `off`                       | Silent swap - no review needed               |
| Restore old version | `trackMineOnly`             | User reviews what changed                    |
| Merge (clean)       | `trackMineOnly`             | User reviews what merge introduced           |
| Merge (conflict)    | n/a - document NOT modified | Surface conflict as warning toast            |
| Show diff           | `off` then `trackMineOnly`  | Silently place From baseline, then insert To |

#### Toast notifications

- Four types: success (green), error (red), warning (amber), info (blue)
- Auto-dismiss after 4s; × to dismiss early
- Slide in/out animation
- Positioned fixed at bottom of task pane

---

## Progress tracker

### Phase 0 - Design ✅ DONE

- [x] Architecture diagrams (4 SVGs)
- [x] Project tracker markdown

### Phase 1 - Scaffold ✅ DONE

- [x] `yo office` → Task Pane Add-in, TypeScript
- [x] Sideloading in Word desktop confirmed
- [x] Custom XML Part hello world + round-trip verified

### Phase 2 - Core VCS engine ✅ DONE

- [x] `ObjectStore`, `CommitManager`, `BranchManager`, `DiffEngine`, `MergeManager`
- [x] Root commit (snapshot), subsequent commits (dmp deltas)

### Phase 3 - Checkout & log view ✅ DONE

- [x] `reconstruct(hash)`, `checkout(hash)`, `checkoutSilent(hash)`
- [x] Log panel with expandable commit cards

### Phase 4 - Branches & merge ✅ DONE

- [x] Branch switch, create, list
- [x] 3-way merge with normalised OOXML
- [x] Merge conflict surfaced as toast

### Phase 5 - UX polish ✅ DONE

- [x] Toast notification system
- [x] Office theme support via CSS custom properties
- [x] `checkoutSilent` for branch switches

### Phase 6 - Polish ✅ DONE

- [x] Export history as JSON (pinned to bottom of Log tab)
- [x] Diff moved into log cards - "Show diff" diffs commit vs its parent directly (no separate Diff tab)
- [x] Separation of concerns: layout/components in `taskpane.css`; button colours in HTML `<style>` block
- [x] Init state: Commit input + button disabled until repo initialised; Init repo disables after use
- [x] Accordion log cards - one open at a time
- [x] Active branch highlight + "current" badge
- [x] Active tab tinted background
- [x] Enter key on branch name input → create branch
- [x] Dark mode detection via luminance check; `body.wg-dark` class toggled; dark CSS vars applied

### Phase 7 - Branch graph & extras ✅ DONE

- [x] Branch graph SVG visualization (`GraphBuilder.ts` + `renderBranchGraph()`)
- [x] Branch graph layout: horizontal flex model - pill | circle | text all on same Y
- [x] Dynamic lane width computed from actual branch name pill widths + safety pad (no fixed COL_W; no overlap possible)
- [x] Branches panel reordered: Create branch at top, then branch list, then graph below
- [x] Removed "Commit" section label from Commit tab (tab name is sufficient)
- [x] Graph container plain by default; `.has-graph` class added by JS only when SVG is ready
- [x] "No branches yet" hint removed - silent empty state
- [x] Keyboard shortcuts:
  - Alt+C → commit with timestamp message; auto-inits repo if not yet initialised
  - Alt+L → switch to Log tab + refresh
  - Alt+B → switch to Branches tab + refresh
  - Alt+S → download `<name>_clean.docx` with all WordGit Custom XML Parts stripped (via JSZip)
- [x] LZ-string compression of all stored deltas/snapshots (`encodeDelta`/`decodeDelta` in CommitManager; encoding values `snapshot-lz` / `dmp-lz`; backward-compatible with legacy `snapshot`/`dmp`)
- [x] **Per-paragraph blame via `w:rsid` - investigated, NOT implemented (see "Why blame was dropped" below)**

---

## 🎉 PROJECT STATUS: COMPLETE

All planned phases (0–7) are done. WordGit is a fully functional in-document Git-like VCS for Word: init, commit, log with diffs, branches, merges, a commit DAG graph, keyboard shortcuts, compressed storage, and clean export. No further phases are planned. Future work would be opportunistic polish, not a structured phase.

### Why blame was dropped

Per-paragraph blame via `w:rsid` was investigated but not built. The reason: `w:rsid*` attributes (rsidR, rsidRPr, rsidP, rsidRDefault, etc.) are Word _session_ IDs - Word regenerates/reassigns them every time `body.insertOoxml(..., InsertLocation.replace)` runs, which is exactly what WordGit does on every checkout, branch switch, and merge. So any rsid → commit mapping recorded at commit time would be silently invalidated the next time the user switches branches or restores a version, making the blame data wrong (and impossible to detect as wrong) without warning.

A workable alternative exists - diff each paragraph's text content between consecutive commits (keyed by `w:paraId` with a position-based fallback) and report "last changed in commit X" per paragraph, independent of Word's rsid churn. This would give the same user-facing result as git blame. It was deferred rather than built in this session; if revisited, implement it as a new `BlameEngine.ts` that walks the commit chain via `CommitManager.reconstruct()` + `CommitManager.log()`, diffs paragraph-by-paragraph (not full-document dmp), and renders results as a 4th tab or as an overlay in the Log tab.

---

## Known hard problems

1. **Large docs** - root commit snapshot can be large. Mitigated: all snapshots/deltas now LZ-string compressed (`snapshot-lz`/`dmp-lz`); further mitigation would be periodic re-snapshotting to cap delta-chain length.
2. **Binary assets** - images live in `word/media/` ZIP; Office JS can't manipulate the ZIP directly.
3. **Merge conflicts** - Word TrackedChanges has no "conflict marker" concept. Plan: CONFLICT-A / CONFLICT-B author names.
4. **Performance on checkout** - many deltas = slow replay. Mitigation: periodic pack snapshots.
5. **Alt+S clean export reads last-saved file** - `getFileAsync(Office.FileType.Compressed)` reads the document as saved on disk, NOT live in-memory edits. If the user has unsaved changes, the downloaded "clean" copy won't include them. Consider prompting to save first, or calling `document.save()` before export if Office JS API set supports it in the target Word version.
6. **No per-paragraph blame** - Word regenerates `w:rsid*` attributes on every `insertOoxml` replace (checkout/switch/merge all do this), so rsid-based blame would silently go stale. Not implemented. A paragraph-text-diff approach (keyed on `w:paraId`) would work but wasn't built - see Phase 7 notes.

---

## Files in this project

| File                 | What it is                                   |
| -------------------- | -------------------------------------------- |
| `WORDGIT_PROJECT.md` | This tracker                                 |
| `WORDGIT_SETUP.md`   | Reproduction guide, pitfalls, file structure |
| Diagrams (4 SVGs)    | Re-upload with tracker to resume sessions    |

---

## If you want to pick this back up later

> "Resume the WordGit project. Tracker and diagrams attached. The project is feature-complete (Phases 0–7 done) - read the UI decisions section and the 'Why blame was dropped' note before changing anything.
>
> The UI is fully polished - read the UI decisions section carefully before touching taskpane.html, taskpane.css, or taskpane.ts. The button colour system, accordion log cards, init state management, dark mode detection, branch graph layout, and keyboard shortcuts are all intentional and tested.
>
> Branch graph layout: horizontal flex model - pill LEFT of circle | circle | text RIGHT of circle - all on same Y. Lane widths are dynamic (computed from branch name pill widths). DO NOT revert to fixed COL_W or above/below pill placement.
>
> Keyboard shortcuts (Alt+C/L/B/S) are wired globally on `document` in Office.onReady - don't duplicate listeners per-element.
>
> Storage: all commit deltas/snapshots are LZ-string compressed (`snapshot-lz`/`dmp-lz` encodings in CommitManager's encodeDelta/decodeDelta). Legacy uncompressed `snapshot`/`dmp` commits are still readable - don't break that backward compatibility.
>
> If asked for blame/history-per-paragraph: do NOT use w:rsid (Word regenerates it on every insertOoxml replace). Use a paragraph-text-diff approach keyed on w:paraId instead, as described in the Phase 7 notes.
>
> All source files are in WORDGIT_SETUP.md."
