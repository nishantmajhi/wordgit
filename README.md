# WordGit

A Microsoft Word add-in that keeps a version history inside the document.

Commits, branches, diffs and merges, with the whole repository stored in the
`.docx` file's own Custom XML Parts. There is no server and no account. If you
email someone the file, you email them the history; if they open it in Word with
WordGit installed, they can read it, branch from it and commit to it.

That is the whole idea, and it is also the main thing to understand before
relying on it. See [What will lose your history](#what-will-lose-your-history).

---

## What it does

**Commit.** Press Commit and the current state of the document is stored. The
first commit holds a complete snapshot of the document's OOXML; every commit
after it holds a diff-match-patch delta against its parent, compressed. A
commit is reconstructed by replaying from the most recent snapshot at or below
it.

**Restore.** Restore a past commit and its content is inserted with Word's
change tracking on, so you see it as tracked changes and can reject it. Nothing
is lost until you accept.

**Show diff.** Between a commit and its parent, rendered the same way — the
parent is laid down silently, then the child is inserted with tracking on, so
Word itself draws the difference. Note that this replaces the document body, so
**commit your work before using it.**

**Branch and merge.** Branches are named pointers to a commit, as in git.
Merging is a three-way merge over the documents' OOXML. It either combines both
sides or refuses; it never guesses. See [Merging](#merging).

**Clean export.** `Alt+S` downloads a copy of the `.docx` with every WordGit
part removed, along with the package entries that referred to them — for when
you want to hand the document on without the history.

**Export history.** The Log tab writes the commit list to a JSON file. This is
metadata only: hashes, authors, timestamps and messages, not the content.

---

## Keyboard shortcuts

| Key     | Action                                                     |
| ------- | ---------------------------------------------------------- |
| `Alt+C` | Commit with the current timestamp as the message           |
| `Alt+L` | Log tab, refreshed                                         |
| `Alt+B` | Branches tab, refreshed                                    |
| `Alt+S` | Download a clean copy with the WordGit data stripped        |

`Alt+C` initialises the repository first if the document does not have one.

---

## What will lose your history

Custom XML Parts survive a `.docx` round trip through Word. They do not survive
much else. The history is **silently discarded**, with no error and nothing to
undo, by:

- saving as `.doc`, `.rtf`, `.odt` or anything else that is not `.docx`;
- exporting to PDF;
- uploading to Google Docs, or most other web editors, and downloading again;
- most "convert this document" services;
- some mail systems that sanitise attachments.

Word itself also strips Custom XML Parts when the document is saved in
compatibility mode. If the history matters, keep the `.docx` and keep a backup
of the file, not of the history.

WordGit will tell you when the data it finds is damaged rather than missing, and
will refuse to write anything until you have dealt with it — but it cannot
recover a part that is no longer there.

---

## Storage format

One Custom XML Part per document, in the namespace
`http://schemas.wordgit.local/vcs/1`:

```xml
<wordgit:repo xmlns:wordgit="http://schemas.wordgit.local/vcs/1">
  <HEAD>refs/heads/main</HEAD>
  <config><author>Ada</author></config>
  <refs>
    <branch name="main" tip="mh3k91af2c7e"/>
    <branch name="draft-2" tip="mh3k95bq81d4"/>
  </refs>
  <objects>
    <commit hash="mh3k91af2c7e" parent="00root">
      <author>Ada</author>
      <timestamp>2026-01-14T10:42:00.000Z</timestamp>
      <message>Draft intro section</message>
      <delta encoding="snapshot-lz">…</delta>
    </commit>
    <commit hash="mh3ka2c7f019" parent="mh3k91af2c7e" parent2="mh3k95bq81d4">
      …
    </commit>
  </objects>
</wordgit:repo>
```

- `parent="00root"` marks the first commit. There is no root object.
- `parent2` is present only on a merge commit, and names the branch that was
  merged in.
- `encoding` says what the payload is: `snapshot` (complete OOXML), `dmp` (a
  diff-match-patch patch against the parent), or either of those with `-lz`,
  meaning compressed with `LZString.compressToBase64`. Compression is skipped
  when it would make the payload longer, which it does on short patches.
- Payload text is escaped, not wrapped in CDATA. It has to be: an XML parser
  normalises line endings before anything else sees the text, so a literal
  carriage return becomes a line feed and is gone, and a CDATA section cannot
  carry the `&#13;` that would survive. Repositories written with CDATA, by
  earlier versions, still read correctly.

The rules for changing this format are in `src/vcs/RepoFormat.ts`. The short
version: the data is inside users' files and cannot be migrated centrally, so
anything an earlier version wrote must still be readable, and anything written
now must still be readable later. An unrecognised payload encoding is refused by
name rather than guessed at.

### Damaged data

`parseRepoXml` refuses, rather than interpreting, a part that is not valid XML,
is not a WordGit repository, has a commit with no hash or no payload, names a
payload encoding it does not know, or lists the same branch or commit twice.
When that happens the task pane disables every action that writes, and
`ObjectStore` refuses to save. A repository that cannot be read is not the same
thing as a document that has never had one, and the difference is your history.

---

## Merging

A merge compares three documents: the common ancestor, your branch and theirs.

Before anything is combined, the two sides are checked for **overlap**: the
spans of the ancestor that each side rewrote, including the few characters of
context diff-match-patch keeps around each change. If those spans touch, the
merge is refused as a conflict and neither the document nor the branches are
touched.

This is not the same as asking whether both patch sets applied. `patch_apply`
is deliberately fuzzy — it searches near where it expects a change to go — so
when two people rewrite the same sentence it will usually find somewhere
plausible to put the second one and report success. What comes out is a
sentence neither of them wrote.

As a second guard, both patch sets are applied in each order and the results
must agree. Disjoint changes commute; if they do not, something landed in the
wrong place whatever it reported.

Before diffing, attributes Word regenerates on every save are removed:
`w:rsid*`, and `paraId`/`textId` under any prefix. `w:id` is **not** removed —
it is the identity of a bookmark, a comment anchor, a tracked revision and a
footnote reference, and removing it unpairs every one of them.

A merge commit is stored as a full snapshot, and records both parents.

---

## Setup

### Prerequisites

- Node 20 or newer
- Microsoft Word, desktop, Windows or macOS
- A Microsoft 365 account signed in to Word

### Install and run

```bash
npm install
npm start            # builds, starts the dev server, sideloads into Word
```

`npm start` runs `office-addin-debugging`, which generates and installs a local
development certificate the first time. Word will open with the add-in in the
Home tab.

```bash
npm stop             # unregister the add-in and stop the dev server
npm run build        # production bundle into dist/
npm test             # the test suite
```

### Sideloading by hand

**Windows.** File → Options → Trust Center → Trust Center Settings → Trusted
Add-in Catalogs, and add the folder containing `manifest.xml`.

**macOS.** Insert → Add-ins → My Add-ins → Upload My Add-in, and select
`manifest.xml`.

---

## Hosting

`manifest.xml` points at `https://localhost:3000`, which is correct for
development and wrong for anything else. To install the add-in somewhere it
keeps working without a dev server running: build it, host `dist/` over HTTPS,
and replace every `localhost:3000` URL in the manifest with the hosted base.

There are four such URLs: `SourceLocation` for the task pane, `IconUrl`,
`HighResolutionIconUrl`, the three `bt:Image` resources and the two `bt:Url`
resources. Keep the `<Id>` GUID unchanged — changing it registers a second,
separate add-in rather than updating this one.

Any HTTPS static host works: GitHub Pages, Cloudflare Pages, Netlify, Azure
Static Web Apps, or an internal server. An internal server must serve valid
HTTPS and send `Access-Control-Allow-Origin: *`.

For an organisation, upload the edited manifest at
[admin.microsoft.com](https://admin.microsoft.com) under Settings → Integrated
apps → Upload custom apps, choose the scope, and deploy. Rollout takes up to 24
hours, after which it appears under Insert → Add-ins → Admin Managed.

---

## Layout

```
manifest.xml              What Word is told about the add-in
src/
  taskpane/
    taskpane.html         Markup and button styles
    taskpane.css          Everything else visual
    taskpane.ts           Handlers, and the wiring between them and the VCS
    render.ts             The markup built at runtime, escaped
    graphSvg.ts           The commit graph, drawn
    download.ts           Handing a file to the user
  commands/               The FunctionFile the manifest declares; no commands
  vcs/
    RepoFormat.ts         The storage contract: serialise, parse, validate
    ObjectStore.ts        Reading and writing the document's Custom XML Parts
    History.ts            Graph walks: replay, ancestry, merge base, log
    CommitManager.ts      init, commit, checkout, merge
    BranchManager.ts      Branch names and pointers
    DiffEngine.ts         diff-match-patch, and OOXML normalisation
    MergeManager.ts       Three-way merge and conflict detection
    GraphBuilder.ts       Commit DAG layout
    DocxCleaner.ts        Stripping WordGit out of a .docx
tests/                    See below
docs/                     Design notes and architecture diagrams
```

The split that matters is that nothing in `src/vcs/` calls Office.js directly.
`ObjectStore` takes a `CustomXmlPartHost` and `CommitManager` takes a
`DocumentIO`; the Office.js implementations of both live alongside them and are
about forty lines between them. Everything else is decisions over strings.

---

## Tests

```bash
npm test
```

`pretest` compiles `src/` to `tests/build/` as CommonJS with `tsc`, so the
suite exercises the real source rather than a copy of it. There is no test
framework: `node --test`, plain assertions, no dependencies beyond what the
add-in already has.

| File                 | What it covers                                               |
| -------------------- | ------------------------------------------------------------ |
| `format.test.js`     | Round trips, escaping, and every reason to refuse a part      |
| `store.test.js`      | When it is safe to write to the document, and when it is not  |
| `history.test.js`    | Replay, ancestry, merge base, log ordering, payload encoding  |
| `vcs.test.js`        | init, commit, branch, switch, merge, restore, end to end      |
| `merge.test.js`      | What merges, what conflicts, and OOXML normalisation          |
| `graph.test.js`      | Commit DAG layout and the SVG it becomes                      |
| `render.test.js`     | That document content cannot change the shape of the markup   |
| `docx.test.js`       | Stripping WordGit out of real zip packages                    |
| `download.test.js`   | Anchor attachment and object-URL lifetime                     |
| `manifest.test.js`   | That no scaffolding placeholder has come back                 |
| `original.test.js`   | Each fix, run against the original code to show it was needed |

Two pieces of scaffolding are worth knowing about. `tests/xml.js` is a DOMParser
with the three behaviours this format depends on: line-ending normalisation
before anything else, a `<parsererror>` document instead of an exception on
malformed input, and `textContent` that folds CDATA in. `tests/harness.js` holds
an in-memory Custom XML Part store and an in-memory document, which is what
makes a whole init-commit-branch-merge sequence runnable without Word.

`tests/original.js` is a transcription of the code each fix replaced, and
`original.test.js` reads the baseline commit out of git and asserts the
transcription still matches it — so a transcription that has drifted fails
loudly rather than quietly asserting nothing.

---

## Limitations

**Images and other binary content are not versioned.** They live in
`word/media/` inside the zip, and Office.js offers no way to reach it from a
task pane. Only the text content is in a commit, so an image added on one branch
will not survive a switch away and back.

**Every commit rewrites the whole repository.** The repository is one XML part,
so commit number two hundred rewrites all its predecessors. The delta encodings
and the compression keep the size down; the rewrite cost is still linear in
history length on every commit. One part per commit plus a small index would fix
it, and would be a format change with a migration.

**Clean export reads the last saved state.** `Alt+S` asks Office for the
document as it is on disk. Save first if you have unsaved edits.

**Long histories are slow to check out.** Reconstruction replays from the most
recent snapshot, and only the first commit and merge commits are snapshots.

**No per-paragraph blame.** Word regenerates `w:rsid*` on every `insertOoxml`,
which WordGit does for every checkout, branch switch and merge, so anything
keyed on them would go stale with no way to detect it.

**The author name is the document's, not the user's.** Word's JavaScript API
does not expose the signed-in user's name to a task pane. WordGit asks for the
document's `lastAuthor` property on each commit and falls back to the literal
`Author` where the host will not say.

---

## Built with

TypeScript, bundled with webpack 5. `diff-match-patch` for the deltas and the
merge, `lz-string` for compression, `jszip` for the clean export, and Office.js
for `Word.run`, `CustomXmlParts` and change tracking. The task pane is plain
CSS, with dark mode chosen from the luminance of the Office theme's background.

---

## Licence

Free for personal, non-commercial use. You may use, study and modify this
software for yourself, but you may not use it commercially, and you may not
share or publish it, in its original form or modified. See [LICENSE](LICENSE)
for the full terms.

`src/commands/commands.html` comes from Microsoft's Office Add-in template
and keeps its MIT notice.
