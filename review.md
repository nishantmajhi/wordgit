# Review and remediation — wordgit

First reviewed 2026-09-04. Remediated 2026-09-19 on branch `cleanup/2026-09-19`.
195 tests, all passing. Nothing pushed.

---

## Corrections to the first review

The first review was accurate about everything it looked at. It looked at
roughly half the project, and the half it did not look at held the worst bugs.
Two of its conclusions need amending.

**"The XML escaping is correct, including the `]]>` splitting trick in
`escapeCdata` that most hand-rolled serialisers get wrong."** The splitting
trick is correct and the escaping of `&`, `<`, `>` and `"` is correct. But
choosing CDATA at all is the bug: an XML parser normalises line endings before
anything else sees the text, and CDATA cannot carry the `&#13;` that would
survive it, so **a payload containing a carriage return did not survive a save
and load**. `tests/original.test.js` demonstrates it: `<w:p>one</w:p>\r\n…`
comes back as `<w:p>one</w:p>\n…`. A commit whose payload changes underneath it
means the next delta is computed against a base that no longer matches, and the
failure surfaces later as a patch that will not apply. Getting the escaping
right inside a construct that cannot represent the data is not the same as
getting the serialisation right.

**"No secrets, no debug output, no analytics, no committed build output."**
True, and it is the sentence that closes the "Cleanup" section — but the review
never looked at the task pane's markup building, where **every string the
document carries was written into `innerHTML` unescaped**. Commit messages,
branch names, the author line and error text, in four separate builders. A
WordGit repository travels inside the `.docx`, so those strings are written by
whoever edited the file before you, and the task pane is a webview with your
document open through Office.js. Opening a document someone else had committed
to was enough to run their script against it. The review checked the *storage*
escaping carefully and did not think to check the *display* escaping.

The first review's other finding worth flagging as understated rather than
wrong: it asked, of the 52-line `MergeManager`, to "confirm what merge actually
does when two branches change the same paragraph" and noted that "a merge that
silently picks one side is much worse than one that refuses". It does neither.
See below; it is the headline bug.

---

## What was wrong

### P1 — Damaged history presented as no history, and was then destroyed

`xmlToRepo` parsed with `DOMParser` and read fields with `?? default`. A browser
does not throw on malformed XML; it returns a document whose root is a
`<parsererror>`. So **any damage to the stored XML — a truncating converter, a
failed write, a hand edit — produced an empty repository**: no branches, no
commits, default HEAD.

`checkRepoExists()` used `repo.refs.length > 0` to decide whether the document
had ever been initialised. Zero refs enabled the Init button. `handleInit` then
called `ObjectStore.save()`, which **deleted the existing Custom XML Part** and
added a new one. One click, and a history that was merely unreadable was gone.
`Alt+C` reached the same code with no click at all: it initialised first if it
saw no refs.

Fixed by separating the two cases at the bottom. `parseRepoXml` throws
`RepoFormatError` for anything it cannot interpret — unparseable XML, a root
element that is not a WordGit repository, a commit with no hash or no payload,
an unrecognised payload encoding, a duplicate branch or commit. `ObjectStore`
remembers that the load failed and `save()` refuses to write over it. The task
pane disables every writing action and says so for twelve seconds rather than
four. What is *not* rejected is as deliberate as what is: a missing HEAD still
falls back to `main`, a branch with an empty tip is still allowed, and a
repository written in the old CDATA format still reads, because refusing to open
a document a user can still read is its own kind of data loss.

### P1 — Two branches editing the same sentence merged silently into a third

`MergeManager.threeWayMerge` applied ours to the base, then theirs to the
result, and treated `patch_apply` reporting success as proof the merge was
clean. `patch_apply` is deliberately fuzzy: it searches near where it expects a
hunk and places it somewhere plausible. Given a base sentence "The quick brown
fox jumps over the lazy dog.", one branch changing it to "The quick brown fox
leaps over the idle dog." and the other to "A slow grey badger ambles past the
lazy dog.", the original produced, with no conflict and no warning:

> A slow grey badger ambles past the idle dog.

Their subject, our object, a sentence neither author wrote, inserted into the
document as tracked changes under a toast that said the merge had succeeded.

The fix is to ask a different question. Before anything is applied, the spans of
the base that each side rewrote — including the context diff-match-patch keeps
around each hunk — are compared, and overlapping spans are a conflict. Then both
patch sets are applied in each order and the results must agree, because
disjoint edits commute and edits that do not commute did not land where they
belonged. `tests/merge.test.js` covers both directions: rewrites of the same
sentence, of the same paragraph, an edit against a deletion and two edits inside
the shared context window all conflict; edits to different paragraphs, an
insertion against a distant edit, and identical edits on both sides all merge.

### P1 — Commit messages and branch names ran as script in the task pane

Four sinks, all `innerHTML`, all fed from the document:

| Where | Field |
| --- | --- |
| `handleRefreshLog` | commit message, author, hash, parent hash |
| `handleRefreshBranches` | branch name, in text and in two `data-` attributes |
| `showToast` | any error message, which quotes branch names back |
| `renderBranchGraph` | the commit hash, and the graph's own error message |

The graph escaped the message and the branch pill through `gcEsc` and missed
the hash. The hash is truncated to six characters before it is written, so it
cannot carry a payload — but `</text` is exactly six characters, which is enough
to unbalance the element and corrupt the graph below it. That one is a
robustness bug, not an injection; the other three are injections.

The builders are now in `src/taskpane/render.ts` and `src/taskpane/graphSvg.ts`,
escaped, and tested with a shape check: the same markup is rendered twice, once
with a hostile string and once with an inert one, and the two must have the same
elements with the same attributes in the same order. Content that cannot change
the shape of the markup cannot inject anything, so there is no payload list to
maintain. `tests/render.test.js` also asserts the check **fails** against the
original builders, so a shape function that returned a constant could not make
the suite pass.

A related bug fell out of the same code: the handlers were wired by building a
selector from the value — `li.querySelector('[data-switch="' + name + '"]')`.
A branch named `say "hi"` throws a `SyntaxError` out of `querySelector` and
leaves the whole branch list inert. They are now found by attribute name and the
value is read off the element.

### P1 — The add-in identified itself as published by Contoso

As the first review said. `ProviderName`, the description users see, the support
URL, the task pane id, the ribbon labels and the get-started copy were all
generator placeholders, and `https://www.contoso.com` was declared as a trusted
app domain — widening what the add-in's webview may navigate to, for nothing.
The `AppDomains` element is gone entirely; the add-in loads only its own
`SourceLocation`, which is trusted implicitly. `package.json` still called the
project `office-addin-taskpane` and pointed at Microsoft's template repository.
`tests/manifest.test.js` fails if any of it comes back, and also asserts every
URL in the manifest is https and points at either localhost or this project.

### P2 — The merge stripped `w:id` out of the document, permanently

`VOLATILE_ATTR_RE` removed `w:rsid*`, `w:paraId`, `w:textId` and **`w:id`** as
"attributes Word regenerates on every save". Three of those four claims are
wrong.

`w:id` is not volatile. It is the identity of a bookmark, a comment anchor, a
tracked insertion or deletion, and a footnote reference. Stripping it unpairs
every `w:bookmarkStart` from its `w:bookmarkEnd` and detaches every comment from
the text it was about — and because the merge result was stored as the new
snapshot *and* pushed into the document, the loss was permanent from the first
merge onwards.

`w:paraId` and `w:textId` do not exist. In real OOXML they are `w14:paraId` and
`w14:textId`, in the 2010 extension namespace, and the pattern was anchored to
the `w:` prefix, so neither alternative ever matched anything. The
normalisation the code documented was, in practice, rsids only.

Both fixed, and `tests/merge.test.js` asserts a merged document keeps its
bookmark pairing.

### P2 — A merge recorded only one parent

`mergeCommit.parent` was set to our tip and the merged branch was not recorded
at all. Three consequences, all of them real rather than cosmetic:

- The log walked the first-parent chain, so **none of the work that had just
  been merged in appeared in the history** — although it was in the document.
- The graph had no way to draw a merge edge; `GraphCommit` had no field for it.
- Merging the same branch twice replayed the whole thing, because the search for
  a common ancestor could not see past our own first-parent chain.

The format now carries an optional `parent2` on a commit, written only when
there is one, and ignored by any reader that predates it. `ancestors`,
`mergeBase`, `logFrom` and `buildGraph` all follow both parents, and merging an
already-merged branch is refused by name.

### P2 — The log could print a commit above its own parent

Both the log and the graph ordered by traversal order, which happened to work on
a linear history and interleaved two branches of different lengths by depth
rather than by date. Both now use a topological order with the timestamp as the
tie-break: of the commits whose children have all been placed, the newest goes
next. That gives date order where the graph allows it and graph order where it
does not — which matters, because a clock that is a minute out, or two machines
editing the same document, would otherwise print a commit above its own parent.

### P2 — Every read of the repository was a separate round trip to the document

`CommitManager` held an `ObjectStore` and called `store.load()` — which reads
and re-parses the whole Custom XML Part — inside every method, including inside
the graph walks those methods called. A single `merge()` loaded and parsed the
entire repository **six times**: once for itself, three times through
`reconstruct`, and twice more through `checkoutSilent` on the fast-forward path.
`handleSwitchBranch` did it three times for a branch switch.

The walks are now pure functions over a `RepoXml` in `src/vcs/History.ts`, and
each operation loads once. `BranchManager.switchBranch` returns the tip it
switched to, for the same reason.

Two more things fell out of that. `reconstruct` walked to the root every time
and then insisted the root be a snapshot, so rebuilding the tip of a long branch
replayed every delta ever made; it now stops at the most recent snapshot at or
below the commit, which is what makes storing a merge as a snapshot worth
anything. And a cycle in the parent chain — which a damaged part can produce —
looped forever; it is now reported.

### P2 — Deleting the part before adding its replacement

`ObjectStore.save()` deleted the existing Custom XML Part, synced, then added
the new one, with a comment saying Office.js had no way to update a part in
place. It does: `CustomXmlPart.setXml`, in WordApi 1.4. The original's order
meant that anything going wrong between the two calls — an error, the document
closing, Word stopping — left the document with no history at all. It now
replaces in place, and where the host cannot, adds before it removes. If a
remove is what fails, the document is left with two parts, and `load()` takes
the one with the longer history, since the newer write is always the longer one.

### P3 — Smaller things

- **`sha6` was not a hash and lost entropy at random.**
  `Math.random().toString(36).slice(2, 6)` assumes a fixed-length rendering:
  `0.5` renders as `"0.i"`, so the slice yields one character and `padEnd` filled
  the rest with zeroes. Commit ids are now generated against the set already in
  use and retried on collision, so two commits in the same millisecond cannot
  collide at all.
- **The theme-change handler was never registered.** `Office.EventType` has
  `OfficeThemeChanged`; the code asked for `officeThemeChanged`, which is
  `undefined`, and silenced the resulting type error with `as any`. The task
  pane kept its startup colours when Office switched between light and dark.
- **Both downloads revoked their object URL on the line after `click()`**, and
  clicked an anchor that was never in the document. `click()` only *starts* a
  download; revoking synchronously can invalidate the URL before the fetch
  behind it has read it, and a detached anchor does nothing at all in some
  engines. Both are now in `src/taskpane/download.ts` with a test.
- **The clean export left orphans and matched the wrong entries.** It did not
  remove the `itemProps` part the item's rels file names, or that part's
  content-type entry. It matched the `[Content_Types].xml` and relationship
  entries by interpolating the part name into a pattern, which also matches a
  part whose name merely *contains* it (`item1.xml.bak`) and misses an element
  written with a separate closing tag. And it identified a WordGit part by
  searching the whole text for the namespace, so a user's own Custom XML Part
  that merely quoted it — a document *about* WordGit — would have been deleted
  from their file. Now in `src/vcs/DocxCleaner.ts`, matching on the root element
  and reading attributes off whole elements, with 21 tests over real zips.
- **A slice failure in `getDocumentBytes` could settle the promise twice** and
  close the file twice, because the other callbacks kept running.
- **`src/commands/commands.ts` was the Outlook sample.** It built an
  `Office.NotificationMessageDetails` and called
  `Office.context.mailbox.item.notificationMessages.replaceAsync`. There is no
  mailbox in Word. Nothing invoked it — the manifest declares the FunctionFile
  but no command that uses it — so it was dead rather than broken, but it was
  Outlook code in a Word add-in referencing a LICENSE that did not exist.
- **Compression was applied even when it made the payload longer.** LZString's
  base64 form is about a third larger than its raw output, so a short patch came
  back longer than it went in — and since the whole repository is rewritten on
  every commit, storing the larger form costs again on every future write.
- **The author was the literal string `"Author"`.** See the correction below.
- **No `LICENSE`.** Added, MIT, matching what `package.json` declares.

---

## A correction to one of the first review's suggestions

The review suggested pulling the author name from `Office.context.displayName`.
That property is the **add-in's** display name, not the user's. Word's
JavaScript API does not expose the signed-in user's name to a task pane at all;
only Outlook does, through `Office.context.mailbox.userProfile`.

What is available is the document's own `lastAuthor` property, which
`WordDocumentIO.authorName()` now reads, falling back to the literal `Author`
where the host will not say. It is asked per commit rather than once at
initialisation, so a document passing between two people records who wrote which
commit — which is why the format carries an author on the commit as well as one
in the config. It is the document's author rather than the user's, and the
README says so.

---

## What was deliberately not done

**The repository is still one Custom XML Part.** The first review's P2 — one
part per commit plus a small index — is the right structural answer and is still
open. It is a storage-format change needing a version field and a migration
path, and the data lives inside users' files where it cannot be migrated
centrally. Doing it in the same pass as the correctness fixes would have meant
two format changes landing together with no way to tell which one broke a
document. The compatibility rules are written down at the top of
`src/vcs/RepoFormat.ts` for whoever does it.

**No production manifest variant.** The review asked for a second manifest, or a
build-time substitution, for a non-localhost `SourceLocation`. Nobody knows where
this would be hosted, and a manifest with an invented URL in it is worse than no
manifest — it is the kind of file that gets deployed by accident. The README says
exactly which URLs to change and warns not to touch the `<Id>`.

**The UI was not changed**, per the standing constraint. Three things are worth
knowing and were left alone:

- **"Show diff" replaces the document body with no warning.** It lays down the
  parent silently, then inserts the child with tracking on, so Word draws the
  difference — which means any uncommitted work in the document is gone, with no
  undo step of its own. It reads like a read-only view and is not one. The
  README says so now; a confirmation dialogue would be a UI change.
- **"Download clean copy" has no button.** It is on `Alt+S` and nowhere else.
- **The commit and branch panels do not refresh each other.** Committing does
  not update the log tab.

**`npm run lint` has never passed and still does not.** It reports 70 errors, of
which about 60 are `no-undef` for DOM globals — the project's `.eslintrc.json`
is not the config `office-addin-lint` actually reads, so setting `env.browser`
there changes nothing (tried; the error count was identical, and the change was
reverted rather than left in as decoration). The baseline was 79 errors. Fixing
this means working out how `office-addin-lint` resolves its configuration, which
is a separate job from this one.

**`docs/` was renamed but its contents were not rewritten.** They are the
author's own design notes and are worth keeping — the architecture, the UI
decisions and the "why blame was dropped" reasoning are the best account of why
the project is shaped as it is. Both files now carry a note at the top saying
they are superseded in places and pointing at the README and `RepoFormat.ts` for
the current contract. **Both still open with "Instructions for Claude" blocks
telling a future session to update them; that is the author's call, not a
cleanup decision.**

---

## Structure

`src/vcs/` was already the better half of this project. It now has nothing in it
that calls Office.js directly:

- `RepoFormat.ts` — serialise, parse, validate. No DOM global at load time; the
  parser is injectable, which is how the format is tested without a browser.
- `ObjectStore.ts` — the Custom XML Parts, behind a `CustomXmlPartHost`
  interface. The Office.js implementation is 40 lines at the top of the file.
- `History.ts` — replay, ancestry, merge base, log, payload encoding. Pure.
- `CommitManager.ts` — takes a `DocumentIO`, which is `getOoxml`, `replaceBody`
  and `authorName`. That interface is what makes a whole
  init-commit-branch-merge sequence runnable in Node.
- `DocxCleaner.ts` — new, extracted from the task pane.

`taskpane.ts` went from 942 lines to 595, with the markup in `render.ts`,
the graph in `graphSvg.ts` and the download in `download.ts`.

---

## Tests

`npm test`. 195 checks, `node --test`, no framework. `pretest` compiles `src/` to
`tests/build/` with `tsc`, so the suite runs the real source.

Each fix has a matching assertion in `tests/original.test.js` that runs the code
it replaced and asserts the wrong answer it gave. `tests/original.js` is that
code transcribed, and the first test in the file reads the baseline commit out of
git and asserts every transcribed line is still in it — so a transcription that
has drifted, or been edited to make a test pass, fails loudly.

Things worth stealing:

- `tests/xml.js` is a DOMParser in about 320 lines with the three behaviours this
  format depends on: **line-ending normalisation before anything else sees the
  text**, which is the bug the CDATA fix is about; **a `<parsererror>` document
  instead of an exception**, which is why the corruption bug is a bug at all and
  which a shim that throws would have hidden; and `textContent` that folds CDATA
  in. It is strict about well-formedness on purpose — a test that feeds in
  damaged XML has to get the same answer a browser would.
- The shape check in `tests/render.test.js`, and the test that asserts it fails
  against the original builders. A shape check that passes everything proves
  nothing, and that is not visible from reading it.
- `tests/harness.js`'s `FakeParts` and `FakeDocument`. Forty lines each, and
  between them they turn "untestable without Word" into 28 end-to-end checks.

---

## Notes for the next session

- **Nothing here has been loaded into Word.** Everything above is reasoned from
  the source, the Office.js typings and 195 automated checks, and none of it has
  seen a real document. The manual pass is at the bottom of this file.
- `CustomXmlPart.setXml` is WordApi 1.4. If the add-in has to support an older
  host, the add-before-remove path is already there and tested; if not, the
  runtime check in `WordCustomXmlParts.setXml` can go.
- The one-part-per-commit change is the next real piece of work, and the format
  rules for doing it safely are at the top of `RepoFormat.ts`.
- If you touch the merge, read `tests/merge.test.js` before you loosen anything.
  The overlap check is conservative on purpose: two edits inside diff-match-patch's
  four-character context window conflict, and that is wanted.

---

## Manual pass — about twenty minutes, needs Word

Nothing in this environment can sideload an add-in, and sideloading also needs a
trusted development certificate, which `npm start` generates and installs the
first time. Run `npm install` then `npm start`; Word opens with WordGit in the
Home tab.

1. **Init and commit.** Type a paragraph, press **Init repo**, type another, and
   press **Commit** with a message. The Log tab should show two commits, newest
   first, with your name on them if Word knows it. This is the path that reads
   and writes the Custom XML Part for real, and the one thing no test here can
   confirm is that `setXml` behaves as the typings say.

2. **The one that matters: a damaged repository.** Save the document, close
   Word, open the `.docx` as a zip, and corrupt `customXml/itemN.xml` — delete
   its last forty bytes. Reopen in Word and open the task pane. It must say the
   data could not be read, and **Init, Commit and the message box must all be
   disabled.** In the original this presented as a document with no repository
   and an enabled Init button, one click from destroying the history. If this
   step behaves as it did before, stop and read `parseRepoXml` — everything else
   in this pass is cosmetic next to it.

3. **A merge that should conflict.** Commit a sentence. Create a branch, switch
   to it, rewrite that sentence one way, commit. Switch back to main, rewrite the
   same sentence a different way, commit. Merge the branch in. It must refuse,
   with a warning toast, and **the document must be unchanged**. The original
   produced a third sentence nobody wrote and reported success.

4. **A merge that should succeed.** Same setup, but each branch edits a
   different paragraph. Both edits should be in the document as tracked changes,
   and the Log tab should list the commits from *both* branches afterwards —
   which is the `parent2` fix, and which the original could not do.

5. **Bookmarks through a merge.** Before branching, insert a bookmark (Insert →
   Bookmark) around some text that neither branch will touch. After the merge,
   check the bookmark is still there and still spans the same text. The original
   stripped `w:id`, which unpaired it.

6. **Clean export.** Press **Alt+S** with the repository initialised. A file
   should download — that is the anchor and object-URL fix, and nobody has
   watched a real download. Open it in Word, open the task pane: it should offer
   Init, because the history is genuinely gone. Then confirm the original still
   has its history.

7. **Theme.** Switch Office between light and dark (File → Account → Office
   Theme) with the task pane open. It should follow. It never did.

## Licence, 2026-10-02

The author chose one licence for all of their projects: free for personal,
non-commercial use, modifiable for the user's own use, and not to be
distributed, original or modified. No standard licence says that, so `LICENSE`
is a short custom text, the same in every project apart from the year and what
it excludes. The README's licence section says the same, and `package.json`
declares `SEE LICENSE IN LICENSE`, npm's form for a licence with no SPDX
identifier. Anything above that calls the licence an open question, or calls
the project MIT, is superseded.

The manifest's `ProviderName` is now Manish Majhi rather than `nishantmajhi`,
at the author's request: it is the publisher name Office shows for the add-in.
`tests/manifest.test.js` pins both, and fails 2 of 16 against the previous
commit.

## CI, 2026-10-02

`.github/workflows/ci.yml` runs `npm test` (with its `tsc` pretest) and the
production webpack build on `ubuntu-latest`. Nothing has been pushed, so it
has not run on GitHub yet. Every job was run before committing, from a fresh
clone of this branch, the way a runner would run it: Linux jobs in a throwaway
Ubuntu 24.04 (the image `ubuntu-latest` uses) with the same tool versions.
196.

It fetches full history: the transcription tests read the baseline commit and
fail without it (checked). The uncommitted `package-lock.json` resync noted
under Licence is still uncommitted; CI uses the committed lockfile, which `npm
ci` accepts.
