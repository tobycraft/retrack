# ReTrack — Word Add-in

## Problem

Microsoft Word's Track Changes treats a retyped or pasted replacement as "delete entire selection, insert entire replacement" — even if only one word changed. This makes reviewing tracked changes painful when iterating on paragraphs.

## Solution

A Word Add-in hosted on GitHub Pages with a single ribbon button, no task pane: it scans the document's *existing* tracked changes for delete+insert pairs (the clunky result of a sloppy edit), computes a word-level diff between each pair's deleted and inserted text, and replays it as native Word edits — with Track Changes on — so only the words that actually changed stay marked. Everything else reverts to plain, unmarked text.

No clipboard, no name to type. The "before" and "after" text for a change both already live in the document, as the tracked-changes revisions Word itself recorded — the add-in just re-diffs and re-marks them at the word level instead of the whole-selection level. And the reviewer name it filters by isn't typed in anywhere — it's read directly off a throwaway tracked edit the add-in makes itself (see "Locating and rewriting each change" below).

## Architecture

```
GitHub Pages (docs/)                Word (desktop/web)
commands.html   ◄──HTTPS── loaded hidden, runs minimizeChanges() when
                            the ribbon button (Add-in Command) is clicked
dialog.html     ◄──HTTPS── opened only on error, via displayDialogAsync()
assets/icons    ◄──HTTPS── ribbon button icon

No server logic. No build step. No npm.

Office.js API used (WordApi 1.6+ unless noted):
  - Office.actions.associate           (wire the ribbon button to a function)
  - body.paragraphs / paragraph.getRange()
  - range.getTrackedChanges() / TrackedChange.reject()
  - range.insertText() / delete() / search()
```

- **Hosting:** GitHub Pages serving the `docs/` folder. Free HTTPS on `*.github.io`.
- **Manifest:** `manifest.xml` at repo root, sideloaded into Word. Defines an Add-in Command (ribbon button) via `VersionOverrides` instead of a task pane, pointing at `commands.html` as its `FunctionFile`.
- **Add-in logic:** `commands.html` — a hidden page Word loads only to run `minimizeChanges()`; no HTML UI of its own. Contains all the diff logic and the tracked-changes scan/rewrite loop.
- **Error reporting:** `dialog.html` — a small static page opened via `Office.context.ui.displayDialogAsync()` only when something goes wrong. Nothing appears on success.
- **Only external dependency:** Microsoft's `office.js` CDN (`appsforoffice.microsoft.com/lib/1/hosted/office.js`), mandatory for all Office Add-ins. `dialog.html` doesn't even need this — it only ever displays a message baked into its own URL.

## Diff approach

Tokenize both texts into word/whitespace tokens (each non-space run absorbs its trailing whitespace), then run a **Myers diff** over the token arrays — the same LCS-based algorithm `git diff` uses — to get the minimal equal/delete/insert edit script.

(An earlier VBA prototype used prefix/suffix matching — count matching tokens from each end, treat the differing middle as one change — to sidestep persistent array-bounds bugs in an LCS backtrack. The add-in runs in a modern JS engine and doesn't hit that problem, so it uses the more accurate Myers diff instead. See decision 3 below.)

## Locating and rewriting each change

0. Before scanning anything, figure out "your name" the way Word itself would stamp it — Office.js has no API to read the current Word user's display name outside Outlook, so instead the add-in inserts one throwaway zero-width-space character at the very start of the document (with Track Changes on), reads the `author` Word just stamped on that tracked insertion via `getTrackedChanges()`, then calls `reject()` on it to remove the probe. The document is left exactly as it was found; the add-in now knows the exact author string to filter by, with no dialog and nothing stored between runs.
1. Walk `document.body.paragraphs`. Within each paragraph, scan left to right for an adjacent `Deleted`-run + `Added`-run pair (in either order) via `paragraph.getTrackedChanges()` — Word's own record of a replace-style edit. Only a pair where every revision is authored by the name from step 0 counts as a match; a lone deletion/insertion (nothing to minimize) or a pair by another author is skipped.
2. Before rejecting anything, read the pair's inserted text off the `Added` revision's `TrackedChange.text` — this is safe to read directly (see the `TrackedChange.text` note below). Then capture the pair's combined range and call `reject()` on every revision in it. This restores plain, untracked text — the original "before" content, un-marked — which the range's own `.text` now holds (see below for why this can't come from the `Deleted` revision's `TrackedChange.text` instead).
3. Diff that restored range text against the inserted text captured in step 2 (see Diff approach above), then replay the edit script as `range.insertText()` / `range.delete()` calls with Track Changes on — Word itself records the new minimal `<w:ins>`/`<w:del>` marks. No manual OOXML construction, and (unlike `getOoxml()`/`insertOoxml()`, which hang indefinitely on Word for the web despite being nominally supported) this works cross-platform.
4. Resume scanning from the end of the just-rewritten group, never from the paragraph start — otherwise the freshly minimized pair (itself an adjacent Deleted+Added pair) would be rediscovered and re-diffed into itself forever on words with no shorter minimization.
5. Nothing is reported on success — the rewritten tracked changes in the document *are* the feedback. If anything throws (e.g. the author probe fails, or an expected token can't be relocated mid-rewrite), `commands.html` opens `dialog.html` with the error message in the query string via `displayDialogAsync()`; on success it's never opened at all.

## Setup

```powershell
.\setup.ps1 <github-username>    # rewrites manifest URLs
git push                          # deploys via Pages
# Word → Insert → Get Add-ins → Upload My Add-in → manifest.xml
```

GitHub Pages config: Settings → Pages → Source: main branch, `/docs` folder.

## File structure

```
retrack/
├── manifest.xml              # Office Add-in manifest (sideload this)
├── setup.ps1 / setup.sh      # One-time: replace YOUR_GITHUB_USERNAME in manifest
├── .gitignore
├── README.md
└── docs/                     # Served by GitHub Pages
    ├── commands.html         # Entire add-in: diff engine + tracked-changes scan/rewrite, no visible UI
    ├── dialog.html            # Error popup only (opened via displayDialogAsync on failure)
    └── assets/
        └── icon-{16,32,80}.png
```

## Known limitations

- **Author scope:** Only rewrites tracked-change pairs where every revision is authored by the same name Word would stamp on a new edit right now (detected via the throwaway-probe technique in step 0 above, matched exactly against `TrackedChange.author`, no fuzzy matching). Other reviewers' changes are left untouched. Any group it does rewrite loses its original author/timestamp — the new minimal revisions are stamped with the current time, same as the probe-detected name.
- **Single diff region per group:** A replace-style edit whose deleted/inserted text spans a paragraph break isn't merged into one group — each paragraph is scanned independently. Fine for the typical single-paragraph edit.
- **Mixed formatting:** Inserted text inherits the formatting of the text immediately before it (native `insertText()` behavior) — no manual run-properties handling needed. Single-style paragraphs are preserved perfectly.
- **Plain text only:** Tables, images, and structured content are not scanned or diffed.
- **Silent success:** There's no confirmation that anything happened beyond the tracked changes themselves looking different — by design (see decision 6), but worth knowing if you click the button on a document with nothing to minimize.
- **`TrackedChange.text` reads as empty for a `Deleted` revision:** confirmed against real Word for Mac (not just documentation) — `TrackedChange.text` reflects the text as it would read if the change were accepted, which is correct for an `Added` revision (accepting keeps the text) but always empty for a `Deleted` one (accepting removes it), even though the struck-through text is still visually present. The add-in works around this by reading the inserted text directly from the `Added` revision's `.text` before rejecting, but recovering the deleted text from the group's range `.text` only after `reject()` has restored it to plain text (see step 2 above) — never from the `Deleted` revision's own `.text`.

## Design decisions & history

1. **Started as a Python CLI tool** — generated .docx files with tracked changes. Rejected: separate tool, not integrated into Word workflow.
2. **Office Add-in with npm/Node dev server** — full taskpane add-in with PowerShell HTTPS server. Rejected: user asked why a task pane needs a running server. Fair point.
3. **VBA macro** — zero infrastructure, lives inside Word. LCS diff engine had persistent "subscript out of range" errors across two rewrites (array bounds in backtrack/merge pipeline). Simplified to prefix/suffix approach which worked but VBA is desktop-only and can't be published.
4. **GitHub Pages add-in, clipboard-based** — same prefix/suffix diff and OOXML approach as the working VBA version, but as a proper Office Add-in hosted for free on Pages: select text in Word, paste the revised version into a textarea in the task pane, click Apply. Cross-platform, distributable via manifest, zero ongoing infrastructure. The diff engine was later upgraded from prefix/suffix to a full Myers diff, and `insertOoxml()` (which hangs indefinitely on Word for the web) was replaced with native `insertText()`/`delete()` calls.
5. **Document-scan-based diff, no clipboard** — Office.js's `getTrackedChanges()`/`TrackedChange.reject()` (WordApi 1.6) let the add-in read the "before" and "after" text directly off the delete+insert pairs Word already recorded for a sloppy edit, instead of from the clipboard. One button, in a task pane, scanned the whole document, re-diffed every matching pair at the word level, and replayed the result as native edits — the user had to type their reviewer name once, up front, since Office.js can't read it directly.
6. **Ribbon command, no task pane, name auto-detected** — final form. Asked "why do I need a task pane at all, and why does it ask for my name when Word already knows it?" Both were fair. The task pane's only remaining jobs were hosting a name field and a status line — neither was load-bearing. Converted the manifest to an Add-in Command: a plain ribbon button (`Action xsi:type="ExecuteFunction"`) that runs a hidden `commands.html` with no UI of its own. For the name: rather than ask, the add-in now makes one throwaway tracked edit of its own (an invisible zero-width space) right before scanning, reads the `author` Word just stamped on it, and rejects the edit away — that's Word telling the add-in its own user name the same way it stamps every real edit, so there's nothing to type and nothing that can go stale. For feedback: success is silent (the rewritten tracked changes are the only visible result); failure opens a small `dialog.html` popup via `displayDialogAsync()`, since Word has no toast/notification API outside Outlook and this was judged better than resurrecting a task pane just for status text.

## Distribution options

- **Personal:** Sideload `manifest.xml` directly.
- **Org-wide:** Upload manifest through Microsoft 365 admin centre.
- **Public:** Submit to Microsoft Marketplace (requires Partner Center account and certification).
