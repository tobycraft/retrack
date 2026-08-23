# Minimize My Changes — Word Add-in

## Problem

Microsoft Word's Track Changes treats a retyped or pasted replacement as "delete entire selection, insert entire replacement" — even if only one word changed. This makes reviewing tracked changes painful when iterating on paragraphs.

## Solution

A Word Add-in hosted on GitHub Pages with a single button: it scans the document's *existing* tracked changes for delete+insert pairs (the clunky result of a sloppy edit), computes a word-level diff between each pair's deleted and inserted text, and replays it as native Word edits — with Track Changes on — so only the words that actually changed stay marked. Everything else reverts to plain, unmarked text.

No clipboard involved. The "before" and "after" text for a change both already live in the document, as the tracked-changes revisions Word itself recorded — the add-in just re-diffs and re-marks them at the word level instead of the whole-selection level.

## Architecture

```
GitHub Pages (docs/)          Word (desktop/web)
┌─────────────────┐          ┌───────────────────────────┐
│ taskpane.html    │◄─HTTPS──│ Webview (task pane)        │
│ assets/icons     │          │                            │
└─────────────────┘          │ Office.js API (WordApi 1.6+):
                              │  - body.paragraphs         │
No server logic.              │  - paragraph.getTrackedChanges()
No build step.                │  - TrackedChange.reject()  │
No npm.                       │  - range.insertText()/delete()/search()
                              └───────────────────────────┘
```

- **Hosting:** GitHub Pages serving the `docs/` folder. Free HTTPS on `*.github.io`.
- **Manifest:** `manifest.xml` at repo root, sideloaded into Word. Contains GitHub Pages URLs.
- **Add-in UI:** Single `taskpane.html` — all HTML, CSS, JS, diff logic, and the tracked-changes scan/rewrite loop in one file.
- **Only external dependency:** Microsoft's `office.js` CDN (`appsforoffice.microsoft.com/lib/1/hosted/office.js`), mandatory for all Office Add-ins.

## Diff approach

Tokenize both texts into word/whitespace tokens (each non-space run absorbs its trailing whitespace), then run a **Myers diff** over the token arrays — the same LCS-based algorithm `git diff` uses — to get the minimal equal/delete/insert edit script.

(An earlier VBA prototype used prefix/suffix matching — count matching tokens from each end, treat the differing middle as one change — to sidestep persistent array-bounds bugs in an LCS backtrack. The add-in runs in a modern JS engine and doesn't hit that problem, so it uses the more accurate Myers diff instead. See decision 3 below.)

## Locating and rewriting each change

1. Walk `document.body.paragraphs`. Within each paragraph, scan left to right for an adjacent `Deleted`-run + `Added`-run pair (in either order) via `paragraph.getTrackedChanges()` — Word's own record of a replace-style edit. Only a pair where every revision is authored by the name entered in the task pane counts as a match; a lone deletion/insertion (nothing to minimize) or a pair by another author is skipped.
2. Capture the pair's combined range, then call `reject()` on every revision in it. This restores plain, untracked text — the original "before" content, un-marked.
3. Diff that restored text against the pair's original inserted text (see Diff approach above), then replay the edit script as `range.insertText()` / `range.delete()` calls with Track Changes on — Word itself records the new minimal `<w:ins>`/`<w:del>` marks. No manual OOXML construction, and (unlike `getOoxml()`/`insertOoxml()`, which hang indefinitely on Word for the web despite being nominally supported) this works cross-platform.
4. Resume scanning from the end of the just-rewritten group, never from the paragraph start — otherwise the freshly minimized pair (itself an adjacent Deleted+Added pair) would be rediscovered and re-diffed into itself forever on words with no shorter minimization.

## Setup

```powershell
.\setup.ps1 <github-username>    # rewrites manifest URLs
git push                          # deploys via Pages
# Word → Insert → Get Add-ins → Upload My Add-in → manifest.xml
```

GitHub Pages config: Settings → Pages → Source: main branch, `/docs` folder.

## File structure

```
track-minimal-changes/
├── manifest.xml              # Office Add-in manifest (sideload this)
├── setup.ps1 / setup.sh      # One-time: replace YOUR_GITHUB_USERNAME in manifest
├── .gitignore
├── README.md
└── docs/                     # Served by GitHub Pages
    ├── taskpane.html         # Entire add-in: UI + diff engine + tracked-changes scan/rewrite
    └── assets/
        └── icon-{16,32,80}.png
```

## Known limitations

- **Author scope:** Only rewrites tracked-change pairs where every revision is authored by the exact name typed into the task pane (matched against `TrackedChange.author`, no fuzzy matching). Other reviewers' changes are left untouched. Any group it does rewrite loses its original author/timestamp — the new minimal revisions are stamped with your name and the current time, since Office.js has no API to read the current Word user's name outside Outlook (would require SSO/Azure AD — out of scope for a zero-infrastructure add-in).
- **Single diff region per group:** A replace-style edit whose deleted/inserted text spans a paragraph break isn't merged into one group — each paragraph is scanned independently. Fine for the typical single-paragraph edit.
- **Mixed formatting:** Inserted text inherits the formatting of the text immediately before it (native `insertText()` behavior) — no manual run-properties handling needed. Single-style paragraphs are preserved perfectly.
- **Plain text only:** Tables, images, and structured content are not scanned or diffed.

## Design decisions & history

1. **Started as a Python CLI tool** — generated .docx files with tracked changes. Rejected: separate tool, not integrated into Word workflow.
2. **Office Add-in with npm/Node dev server** — full taskpane add-in with PowerShell HTTPS server. Rejected: user asked why a task pane needs a running server. Fair point.
3. **VBA macro** — zero infrastructure, lives inside Word. LCS diff engine had persistent "subscript out of range" errors across two rewrites (array bounds in backtrack/merge pipeline). Simplified to prefix/suffix approach which worked but VBA is desktop-only and can't be published.
4. **GitHub Pages add-in, clipboard-based** — same prefix/suffix diff and OOXML approach as the working VBA version, but as a proper Office Add-in hosted for free on Pages: select text in Word, paste the revised version into a textarea in the task pane, click Apply. Cross-platform, distributable via manifest, zero ongoing infrastructure. The diff engine was later upgraded from prefix/suffix to a full Myers diff, and `insertOoxml()` (which hangs indefinitely on Word for the web) was replaced with native `insertText()`/`delete()` calls.
5. **Document-scan-based diff, no clipboard** — final form. Testing showed the manual copy/paste step from decision 4 was still too much friction. Office.js's `getTrackedChanges()`/`TrackedChange.reject()` (WordApi 1.6) let the add-in read the "before" and "after" text directly off the delete+insert pairs Word already recorded for a sloppy edit, instead of from the clipboard. One button now scans the whole document, re-diffs every matching pair at the word level, and replays the result as native edits — the user only has to type their reviewer name once, up front.

## Distribution options

- **Personal:** Sideload `manifest.xml` directly.
- **Org-wide:** Upload manifest through Microsoft 365 admin centre.
- **Public:** Submit to Microsoft Marketplace (requires Partner Center account and certification).
