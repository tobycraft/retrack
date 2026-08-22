# Paste Minimal Changes — Word Add-in

## Problem

Microsoft Word's Track Changes treats paste operations as "delete entire selection, insert entire clipboard" — even if only one word changed. This makes reviewing tracked changes painful when iterating on paragraphs.

## Solution

A Word Add-in hosted on GitHub Pages that computes a word-level diff between the selected text and the clipboard, then writes OOXML with `<w:ins>` and `<w:del>` tags wrapping only the actual changes. The unchanged prefix and suffix remain clean.

## Architecture

```
GitHub Pages (docs/)          Word (desktop/web)
┌─────────────────┐          ┌──────────────────────┐
│ taskpane.html    │◄─HTTPS──│ Webview (task pane)   │
│ assets/icons     │          │                      │
└─────────────────┘          │ Office.js API:        │
                              │  - getSelection()    │
No server logic.              │  - getOoxml()        │
No build step.                │  - insertOoxml()     │
No npm.                       └──────────────────────┘
```

- **Hosting:** GitHub Pages serving the `docs/` folder. Free HTTPS on `*.github.io`.
- **Manifest:** `manifest.xml` at repo root, sideloaded into Word. Contains GitHub Pages URLs.
- **Add-in UI:** `taskpane.html` (task pane button) and `commands.html` (ribbon button, prototype — see Design decisions & history) both load the diff logic and OOXML builder from `paste-diff.js`.
- **Only external dependency:** Microsoft's `office.js` CDN (`appsforoffice.microsoft.com/lib/1/hosted/office.js`), mandatory for all Office Add-ins.

## Diff approach

**Prefix/suffix matching**, not LCS. Tokenize both texts into word/whitespace tokens, count matching tokens from the start (prefix) and end (suffix), replace only the middle region that differs.

Tradeoff: if two changes are far apart in the same selection (word 3 and word 50), the entire middle is marked as one change. For the typical "I edited a paragraph" case, the result is identical to a full LCS diff. We tried LCS first in VBA but it had persistent array bounds issues and was 10x the code.

## OOXML generation

1. Extract `<w:rPr>` (run properties — font, size, bold, etc.) from the selection's OOXML to preserve formatting.
2. Build runs: unchanged text in `<w:r>`, deleted text in `<w:del><w:r><w:delText>`, inserted text in `<w:ins><w:r><w:t>`.
3. Wrap in a `<pkg:package>` envelope.
4. Replace selection via `Range.insertOoxml(ooxml, Word.InsertLocation.replace)`.

## Setup

```powershell
.\setup.ps1 <github-username>    # rewrites manifest URLs
git push                          # deploys via Pages
# Word → Insert → Get Add-ins → Upload My Add-in → manifest.xml
```

GitHub Pages config: Settings → Pages → Source: main branch, `/docs` folder.

## File structure

```
paste-minimal-changes/
├── manifest.xml              # Office Add-in manifest (sideload this)
├── setup.ps1 / setup.sh      # One-time: replace YOUR_GITHUB_USERNAME in manifest
├── .gitignore
├── README.md
└── docs/                     # Served by GitHub Pages
    ├── taskpane.html         # Task pane UI (button + status)
    ├── commands.html         # Ribbon button function file (prototype, no task pane)
    ├── paste-diff.js         # Shared diff engine + OOXML builder
    └── assets/
        └── icon-{16,32,80}.png
```

## Known limitations

- **Single diff region:** Prefix/suffix approach, not multi-region. Fine for paragraph-level edits.
- **Mixed formatting:** Applies first run's `<w:rPr>` to all generated runs. Single-style paragraphs are preserved perfectly.
- **Plain text only:** Tables, images, and structured content in the selection are not preserved through the diff.

## Design decisions & history

1. **Started as a Python CLI tool** — generated .docx files with tracked changes. Rejected: separate tool, not integrated into Word workflow.
2. **Office Add-in with npm/Node dev server** — full taskpane add-in with PowerShell HTTPS server. Rejected: user asked why a task pane needs a running server. Fair point.
3. **VBA macro** — zero infrastructure, lives inside Word. LCS diff engine had persistent "subscript out of range" errors across two rewrites (array bounds in backtrack/merge pipeline). Simplified to prefix/suffix approach which worked but VBA is desktop-only and can't be published.
4. **GitHub Pages add-in** — final form. Same prefix/suffix diff and OOXML approach as the working VBA version, but as a proper Office Add-in hosted for free on Pages. Cross-platform (desktop + web), distributable via manifest, zero ongoing infrastructure.
5. **Exploring: ribbon button instead of task pane** — user feedback: the task pane itself (open sidebar, click a button inside it) is the complicated part, not the setup. Researched Office Add-in "Commands" (`ExecuteFunction` ribbon buttons, function-file/UI-less runtime): documented, real limitation is that `navigator.clipboard.readText()` requires a focused document + a user gesture that originates inside that document's DOM (OfficeDev/office-js#1991, #986); a ribbon click is a native Office UI event that never touches the add-in's iframe, so clipboard read may be unreliable or fail outright with no task pane to hold focus — the existing task pane code already has to special-case an occasional clipboard-read failure ("click inside the task pane first"), which is the *best case* for focus. Built `commands.html` as a live-test prototype (ships alongside the task pane in the same manifest) to get a real answer in Word rather than relying on docs/issues alone; failures are written into the document as a marker paragraph since a function command has no UI to show an error in. If clipboard read proves unreliable from the ribbon, the fallback design is to stop reading the clipboard via JS entirely: let the user do a normal `Ctrl+V` paste with Track Changes on (Word handles this natively, no permissions involved), then have the ribbon button find the `<w:del>`/`<w:ins>` pair Word just created and collapse it into a minimal diff using the same OOXML engine — sidesteps the Clipboard API altogether.

## Distribution options

- **Personal:** Sideload `manifest.xml` directly.
- **Org-wide:** Upload manifest through Microsoft 365 admin centre.
- **Public:** Submit to Microsoft Marketplace (requires Partner Center account and certification).
