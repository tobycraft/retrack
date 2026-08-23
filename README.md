# Minimize My Changes

A Word Add-in that rewrites clunky whole-selection tracked changes into word-level tracked changes, so a one-word edit shows up as a one-word edit instead of "delete the whole sentence, insert the whole sentence."

## Setup

**1. Configure the manifest with your GitHub username:**

```powershell
.\setup.ps1 <your-github-username>
```

**2. Push to GitHub:**

```powershell
git push
```

**3. Enable GitHub Pages:** Settings → Pages → Source: `main` branch, `/docs` folder.

**4. Sideload the add-in in Word:** Insert → Get Add-ins → Upload My Add-in → select `manifest.xml`.

## Usage

1. Edit the document normally in Word with Track Changes on — Word will often record a whole retyped sentence as one big delete+insert.
2. Enter your name in the task pane exactly as it appears in Word's tracked changes (File → Options → General → User name).
3. Click **Minimize My Changes**.

Only the words that actually changed stay marked as tracked insertions/deletions; everything else reverts to plain, unmarked text. No clipboard involved — the add-in reads the "before" and "after" text straight from the tracked changes already in the document.

## How it works

Scans the document paragraph by paragraph for adjacent delete+insert tracked-change pairs authored by the name you entered (via Office.js's `getTrackedChanges()`). For each pair, it rejects both revisions (restoring plain original text), tokenizes the original and replacement text into word/whitespace tokens, runs a Myers diff to find the minimal edit script, and replays it as native `insertText()`/`delete()` calls with Track Changes on — so Word itself records the new word-level `<w:ins>`/`<w:del>` marks.

See [CLAUDE.md](CLAUDE.md) for full architecture and design notes.

## Publishing to Microsoft AppSource

There's no build step for the add-in itself — "building" a release means
validating the manifest and confirming the GitHub Pages site it points to is
live, then submitting that manifest through Partner Center.

**1. Validate the manifest:**

```bash
npm install
npm run validate
```

This runs Microsoft's own `office-addin-manifest` checker (schema, required
fields, HTTPS URLs) against `manifest.xml`.

**2. Confirm the hosted site is live** (Settings → Pages must be enabled,
`main` branch, `/docs` folder — see Setup above). Check that these all load
over HTTPS with no errors:
- `https://<username>.github.io/track-minimal-changes/taskpane.html`
- `https://<username>.github.io/track-minimal-changes/assets/icon-{16,32,80}.png`
- `https://<username>.github.io/track-minimal-changes/privacy.html`

**3. Submit via [Partner Center](https://partner.microsoft.com/dashboard):**
create a new Office Add-in offer and upload `manifest.xml` directly as the
package — there's no zip/build artifact for an add-in-only XML manifest.

**4. Listing assets you'll need to prepare separately** (these live in the
Partner Center form, not the repo):
- A 300×300 PNG store icon — `docs/assets/icon-300.png` (the source artwork
  is also kept at full resolution as `docs/assets/icon-original.png`).
- 1–5 PNG screenshots of the add-in in use in Word.
- A privacy policy URL — use the included `docs/privacy.html`
  (`https://<username>.github.io/track-minimal-changes/privacy.html`), which
  accurately states the add-in collects no data.
- A support URL — `manifest.xml` already points `SupportUrl` at this repo's
  GitHub Issues page.
- Short and long marketplace descriptions (separate from the manifest's
  `<Description>`, which stays short for the "My Add-ins" list).

Expect the AppSource review to take up to a few weeks, and don't be
surprised if the first submission comes back with feedback — that's normal.

## Testing

The add-in itself has zero build/npm dependencies. `package.json` and `tests/` are dev-only tooling for an end-to-end test suite — nothing under them is deployed to GitHub Pages.

Real Word can't be automated in CI, so the tests load the actual `docs/taskpane.html` in a real browser via [Playwright](https://playwright.dev), and swap the Office.js CDN script for a small mock of the Word JS API (`tests/e2e/mock-office.js`) that simulates a document's paragraphs and tracked changes. The tokenizer, diff engine, and scan/rewrite loop run unmodified — only the Word/Office host boundary is faked.

```bash
npm install
npm run test:e2e
```

Covers the core scenarios: a single delete+insert pair minimized to a word-level diff, multiple pairs across paragraphs, a pair skipped because it's authored by someone else, and a document with no matching tracked changes at all.
