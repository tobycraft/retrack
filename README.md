# Paste Minimal Changes

A Word Add-in that pastes with word-level tracked changes instead of the default delete-all/insert-all behaviour.

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

1. Select the paragraph (or text) you want to replace in Word.
2. Copy the revised version to your clipboard.
3. Click **Apply Minimal Paste** in the task pane.

Only the words that actually changed appear as tracked insertions/deletions. The unchanged prefix and suffix are written back as plain runs.

## Prototype: ribbon button, no task pane

There's a second, experimental entry point being validated alongside the task pane: an **"Apply Minimal Paste" ribbon button** on the Home tab (`docs/commands.html`) that runs the same diff engine without ever opening a task pane — select, copy, click the ribbon button, done.

The open question is whether `navigator.clipboard.readText()` works when triggered by a ribbon click instead of a click inside a visible task pane iframe — Office Add-in function commands run UI-less, and the Clipboard API normally requires a focused document and a user gesture that originates inside that document. See [CLAUDE.md](CLAUDE.md) for the research behind this.

**To test it:** sideload the manifest as usual (both entry points ship in the same `manifest.xml`), select text, copy new text to the clipboard, and click **Apply Minimal Paste** in the Home tab ribbon (not the task pane). Because a failed function command has no UI to show an error in, failures are written directly into the document as a red `[MinimalPaste TEST FAILED] ...` paragraph after your selection — that's the signal to watch for. No marker paragraph + correct tracked changes = it worked.

## How it works

Tokenizes both texts into word, punctuation, and whitespace tokens, then finds the longest common prefix and suffix. Only the differing middle region is wrapped in `<w:ins>` / `<w:del>` OOXML tags. The result is injected via `Range.insertOoxml()` using the Office.js API.

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

Real Word can't be automated in CI, so the tests load the actual `docs/taskpane.html` in a real browser via [Playwright](https://playwright.dev), swap the Office.js CDN script for a small mock of the Word JS API (`tests/e2e/mock-office.js`), and stub `navigator.clipboard`. The tokenizer, diff engine, and OOXML builder run unmodified — only the Word/Office host boundary is faked.

```bash
npm install
npm run test:e2e
```

Covers the core diff scenarios: mid-sentence word replacement, pure insertion, whole-selection replacement, and no-op (identical) paste — asserting on the generated `<w:ins>`/`<w:del>` OOXML, including that `<w:rPr>` formatting is preserved on every run.
