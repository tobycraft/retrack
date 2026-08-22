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

## How it works

Tokenizes both texts into word, punctuation, and whitespace tokens, then finds the longest common prefix and suffix. Only the differing middle region is wrapped in `<w:ins>` / `<w:del>` OOXML tags. The result is injected via `Range.insertOoxml()` using the Office.js API.

See [CLAUDE.md](CLAUDE.md) for full architecture and design notes.

## Testing

The add-in itself has zero build/npm dependencies. `package.json` and `tests/` are dev-only tooling for an end-to-end test suite — nothing under them is deployed to GitHub Pages.

Real Word can't be automated in CI, so the tests load the actual `docs/taskpane.html` in a real browser via [Playwright](https://playwright.dev), swap the Office.js CDN script for a small mock of the Word JS API (`tests/e2e/mock-office.js`), and stub `navigator.clipboard`. The tokenizer, diff engine, and OOXML builder run unmodified — only the Word/Office host boundary is faked.

```bash
npm install
npm run test:e2e
```

Covers the core diff scenarios: mid-sentence word replacement, pure insertion, whole-selection replacement, and no-op (identical) paste — asserting on the generated `<w:ins>`/`<w:del>` OOXML, including that `<w:rPr>` formatting is preserved on every run.
