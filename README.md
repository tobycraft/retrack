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
