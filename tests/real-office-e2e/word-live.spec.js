// Real end-to-end test against actual Word for Mac, driven via AppleScript/
// System Events UI automation (tests/real-office-e2e/live-word.applescript)
// — see the "Testing" section in README.md and CLAUDE.md for why this is
// separate from the mocked suite in tests/e2e/.
//
// This does NOT run in CI and is not part of `npm run test:e2e`. It requires:
//   - A local, licensed Word for Mac install
//   - One-time macOS permission grants (System Settings > Privacy &
//     Security > Automation, and > Accessibility) for the terminal running
//     this — macOS will prompt for these on first run
//   - Word may show a one-time "Trust this add-in" prompt the first time
//     the sideloaded manifest loads
//
// Run with: npm run test:live
//
// What it does: sideloads manifest.xml (pointing at the live, deployed
// docs/commands.html on GitHub Pages — verified byte-identical to the local
// file), creates a blank Word document, types a plain sentence, makes a
// real tracked retype (Track Changes on, "brown fox" -> "red fox" — the
// classic whole-selection replace this add-in exists to clean up), clicks
// the real ReTrack ribbon button, saves the result, then asserts on the
// saved .docx's word/document.xml.
//
// No fixture .docx is checked in and no reviewer name is ever hardcoded:
// the tracked edit is made live by this same Word session, so whatever
// identity Word stamps on it is also what the add-in's own author-detection
// probe will read — verification only checks that the resulting w:ins/w:del
// share one author, never a specific name (avoids baking the user's real
// name into this repo).

const { test, expect } = require('@playwright/test');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const REPO_ROOT = path.join(__dirname, '..', '..');
const MANIFEST_SRC = path.join(REPO_ROOT, 'manifest.xml');
const WEF_DIR = path.join(os.homedir(), 'Library', 'Containers', 'com.microsoft.Word', 'Data', 'Documents', 'wef');
const OUTPUT_DIR = '/Users/tobiasruby/Library/CloudStorage/OneDrive-Personal/Sonstiges/retrack - testing';
const SCRIPT = path.join(__dirname, 'live-word.applescript');

const BASELINE = 'The quick brown fox jumps.';
const OLD_SPAN = 'brown fox'; // span that gets found (Cmd+F) and retyped
const NEW_TEXT = 'red fox';

function ensureSideloaded() {
  fs.mkdirSync(WEF_DIR, { recursive: true });
  const dest = path.join(WEF_DIR, 'manifest.xml');
  const src = fs.readFileSync(MANIFEST_SRC, 'utf-8');
  if (!fs.existsSync(dest) || fs.readFileSync(dest, 'utf-8') !== src) {
    fs.writeFileSync(dest, src);
  }
}

function runHandler(handler, ...args) {
  return execFileSync('osascript', [SCRIPT, handler, ...args], { encoding: 'utf-8' });
}

function timestampedOutputPath() {
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  return path.join(OUTPUT_DIR, `retrack-live-test-${stamp}.docx`);
}

// Minimal regex-based scan of a Word document.xml — not a general OOXML
// parser, just enough to pull out <w:ins>/<w:del> blocks (with their
// w:author and concatenated visible text) and the plain text outside them.
function scanRevisions(xml) {
  const blocks = [];
  const blockRe = /<w:(ins|del)\b[^>]*\bw:author="([^"]*)"[^>]*>([\s\S]*?)<\/w:\1>/g;
  let match;
  while ((match = blockRe.exec(xml))) {
    const [, kind, author, inner] = match;
    const textRe = kind === 'del' ? /<w:delText[^>]*>([^<]*)<\/w:delText>/g : /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let text = '';
    let textMatch;
    while ((textMatch = textRe.exec(inner))) text += textMatch[1];
    blocks.push({ kind, author, text });
  }
  const plainXml = xml.replace(blockRe, '');
  const plainTextRe = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let plainText = '';
  let plainMatch;
  while ((plainMatch = plainTextRe.exec(plainXml))) plainText += plainMatch[1];
  return { blocks, plainText };
}

test.describe('ReTrack — live Word (real Office desktop)', () => {
  test('minimizes a real tracked retype to a word-level diff', async () => {
    ensureSideloaded();

    const outputPath = timestampedOutputPath();

    // Deliberately no catch-and-close-on-failure here: if clickRibbonButton
    // or an earlier step throws, the in-progress Word window is left open
    // on screen so a failure can actually be inspected, rather than being
    // silently closed away.
    runHandler('createDocument');
    runHandler('typeBaseline', BASELINE);
    runHandler('retypeTracked', OLD_SPAN, NEW_TEXT);
    runHandler('clickRibbonButton', 'ReTrack');
    runHandler('saveAndClose', outputPath);

    const xml = execFileSync('unzip', ['-p', outputPath, 'word/document.xml'], { encoding: 'utf-8' });
    const { blocks, plainText } = scanRevisions(xml);

    const del = blocks.find(b => b.kind === 'del' && b.text.trim() === 'brown');
    const ins = blocks.find(b => b.kind === 'ins' && b.text.trim() === 'red');

    expect(del, `no minimized w:del "brown" found — revisions were: ${JSON.stringify(blocks)}`).toBeTruthy();
    expect(ins, `no minimized w:ins "red" found — revisions were: ${JSON.stringify(blocks)}`).toBeTruthy();
    expect(ins.author).toBe(del.author);

    expect(plainText).toContain('fox');
    expect(plainText).toContain('jumps');
    expect(blocks.some(b => b.text.includes('fox'))).toBe(false);
  });
});
