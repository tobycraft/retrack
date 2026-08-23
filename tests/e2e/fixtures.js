const fs = require('fs');
const path = require('path');

const mockOfficeJs = fs.readFileSync(path.join(__dirname, 'mock-office.js'), 'utf-8');

// Converts a human-friendly paragraph spec into the {chars, liveRanges}
// document model mock-office.js operates on. Each paragraph is an array of
// parts: {text}, {text, del: {author}}, or {text, ins: {author}}.
function buildParagraphs(spec) {
  return spec.map(parts => {
    const chars = [];
    parts.forEach(part => {
      const revision = part.del ? { type: 'Deleted', author: part.del.author }
        : part.ins ? { type: 'Added', author: part.ins.author }
        : null;
      for (const ch of part.text) chars.push({ ch, revision });
    });
    return { chars, liveRanges: [] };
  });
}

// Loads the real docs/commands.html in the browser, with the Office.js CDN
// script swapped for our mock. Everything else — tokenizer, Myers diff,
// author-detection probe, paragraph scan/rewrite loop — runs unmodified.
// actingAuthor is "the current Word user" as far as the mock is concerned:
// it's stamped on every new tracked edit the mock makes (including the
// commands.html author-detection probe), exactly like a real Word user name.
async function loadCommandsPage(page, { paragraphs, actingAuthor = 'Test Author' }) {
  const doc = { paragraphs: buildParagraphs(paragraphs) };

  await page.addInitScript(
    ({ doc, actingAuthor }) => {
      window.__mock = { doc, actingAuthor };
    },
    { doc, actingAuthor }
  );

  await page.route('**/office.js', (route) =>
    route.fulfill({ status: 200, contentType: 'application/javascript', body: mockOfficeJs })
  );

  await page.goto('/commands.html');
  await page.waitForFunction(() => !!(window.__mock.registeredActions && window.__mock.registeredActions.minimizeChanges));
}

// Invokes the ribbon action exactly as Word would on a button click, and
// waits for event.completed() — the signal commands.html always sends,
// success or failure, once it's done.
function runMinimizeChanges(page) {
  return page.evaluate(() => new Promise(resolve => {
    window.__mock.registeredActions.minimizeChanges({ completed: resolve });
  }));
}

// Reads back the current document state as runs of contiguous same-markup
// text per paragraph, e.g. [{text: 'brown ', revision: {type: 'Deleted', author: 'X'}}, ...].
function serializeParagraphs(page) {
  return page.evaluate(() => {
    return window.__mock.doc.paragraphs.map(p => {
      const runs = [];
      let lastKey = null;
      p.chars.forEach(c => {
        const key = c.revision ? c.revision.type + '|' + c.revision.author : 'plain';
        if (lastKey === key) {
          runs[runs.length - 1].text += c.ch;
        } else {
          runs.push({ text: c.ch, revision: c.revision ? { type: c.revision.type, author: c.revision.author } : null });
          lastKey = key;
        }
      });
      return runs;
    });
  });
}

module.exports = { loadCommandsPage, runMinimizeChanges, serializeParagraphs };
