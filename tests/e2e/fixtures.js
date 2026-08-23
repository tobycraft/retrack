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

// Loads the real taskpane.html in the browser, with the Office.js CDN
// script swapped for our mock. Everything else — tokenizer, Myers diff,
// paragraph scan/rewrite loop, DOM wiring — runs unmodified.
async function loadTaskpane(page, { paragraphs, actingAuthor = 'Test Author' }) {
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

  await page.goto('/taskpane.html');
  await page.waitForFunction(() => window.Office !== undefined);
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

module.exports = { loadTaskpane, serializeParagraphs };
