const fs = require('fs');
const path = require('path');

const mockOfficeJs = fs.readFileSync(path.join(__dirname, 'mock-office.js'), 'utf-8');

// Loads the real taskpane.html in the browser, with the Office.js CDN
// script swapped for our mock and navigator.clipboard stubbed. Everything
// else — tokenizer, Myers diff, OOXML builder, DOM wiring — runs unmodified.
async function loadTaskpane(page, { selectionText, selectionOoxml, clipboardText, author = 'Test Author' }) {
  await page.addInitScript(
    ({ selectionText, selectionOoxml, clipboardText, author }) => {
      window.__mock = {
        selectionText,
        selectionOoxml,
        clipboardText,
        author,
        lastOoxml: null,
        lastInsertLocation: null,
      };
      Object.defineProperty(navigator, 'clipboard', {
        value: { readText: () => Promise.resolve(window.__mock.clipboardText) },
        configurable: true,
      });
    },
    { selectionText, selectionOoxml, clipboardText, author }
  );

  await page.route('**/office.js', (route) =>
    route.fulfill({ status: 200, contentType: 'application/javascript', body: mockOfficeJs })
  );

  await page.goto('/taskpane.html');
  await page.waitForFunction(() => window.Office !== undefined);
}

module.exports = { loadTaskpane };
