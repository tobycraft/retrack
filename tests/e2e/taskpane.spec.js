const { test, expect } = require('@playwright/test');
const { loadTaskpane } = require('./fixtures');

const DOC_TAG = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';

// A minimal but well-formed OOXML fragment shaped like what
// range.getOoxml() returns for a plain selection, so extractRPr /
// extractDocumentOpenTag have real input to parse.
function fakeSelectionOoxml(rPr) {
  return (
    `<?xml version="1.0"?><pkg:package>` +
    `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
    `${DOC_TAG}<w:body><w:p><w:r>${rPr}<w:t>placeholder</w:t></w:r></w:p></w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`
  );
}

async function apply(page) {
  await page.click('#btn');
  await expect(page.locator('#status')).toHaveText('Done.');
  return page.evaluate(() => window.__mock.lastOoxml);
}

test.describe('Paste Minimal Changes — core diff scenarios', () => {
  test('replaces a single word in the middle, leaving prefix/suffix as plain text', async ({ page }) => {
    await loadTaskpane(page, {
      selectionText: 'The quick brown fox jumps over the lazy dog.',
      selectionOoxml: fakeSelectionOoxml('<w:rPr><w:b/></w:rPr>'),
      clipboardText: 'The quick red fox jumps over the lazy dog.',
    });

    const ooxml = await apply(page);

    expect(ooxml).toContain('<w:t xml:space="preserve">The quick </w:t>');
    expect(ooxml).toMatch(/<w:del w:id="1"[^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:delText xml:space="preserve">brown<\/w:delText><\/w:r><\/w:del>/);
    expect(ooxml).toMatch(/<w:ins w:id="2"[^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">red<\/w:t><\/w:r><\/w:ins>/);
    expect(ooxml).toContain('<w:t xml:space="preserve"> fox jumps over the lazy dog.</w:t>');

    // Formatting from the original selection is carried onto every run, not just the edit.
    const rPrCount = (ooxml.match(/<w:rPr><w:b\/><\/w:rPr>/g) || []).length;
    expect(rPrCount).toBe(4); // "The quick " / delete "brown" / insert "red" / " fox...dog."
  });

  test('pure insertion produces no deletions', async ({ page }) => {
    await loadTaskpane(page, {
      selectionText: 'Hello world',
      selectionOoxml: fakeSelectionOoxml(''),
      clipboardText: 'Hello brave world',
    });

    const ooxml = await apply(page);

    expect(ooxml).not.toContain('<w:del ');
    expect(ooxml).toContain('<w:t xml:space="preserve">Hello </w:t>');
    expect(ooxml).toMatch(/<w:ins[^>]*><w:r><w:t xml:space="preserve">brave <\/w:t><\/w:r><\/w:ins>/);
    expect(ooxml).toContain('<w:t xml:space="preserve">world</w:t>');
  });

  test('replaces the whole selection when nothing in common is found', async ({ page }) => {
    await loadTaskpane(page, {
      selectionText: 'abc',
      selectionOoxml: fakeSelectionOoxml(''),
      clipboardText: 'xyz',
    });

    const ooxml = await apply(page);

    expect(ooxml).toMatch(/<w:del[^>]*><w:r><w:delText xml:space="preserve">abc<\/w:delText><\/w:r><\/w:del>/);
    expect(ooxml).toMatch(/<w:ins[^>]*><w:r><w:t xml:space="preserve">xyz<\/w:t><\/w:r><\/w:ins>/);
  });

  test('identical clipboard and selection produce no tracked changes', async ({ page }) => {
    await loadTaskpane(page, {
      selectionText: 'No changes here.',
      selectionOoxml: fakeSelectionOoxml(''),
      clipboardText: 'No changes here.',
    });

    const ooxml = await apply(page);

    expect(ooxml).not.toContain('<w:ins ');
    expect(ooxml).not.toContain('<w:del ');
    expect(ooxml).toContain('<w:t xml:space="preserve">No changes here.</w:t>');
  });

  test('records the acting author and uses replace as the insert location', async ({ page }) => {
    await loadTaskpane(page, {
      selectionText: 'old text',
      selectionOoxml: fakeSelectionOoxml(''),
      clipboardText: 'new text',
      author: 'Jane Reviewer',
    });

    const ooxml = await apply(page);
    const insertLocation = await page.evaluate(() => window.__mock.lastInsertLocation);

    expect(ooxml).toContain('w:author="Jane Reviewer"');
    expect(insertLocation).toBe('Replace');
  });
});
