const { test, expect } = require('@playwright/test');
const { loadTaskpane, serializeParagraphs } = require('./fixtures');

async function minimize(page, authorName) {
  await page.fill('#authorName', authorName);
  await page.click('#btn');
  await expect(page.locator('#status')).not.toHaveText('Working…');
  return serializeParagraphs(page);
}

test.describe('Minimize My Changes — tracked-changes scan/rewrite', () => {
  test('minimizes a single delete+insert pair to a word-level diff', async ({ page }) => {
    await loadTaskpane(page, {
      paragraphs: [
        [
          { text: 'The quick ' },
          { text: 'brown fox', del: { author: 'Test Author' } },
          { text: 'red fox', ins: { author: 'Test Author' } },
          { text: ' jumps.' },
        ],
      ],
    });

    const result = await minimize(page, 'Test Author');

    await expect(page.locator('#status')).toHaveText('Minimized 1 change.');
    // The diff engine's edit script deletes then inserts; the apply loop
    // positions each insertion at the start of the deletion it follows, so
    // the rewritten insertion lands immediately before the (still-present,
    // struck-through) deleted text — matching how the pre-existing
    // selection-based apply loop this logic was copied from already behaved.
    expect(result).toEqual([
      [
        { text: 'The quick ', revision: null },
        { text: 'red ', revision: { type: 'Added', author: 'Test Author' } },
        { text: 'brown ', revision: { type: 'Deleted', author: 'Test Author' } },
        { text: 'fox jumps.', revision: null },
      ],
    ]);
  });

  test('minimizes multiple pairs across paragraphs', async ({ page }) => {
    await loadTaskpane(page, {
      paragraphs: [
        [
          { text: 'Alpha ' },
          { text: 'one two', del: { author: 'Test Author' } },
          { text: 'one three', ins: { author: 'Test Author' } },
          { text: ' end.' },
        ],
        [
          { text: 'Beta ' },
          { text: 'cat', del: { author: 'Test Author' } },
          { text: 'dog', ins: { author: 'Test Author' } },
          { text: ' end.' },
        ],
      ],
    });

    const result = await minimize(page, 'Test Author');

    await expect(page.locator('#status')).toHaveText('Minimized 2 changes.');
    expect(result).toEqual([
      [
        { text: 'Alpha one ', revision: null },
        { text: 'three', revision: { type: 'Added', author: 'Test Author' } },
        { text: 'two', revision: { type: 'Deleted', author: 'Test Author' } },
        { text: ' end.', revision: null },
      ],
      [
        { text: 'Beta ', revision: null },
        { text: 'dog', revision: { type: 'Added', author: 'Test Author' } },
        { text: 'cat', revision: { type: 'Deleted', author: 'Test Author' } },
        { text: ' end.', revision: null },
      ],
    ]);
  });

  test('leaves a pair authored by someone else untouched', async ({ page }) => {
    const original = [
      [
        { text: 'Beta ' },
        { text: 'three', del: { author: 'Someone Else' } },
        { text: 'four', ins: { author: 'Someone Else' } },
        { text: ' end.' },
      ],
    ];

    await loadTaskpane(page, { paragraphs: original });

    const result = await minimize(page, 'Test Author');

    await expect(page.locator('#status')).toHaveText('No matching changes found to minimize.');
    expect(result).toEqual(original.map(parts => parts.map(p => ({
      text: p.text,
      revision: p.del ? { type: 'Deleted', author: p.del.author }
        : p.ins ? { type: 'Added', author: p.ins.author }
        : null,
    }))));
  });

  test('reports no matches for a document with no eligible tracked-change pairs', async ({ page }) => {
    const original = [
      [
        { text: 'Hello ' },
        { text: 'world', ins: { author: 'Test Author' } },
      ],
    ];

    await loadTaskpane(page, { paragraphs: original });

    const result = await minimize(page, 'Test Author');

    await expect(page.locator('#status')).toHaveText('No matching changes found to minimize.');
    expect(result).toEqual([
      [
        { text: 'Hello ', revision: null },
        { text: 'world', revision: { type: 'Added', author: 'Test Author' } },
      ],
    ]);
  });
});
