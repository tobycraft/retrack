const { test, expect } = require('@playwright/test');
const { loadCommandsPage, runMinimizeChanges, serializeParagraphs } = require('./fixtures');

test.describe('ReTrack — ribbon command (commands.html)', () => {
  test('minimizes a single delete+insert pair to a word-level diff', async ({ page }) => {
    await loadCommandsPage(page, {
      paragraphs: [
        [
          { text: 'The quick ' },
          { text: 'brown fox', del: { author: 'Test Author' } },
          { text: 'red fox', ins: { author: 'Test Author' } },
          { text: ' jumps.' },
        ],
      ],
    });

    await runMinimizeChanges(page);

    const dialogUrl = await page.evaluate(() => window.__mock.lastDialogUrl);
    expect(dialogUrl).toBeUndefined();

    // The diff engine's edit script deletes then inserts; the apply loop
    // positions each insertion at the start of the deletion it follows, so
    // the rewritten insertion lands immediately before the (still-present,
    // struck-through) deleted text.
    expect(await serializeParagraphs(page)).toEqual([
      [
        { text: 'The quick ', revision: null },
        { text: 'red ', revision: { type: 'Added', author: 'Test Author' } },
        { text: 'brown ', revision: { type: 'Deleted', author: 'Test Author' } },
        { text: 'fox jumps.', revision: null },
      ],
    ]);
  });

  test('minimizes multiple pairs across paragraphs', async ({ page }) => {
    await loadCommandsPage(page, {
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

    await runMinimizeChanges(page);

    expect(await serializeParagraphs(page)).toEqual([
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

    await loadCommandsPage(page, { paragraphs: original, actingAuthor: 'Test Author' });

    await runMinimizeChanges(page);

    expect(await serializeParagraphs(page)).toEqual(original.map(parts => parts.map(p => ({
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

    await loadCommandsPage(page, { paragraphs: original });

    await runMinimizeChanges(page);

    expect(await serializeParagraphs(page)).toEqual([
      [
        { text: 'Hello ', revision: null },
        { text: 'world', revision: { type: 'Added', author: 'Test Author' } },
      ],
    ]);
  });

  test('auto-detects the reviewer name with no input, even with a pair at the very start of the document', async ({ page }) => {
    // The author-detection probe inserts its throwaway character at the very
    // start of the document — this document puts a real tracked-change pair
    // in exactly that spot, to prove the probe doesn't corrupt it.
    await loadCommandsPage(page, {
      paragraphs: [
        [
          { text: 'old', del: { author: 'Detected Reviewer' } },
          { text: 'new', ins: { author: 'Detected Reviewer' } },
          { text: ' rest.' },
        ],
      ],
      actingAuthor: 'Detected Reviewer',
    });

    await runMinimizeChanges(page);

    const dialogUrl = await page.evaluate(() => window.__mock.lastDialogUrl);
    expect(dialogUrl).toBeUndefined();

    expect(await serializeParagraphs(page)).toEqual([
      [
        { text: 'new', revision: { type: 'Added', author: 'Detected Reviewer' } },
        { text: 'old', revision: { type: 'Deleted', author: 'Detected Reviewer' } },
        { text: ' rest.', revision: null },
      ],
    ]);
  });

  test('opens the error dialog and still signals completion when something fails', async ({ page }) => {
    await loadCommandsPage(page, {
      paragraphs: [[{ text: 'Hello world.' }]],
    });

    await page.evaluate(() => {
      window.Word.run = () => Promise.reject(new Error('boom'));
    });

    await runMinimizeChanges(page); // resolves only once event.completed() is called

    const dialogUrl = await page.evaluate(() => window.__mock.lastDialogUrl);
    expect(dialogUrl).toContain('dialog.html');
    expect(dialogUrl).toContain('msg=boom');
  });
});
