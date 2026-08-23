// Stand-in for Microsoft's office.js, served to the browser in place of the
// real CDN script during tests. Real Word/Office.js can't be automated in
// CI, so this mock reproduces just enough of the Word JS API surface
// (Office.onReady, Office.actions.associate, Office.context.ui.displayDialogAsync,
// Word.run, Body.getRange/paragraphs, Paragraph/Range.getRange,
// Range.expandTo/insertText/delete/search/getTrackedChanges,
// TrackedChange.reject) for commands.html's minimizeChanges() flow to run
// end to end against the real diff/scan logic, driven from
// window.__mock.doc (built by tests/e2e/fixtures.js before this script
// runs). Since commands.html has no visible UI, tests drive it by calling
// the function Office.actions.associate registered, not by clicking
// anything — see window.__mock.registeredActions below.
//
// Document model: each paragraph is { chars: [{ch, revision}], liveRanges }.
// `revision` is null for plain text, or { type: 'Deleted'|'Added', author }
// for tracked text. Nothing is ever physically removed by a "delete" — that
// mirrors Word, where a tracked deletion just marks existing characters,
// it doesn't take them out of the paragraph until the revision is accepted.
// Every Range/TrackedChange handed out is registered in the paragraph's
// liveRanges list so it stays correctly positioned across later
// insertions/removals in the same paragraph, the same way a real Word Range
// auto-adjusts to edits around it.
(function () {
  function shiftForInsert(paragraph, pos, len) {
    paragraph.liveRanges.forEach(r => {
      if (r.start >= pos) r.start += len;
      if (r.end >= pos) r.end += len;
    });
  }

  function shiftForRemove(paragraph, pos, len) {
    paragraph.liveRanges.forEach(r => {
      if (r.start > pos) r.start = Math.max(pos, r.start - len);
      if (r.end > pos) r.end = Math.max(pos, r.end - len);
    });
  }

  function registerRange(paragraph, start, end) {
    const r = { start, end };
    paragraph.liveRanges.push(r);
    return r;
  }

  function scanTrackedChangeGroups(paragraph, scopeStart, scopeEnd) {
    const groups = [];
    let i = scopeStart;
    while (i < scopeEnd) {
      const rev = paragraph.chars[i].revision;
      if (!rev) { i++; continue; }
      let j = i;
      while (
        j < scopeEnd &&
        paragraph.chars[j].revision &&
        paragraph.chars[j].revision.type === rev.type &&
        paragraph.chars[j].revision.author === rev.author
      ) { j++; }
      groups.push({ start: i, end: j, type: rev.type, author: rev.author });
      i = j;
    }
    return groups;
  }

  function makeTrackedChange(paragraph, group) {
    const live = registerRange(paragraph, group.start, group.end);
    return {
      load: function () {},
      get type() { return group.type; },
      get author() { return group.author; },
      get text() { return paragraph.chars.slice(live.start, live.end).map(c => c.ch).join(''); },
      getRange: function (location) {
        const pos = location === 'End' ? live.end : live.start;
        return makeRange(paragraph, registerRange(paragraph, pos, pos));
      },
      reject: function () {
        if (group.type === 'Added') {
          const len = live.end - live.start;
          paragraph.chars.splice(live.start, len);
          shiftForRemove(paragraph, live.start, len);
        } else if (group.type === 'Deleted') {
          for (let k = live.start; k < live.end; k++) paragraph.chars[k].revision = null;
        }
      },
    };
  }

  function makeRange(paragraph, live) {
    return {
      load: function () {},
      getRange: function (location) {
        const pos = location === 'End' ? live.end : live.start;
        return makeRange(paragraph, registerRange(paragraph, pos, pos));
      },
      expandTo: function (other) {
        return makeRange(paragraph, registerRange(paragraph, live.start, other._live.end));
      },
      insertText: function (text) {
        const pos = live.start;
        const newChars = Array.from(text).map(ch => ({
          ch,
          revision: { type: 'Added', author: window.__mock.actingAuthor },
        }));
        paragraph.chars.splice(pos, 0, ...newChars);
        shiftForInsert(paragraph, pos, newChars.length);
        return makeRange(paragraph, registerRange(paragraph, pos, pos + newChars.length));
      },
      delete: function () {
        for (let k = live.start; k < live.end; k++) {
          paragraph.chars[k].revision = { type: 'Deleted', author: window.__mock.actingAuthor };
        }
      },
      search: function (text) {
        const items = [];
        const hay = paragraph.chars.slice(live.start, live.end).map(c => c.ch).join('');
        let idx = 0;
        while (true) {
          const found = hay.indexOf(text, idx);
          if (found === -1) break;
          const start = live.start + found;
          const end = start + text.length;
          items.push(makeRange(paragraph, registerRange(paragraph, start, end)));
          idx = found + Math.max(text.length, 1);
        }
        return { load: function () {}, items: items };
      },
      getTrackedChanges: function () {
        const groups = scanTrackedChangeGroups(paragraph, live.start, live.end);
        return { load: function () {}, items: groups.map(g => makeTrackedChange(paragraph, g)) };
      },
      _live: live,
    };
  }

  function paragraphBoundaryRange(paragraph, location) {
    const pos = location === 'End' ? paragraph.chars.length : 0;
    return makeRange(paragraph, registerRange(paragraph, pos, pos));
  }

  function makeParagraph(paragraph) {
    return {
      getRange: function (location) { return paragraphBoundaryRange(paragraph, location); },
    };
  }

  window.Word = {
    ChangeTrackingMode: { trackAll: 'TrackAll' },
    RangeLocation: { start: 'Start', end: 'End' },
    InsertLocation: { replace: 'Replace' },
    run: function (callback) {
      const doc = window.__mock.doc;
      const ctx = {
        document: {
          changeTrackingMode: null,
          body: {
            // The document-start probe (detectAuthorName) operates on the
            // body directly, not a paragraph, so it needs its own getRange —
            // 'Start' resolves to the very start of the first paragraph.
            getRange: function (location) {
              const paragraph = location === 'End' ? doc.paragraphs[doc.paragraphs.length - 1] : doc.paragraphs[0];
              return paragraphBoundaryRange(paragraph, location);
            },
            paragraphs: {
              load: function () {},
              items: doc.paragraphs.map(makeParagraph),
            },
          },
        },
        sync: function () { return Promise.resolve(); },
      };
      return Promise.resolve().then(function () { return callback(ctx); });
    },
  };

  window.Office = {
    onReady: function (cb) {
      const result = { host: 'Word', platform: 'PC' };
      if (cb) cb(result);
      return Promise.resolve(result);
    },
    actions: {
      associate: function (name, fn) {
        window.__mock.registeredActions = window.__mock.registeredActions || {};
        window.__mock.registeredActions[name] = fn;
      },
    },
    context: {
      ui: {
        displayDialogAsync: function (url, options, callback) {
          window.__mock.lastDialogUrl = url;
          if (callback) callback({ status: 'succeeded' });
        },
      },
    },
  };
})();
