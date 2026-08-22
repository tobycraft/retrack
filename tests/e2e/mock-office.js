// Stand-in for Microsoft's office.js, served to the browser in place of the
// real CDN script during tests. Real Word/Office.js can't be automated in
// CI, so this mock reproduces just enough of the Word JS API surface
// (Office.onReady, Word.run, Range) for taskpane.html's apply() flow to run
// end to end against the real diff/OOXML code, driven from window.__mock
// (set up by tests/e2e/fixtures.js before this script runs).
(function () {
  window.Word = {
    InsertLocation: { replace: 'Replace' },
    run: function (callback) {
      var range = {
        text: window.__mock.selectionText,
        load: function () {},
        getOoxml: function () {
          return { value: window.__mock.selectionOoxml };
        },
        insertOoxml: function (ooxml, location) {
          window.__mock.lastOoxml = ooxml;
          window.__mock.lastInsertLocation = location;
        },
      };
      var ctx = {
        document: {
          getSelection: function () { return range; },
        },
        sync: function () { return Promise.resolve(); },
      };
      return Promise.resolve().then(function () { return callback(ctx); });
    },
  };

  window.Office = {
    onReady: function (cb) {
      var result = { host: 'Word', platform: 'PC' };
      if (cb) cb(result);
      return Promise.resolve(result);
    },
    context: {
      userProfile: { displayName: window.__mock.author },
    },
  };
})();
