// STUB — not implemented. Real end-to-end test against actual Word, using
// @microsoft/office-addin-test-framework, as a follow-up to the mocked
// suite in tests/e2e/.
//
// Why this doesn't exist yet:
// tests/e2e/ mocks the Office.js/Word host boundary and runs in a plain
// browser, so it can verify the add-in's own logic (tokenizer, diff,
// paragraph scan/rewrite loop) but proves nothing about real sideloading,
// the ribbon, or getTrackedChanges()/TrackedChange.reject() actually
// behaving as mocked in a live document.
// @microsoft/office-addin-test-framework drives the real Office app (or
// Office on the web via a real Microsoft 365 tenant) instead of stubbing
// it — that needs infrastructure this repo/sandbox doesn't have:
//   - A Windows or Mac runner with a licensed Word desktop install, OR
//   - Microsoft 365 dev tenant credentials to drive Word on the web
//   - The manifest actually reachable (deployed GitHub Pages, or a local
//     HTTPS server matching the AppDomains in manifest.xml)
//   - A seeded .docx fixture with Track Changes enabled and existing
//     delete+insert tracked-change pairs already in it (e.g. produced by
//     retyping over a selection), authored by a known reviewer name
//
// When that infra exists, implement this by:
//   1. `npm install --save-dev @microsoft/office-addin-test-framework`
//   2. Sideload manifest.xml into the test Word instance / tenant, with the
//      fixture doc's Word user name already set to the fixture's reviewer
//      name (File > Options > General > User name) — there's no field to
//      type it into anymore
//   3. Open the fixture doc, click Minimize My Changes on the ribbon
//   4. Re-open/read the saved document.xml and assert the real <w:ins>/
//      <w:del> runs match the same core scenarios covered in
//      tests/e2e/commands.spec.js (single pair minimized, multiple pairs
//      across paragraphs, a pair skipped for another author, no matching
//      pairs, author auto-detection) — this time verifying Word actually
//      applied them, not just that the add-in's mocked model was updated
//      correctly.
//
// See the "Testing" section in README.md and CLAUDE.md for the current
// mocked-host approach and why it was chosen over this one for now.
