const { defineConfig } = require('@playwright/test');

// Config for the live-Word test in tests/real-office-e2e/ — separate from
// playwright.config.js (which drives the mocked tests/e2e/ suite) because
// this one drives a real local Word install instead of a browser, so it
// needs no browser project. Starts the local HTTPS dev server (used by the
// default LIVE_TARGET=local mode — see word-live.spec.js) so `npm run
// test:live` is a single command; harmless if LIVE_TARGET=prod, just unused.
module.exports = defineConfig({
  testDir: './tests/real-office-e2e',
  fullyParallel: false,
  reporter: 'list',
  timeout: 90000,
  webServer: {
    command: 'node tests/real-office-e2e/local-https-server.js',
    url: 'https://localhost:3000/commands.html',
    ignoreHTTPSErrors: true,
    reuseExistingServer: true,
  },
});
