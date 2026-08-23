const { defineConfig } = require('@playwright/test');

// Config for the live-Word test in tests/real-office-e2e/ — separate from
// playwright.config.js (which drives the mocked tests/e2e/ suite) because
// this one drives a real local Word install instead of a browser, so it
// needs no webServer and no browser project.
module.exports = defineConfig({
  testDir: './tests/real-office-e2e',
  fullyParallel: false,
  reporter: 'list',
  timeout: 60000,
});
