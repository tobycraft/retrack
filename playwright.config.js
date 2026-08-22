const fs = require('fs');
const { defineConfig, devices } = require('@playwright/test');

// Fall back to a pre-installed Chromium (e.g. sandboxed CI images that pin
// their own browser build) only if one is actually present; otherwise let
// Playwright resolve its normally-installed browser.
const preinstalledChromium = process.env.PLAYWRIGHT_CHROMIUM_PATH || '/opt/pw-browsers/chromium';
const launchOptions = fs.existsSync(preinstalledChromium)
  ? { executablePath: preinstalledChromium }
  : {};

module.exports = defineConfig({
  testDir: './tests/e2e',
  fullyParallel: true,
  reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4173',
    launchOptions,
  },
  webServer: {
    command: 'node tests/e2e/static-server.js',
    port: 4173,
    reuseExistingServer: !process.env.CI,
  },
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
    },
  ],
});
