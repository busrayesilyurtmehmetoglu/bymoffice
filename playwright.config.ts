import { defineConfig, devices } from '@playwright/test';

const PORT = Number(process.env.E2E_PORT ?? 4317);
const BASE_URL = `http://127.0.0.1:${PORT}`;
const TRACE = process.env.E2E_TRACE === '1';

export default defineConfig({
  testDir: 'e2e',
  testMatch: '*.spec.ts',
  fullyParallel: true,
  // Every worker is a whole browser. The default (half the cores) has exhausted
  // the memory of a development machine, so the number is fixed and small.
  workers: Number(process.env.E2E_WORKERS ?? 3),
  forbidOnly: !!process.env.CI,
  // A flaky test is a bug to investigate, never something to retry away.
  retries: 0,
  timeout: 30_000,
  reporter: [['list'], ['html', { open: 'never', outputFolder: 'playwright-report' }]],
  outputDir: 'test-results',
  use: {
    baseURL: BASE_URL,
    // A trace records a screenshot of the window after every action of every
    // test, kept or not. On macOS WebKit takes each one through the window
    // server, and thousands of them in one run have brought the window server
    // down, and the whole desktop session with it. Nothing is captured unless
    // asked for: E2E_TRACE=1 with the one spec being looked into.
    trace: TRACE ? 'retain-on-failure' : 'off',
    screenshot: TRACE ? 'only-on-failure' : 'off',
  },
  webServer: {
    command: 'node e2e/server.mjs',
    url: `${BASE_URL}/e2e/pages/script.html`,
    reuseExistingServer: false,
    env: { E2E_PORT: String(PORT) },
  },
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    { name: 'firefox', use: { ...devices['Desktop Firefox'] } },
    { name: 'webkit', use: { ...devices['Desktop Safari'] } },
    { name: 'iphone', use: { ...devices['iPhone 14'] } },
    { name: 'pixel', use: { ...devices['Pixel 7'] } },
  ],
});
