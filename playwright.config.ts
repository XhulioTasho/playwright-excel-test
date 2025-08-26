import { PlaywrightTestConfig } from '@playwright/test';

const config: PlaywrightTestConfig = {
  testDir: './tests',
  timeout: 60000,
  retries: 2,
  workers: 1,
  use: {
    baseURL: 'https://www.office.com',
    trace: 'on-first-retry',
    video: 'on-first-retry',
    screenshot: 'only-on-failure',
  },
  reporter: [
    ['list'],
    ['html', { outputFolder: 'test-reports' }],
    ['junit', { outputFile: 'test-reports/results.xml' }]
  ],
  projects: [
    {
      name: 'Chrome',
      use: {
        channel: 'chrome',
        browserName: 'chromium',
        launchOptions: {
          args: [
              '--start-maximized',
              '--disable-extensions'
            ],
          headless: false,
        },
        viewport: null,
      },
    }
  ],
  outputDir: 'test-results',
};

export default config;
