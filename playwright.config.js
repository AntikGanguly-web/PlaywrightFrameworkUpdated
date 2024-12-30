const { defineConfig } = require('@playwright/test');
 
 
module.exports = defineConfig({
  testDir: './tests',
  timeout: 700*1000,
  workers: 1,
  fullyParallel: false,
  expect:{
    timeout: 80000
  },
  reporter: [['html',{open: 'always'}]],
  projects: [
    {
      use: {
      browserName: "chromium",
      headless: false,
      viewport: null,
      launchOptions:{
        args:["--start-maximized"],
        slowMo: 200
      },
      screenshot: 'on',
      trace: 'on'
    },
    },
  ],
});