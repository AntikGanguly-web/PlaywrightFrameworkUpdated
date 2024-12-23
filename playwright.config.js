const { defineConfig } = require('@playwright/test');
 
 
module.exports = defineConfig({
  testDir: './tests',
  timeout: 300*1000,
  workers: 1,
  fullyParallel: false,
  expect:{
    timeout: 30000
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
        slowMo: 2000
      },
      screenshot: 'on',
      trace: 'on'
    },
    },
  ],
});