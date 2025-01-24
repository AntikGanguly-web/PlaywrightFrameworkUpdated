const { defineConfig } = require('@playwright/test');
 
 
module.exports = defineConfig({
  testDir: './tests',
  timeout: 700*1000,
  workers: 1,
  fullyParallel: false,
  expect:{
    timeout: 50000
  },
  //reporter: [['html',{open: 'always'}]],
  reporter: [['junit',{outputFile:'results.xml'}]],
  projects: [
    {
      use: {
      //browserName: "chromium",
      channel:"msedge",
      headless: false,
      viewport: null,
      launchOptions:{
        args:["--start-maximized"],
        //slowMo: 50
      },
      screenshot: 'on',
      trace: 'on'
    },
    },
  ],
});
