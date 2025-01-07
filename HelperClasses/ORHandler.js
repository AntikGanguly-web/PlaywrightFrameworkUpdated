const ExcelJS = require('exceljs');
const { expect } = require('@playwright/test');

class ORHandler
{
  constructor(context, page)
  {
    this.context=context;
    this.page=page;
  }
async getORSheet(action, worksheet)
{
    let obName = [];
    let obRef = [];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("D:/Users/XY59004/OneDrive - Old Mutual/Desktop/ObjectRepo_CS.xlsx");
    const worksheet1 = workbook.getWorksheet(worksheet);
    let i = 0;
    let j = 0;
    let k = 0;
    worksheet1.eachRow((row,rowNumber) =>
    {
          row.eachCell((cell,colNumber) =>
          {
              if(colNumber == 1 && cell.value!="..ObjName")
              {
                obName[i] = cell.value;
                if(obName[i] === "Header")
                {
                  k=i;
                }
                i=i+1;
              }
              if(colNumber == 2 && cell.value!="LocatorValue")
              {
                obRef[j] = cell.value;
                j=j+1;
              }
          }  )
    })
    try
    {
      await expect(this.page.locator(obRef[k])).toBeVisible();
      await this.page.locator(obRef[k]).evaluate((el)=>{
      el.style.border = '3px solid blue';
      });
    }
    catch(error)
    {
      alert(`Exception - ${action}_${worksheet}`);
      await page.waitForTimeout(10000);
    }
   
    const screenshot = await this.captureScreenShot(action,worksheet)    
    await this.page.waitForTimeout(2000);
    return {screenshot,obName,obRef};
}
async captureScreenShot(action,loc)
{
    const currentDate = new Date();
    const timestamp = currentDate.getTime();
    const ssfilepath = 'D:/Users/XY59004/OneDrive - Old Mutual/Desktop/PlaywrightFrameworkUpdated/test-results';
    const screenshotFilePath = `${ssfilepath}/screenshot_${action}_${loc}_${timestamp}.png`;
    await this.page.screenshot({ path: screenshotFilePath })
    return screenshotFilePath
}
async newTab(newPage)
{
    this.page = newPage;
}
}
module.exports = {ORHandler};