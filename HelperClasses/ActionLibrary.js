const { expect } = require('@playwright/test');
const ExcelJS = require('exceljs');

class ActionLibrary
{
constructor(context, page)
{
    this.context = context;
    this.page = page;
}
async openURL(website)
{
    if(website === "OMSA")
    {
    await this.page.goto("https://secure.advisorweb.snist.dev.oldmutual.co.za/dashboard/sales-dashboard");
    }
}
async enterText(action,obName,loc,value)
{
    const currentDate = new Date();
    const timestamp = currentDate.getTime();
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEditable();
    const ssfilepath = 'D:/Users/XY50035/OneDrive - Old Mutual/Desktop/PlaywrightFramework/test-results';
    const screenshotFilePath = `${ssfilepath}/screenshot_${action}_${obName}_${timestamp}.png`;
    await this.highlightElement(loc);
    let cNum;
    let rNum; 
    let dataVal;
    if(value === "Password@01")
    {
    await this.page.locator(loc).fill(value);
    await this.page.waitForTimeout(2000);
    }
    else
    {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("D:/Users/XY50035/OneDrive - Old Mutual/Desktop/DataSheet.xlsx");
    const worksheet = workbook.getWorksheet('WebEdit');
    worksheet.eachRow((row, rowNumber) =>
    {
    row.eachCell((cell, colNumber) =>
    {
        if(cell.value === obName)
        {
            cNum = colNumber;
        }
        else if(cell.value === value)
        {
            rNum = rowNumber;
        }
    })
    })
    dataVal = worksheet.getCell(rNum,cNum).value;
    await this.page.locator(loc).fill(dataVal.toString());
    await this.page.waitForTimeout(2000);
    }
    await this.page.screenshot({ path: screenshotFilePath });
    return screenshotFilePath;
}
async selectText(action, obName,loc,value)
{
    const currentDate = new Date();
    const timestamp = currentDate.getTime();
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEditable();
    const ssfilepath = 'D:/Users/XY50035/OneDrive - Old Mutual/Desktop/PlaywrightFramework/test-results';
    const screenshotFilePath = `${ssfilepath}/screenshot_${action}_${obName}_${timestamp}.png`;
    await this.highlightElement(loc);
    let cNum;
    let rNum; 
    let dataVal;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("D:/Users/XY50035/OneDrive - Old Mutual/Desktop/DataSheet.xlsx");
    const worksheet = workbook.getWorksheet('WebList');
    worksheet.eachRow((row, rowNumber) =>
    {
    row.eachCell((cell, colNumber) =>
    {
        if(cell.value === obName)
        {
            cNum = colNumber;
        }
        else if(cell.value === value)
        {
            rNum = rowNumber;
        }    
    })
    })
    dataVal = worksheet.getCell(rNum,cNum).value;
    await this.page.locator(loc).selectOption(dataVal.toString());
    await this.page.waitForTimeout(2000);
    await this.page.screenshot({ path: screenshotFilePath });
    return screenshotFilePath;
}
async onClick(action, obName, loc)
{
    const currentDate = new Date();
    const timestamp = currentDate.getTime();
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEnabled();
    const ssfilepath = 'D:/Users/XY50035/OneDrive - Old Mutual/Desktop/PlaywrightFramework/test-results';
    const screenshotFilePath = `${ssfilepath}/screenshot_${action}_${obName}_${timestamp}.png`;
    await this.highlightElement(loc);
    await this.page.locator(loc).click();
    await new Promise(resolve => setTimeout(resolve, 3000));
    await this.page.screenshot({ path: screenshotFilePath });
    return screenshotFilePath;
}
async onClickTabClose(action, obName, loc)
{
    const currentDate = new Date();
    const timestamp = currentDate.getTime();
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEnabled();
    const ssfilepath = 'D:/Users/XY50035/OneDrive - Old Mutual/Desktop/PlaywrightFramework/test-results';
    const screenshotFilePath = `${ssfilepath}/screenshot_${action}_${obName}_${timestamp}.png`;
    await this.highlightElement(loc);
    const [newPage] = await Promise.all([

        this.context.waitForEvent('page'),
        await this.page.locator(loc).click()

    ])
    await this.page.close();
    this.page = newPage;
    const page1 = this.page;
    await this.page.screenshot({ path: screenshotFilePath });
    return {page1,screenshotFilePath};
}
async onClickUploadDoc(action, obName, loc)
{
    const currentDate = new Date();
    const timestamp = currentDate.getTime();
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEnabled();
    const ssfilepath = 'D:/Users/XY50035/OneDrive - Old Mutual/Desktop/PlaywrightFramework/test-results';
    const screenshotFilePath = `${ssfilepath}/screenshot_${action}_${obName}_${timestamp}.png`;
    await this.highlightElement(loc);
    await this.page.locator(loc).click();
    await this.page.waitForTimeout(1000);
    await this.page.locator("[type='file']").setInputFiles("C:/Temp/sticky1.pdf");
    await this.page.waitForTimeout(2000);
    await this.page.screenshot({ path: screenshotFilePath });
    return screenshotFilePath;
}
async checkElement(action, obName, loc)
{
    const currentDate = new Date();
    const timestamp = currentDate.getTime();
    await expect(this.page.locator(loc)).toBeVisible();
    const ssfilepath = 'D:/Users/XY50035/OneDrive - Old Mutual/Desktop/PlaywrightFramework/test-results';
    const screenshotFilePath = `${ssfilepath}/screenshot_${action}_${obName}_${timestamp}.png`;
    await this.highlightElement(loc);
    await this.page.waitForTimeout(1000);
    await this.page.screenshot({ path: screenshotFilePath });
    return screenshotFilePath;
}
async highlightElement(loc)
{
    await expect(this.page.locator(loc)).toBeVisible();
    await this.page.locator(loc).evaluate((el)=>{
        el.style.border = '3px solid blue';
      });
    await this.page.waitForTimeout(1500);
}
}
module.exports = {ActionLibrary};