const { expect } = require('@playwright/test');
const ExcelJS = require('exceljs');

let screenshot

class ActionLibrary
{
constructor(context, page)
{
    this.context = context;
    this.page = page;
}
async openURL(action,website)
{
    if(website === "OMSA")
    {
    await this.page.goto("https://secure.advisorweb.snist.dev.oldmutual.co.za/dashboard/sales-dashboard");
    }
    screenshot = this.captureScreenShot (action,website)
    return screenshot
}
async enterText(action,obName,loc,value)
{
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEditable();
    this.page.locator(loc).clear()
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
    await workbook.xlsx.readFile("D:/Users/XY59004/OneDrive - Old Mutual/Desktop/DataSheet.xlsx");
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
    screenshot = this.captureScreenShot (action,obName)
    return screenshot;
}
async selectText(action, obName,loc,value)
{
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEditable();
    await this.highlightElement(loc);
    let cNum;
    let rNum; 
    let dataVal;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("D:/Users/XY59004/OneDrive - Old Mutual/Desktop/DataSheet.xlsx");
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
    screenshot = this.captureScreenShot(action,obName)
    return screenshot;
}
async onClick(action, obName, loc)
{
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEnabled();
    await this.highlightElement(loc);
    await this.page.locator(loc).click();
    await new Promise(resolve => setTimeout(resolve, 3000));
    screenshot = this.captureScreenShot (action,obName)
    return screenshot;
}
async tabChange(OR,action)
{
    await this.page.locator(".x-icon-menu").click();
    await this.page.getByText('EXISTING BUSINESS SERVICES DASHBOARD').click();
    const [newPage] = await Promise.all(
                                        [this.context.waitForEvent('page'),
                                        await this.page.getByText('OM Protect/Savings & Income').click()]
                                        );
    await this.page.close()
    this.page = newPage
    const page1 = this.page
    await OR.newTab(page1);
    let ORSheet = await OR.getORSheet(action,"ServiceDashboard")
    return {ORSheet,page1}
}
async tabChangeGL(OR,action)
{
    await this.page.locator(".x-icon-menu").click();
    await this.page.getByText('EXISTING BUSINESS SERVICES DASHBOARD').click();
    const [newPage] = await Promise.all(
                                        [this.context.waitForEvent('page'),
                                        await this.page.getByText('Greenlight').click()]
                                        );
    await this.page.close()
    this.page = newPage
    const page1 = this.page
    await OR.newTab(page1);
    let ORSheet = await OR.getORSheet(action,"ServiceDashboard")
    return {ORSheet,page1}
}
async tabChangeCollectionHistory(OR,action)
{
    await this.page.locator(".x-icon-menu").click();
    const [newPage] = await Promise.all(
                                        [this.context.waitForEvent('page'),
                                        await this.page.getByText('COLLECTION HISTORY').click()]
                                        );
    await this.page.close()
    this.page = newPage
    const page1 = this.page
    await OR.newTab(page1);
    let ORSheet = await OR.getORSheet(action,"CollectionHistory")
    return {ORSheet,page1}
}
async tabChangeConservation(OR,action)
{
    await this.page.locator(".x-icon-menu").click();
    await this.page.getByText('CONSERVATION').click();
    const [newPage] = await Promise.all(
                                        [this.context.waitForEvent('page'),
                                        await this.page.getByText('OM Protect/Savings & Income').click()]
                                        );
    await this.page.close()
    this.page = newPage
    const page1 = this.page
    await OR.newTab(page1);
    let ORSheet = await OR.getORSheet(action,"ConservationDashboard")
    return {ORSheet,page1}
}
async tabChangeConservationGL(OR,action)
{
    await this.page.locator(".x-icon-menu").click();
    await this.page.getByText('CONSERVATION').click();
    const [newPage] = await Promise.all(
                                        [this.context.waitForEvent('page'),
                                        await this.page.getByText('Greenlight').click()]
                                        );
    await this.page.close()
    this.page = newPage
    const page1 = this.page
    await OR.newTab(page1);
    let ORSheet = await OR.getORSheet(action,"ConservationDashboard")
    return {ORSheet,page1}
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
async onClickUploadDoc(action, obName, loc)
{
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEnabled();
    await this.highlightElement(loc);
    await this.page.locator(loc).click();
    await this.page.waitForTimeout(1000);
    await this.page.locator("[type='file']").setInputFiles("C:/Temp/sticky1.pdf");
    await this.page.waitForTimeout(1000);
    screenshot = this.captureScreenShot (action,obName)
    return screenshot;
}
async onClickIEAddFile(action, obName, loc)
{
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEnabled();
    await this.highlightElement(loc);
    await this.page.waitForTimeout(1000);
    await this.page.locator(loc).setInputFiles("C:/Temp/sticky1.pdf");
    await this.page.waitForTimeout(1000);
    screenshot = this.captureScreenShot (action,obName)
    return screenshot;
}
async checkElement(action, obName, loc)
{
    await expect(this.page.locator(loc)).toBeVisible();
    await this.highlightElement(loc);
    await this.page.waitForTimeout(1000);
    screenshot = this.captureScreenShot (action,obName)
    return screenshot;
}
async highlightElement(loc)
{
    await expect(this.page.locator(loc)).toBeVisible();
    await this.page.locator(loc).evaluate((el)=>{
        el.style.border = '3px solid blue';
      });
    await this.page.waitForTimeout(500);
}
async carPageAction(action,loc)
{
    ORSheet = await OR.getORSheet(locator[k])
    let screenshot = []
    await this.highlightElement("text='FINANCIAL PLAN'")
    screenshot [0]= await this.captureScreenShot("Financial","Tab")
    await this.highlightElement("text='ADVICE DETAILS'")
    screenshot [1]= await this.captureScreenShot("Advice Details","Tab")
    await this.highlightElement("text='TRANSACTION(S)'")
    screenshot [2]= await this.captureScreenShot("Transaction","Tab")
    await this.page.locator("//h3[contains(text(),'ADVICE DETAILS')]//following::div[8]").click()
    screenshot [3]= await this.captureScreenShot("Question1","Yes")
    await this.page.locator("//h3[contains(text(),'ADVICE DETAILS')]//following::div[16]").click()
    screenshot [4]= await this.captureScreenShot("Question2","Yes")
    await this.highlightElement("text='IMPLICATION & RECOMMENDATION '")
    screenshot [5]= await this.captureScreenShot("Implication","Tab")
    await this.page.locator("[placeholder*='Type the <Implication of Transaction(s)>']").fill("Implication")
    screenshot [6]= await this.captureScreenShot("Implication","Text")
    await this.page.locator("[placeholder*='Type <Reason for Recommendation>']").fill("Reason For Recommendation")
    screenshot [7]= await this.captureScreenShot("Recommendation","Text")
    await this.page.locator(".next-btn").click()
    screenshot [8]= await this.captureScreenShot("Save","Button")
    await this.page.locator(".next-btn").click()
    screenshot [9]= await this.captureScreenShot("Next","Button")
        
    return {ORSheet,screenshot}
}
async captureSr(action, obName, loc)
{  
    await expect(this.page.locator(loc)).toBeVisible();
    await this.page.waitForTimeout(1000);
    await this.highlightElement(loc);
    let SRNum = await this.page.locator(loc).innerText();
    SRNum = SRNum.split(":");
    SRNum = SRNum[1];
    SRNum = SRNum.trim();
    console.log(SRNum);
    screenshot = this.captureScreenShot (action,obName)
    return {screenshot,SRNum};
}
 
async enterData(action, obName, loc, SRNum)
{  
    await expect(this.page.locator(loc)).toBeVisible();
    await this.page.waitForTimeout(1000);
    await this.highlightElement(loc);
    await this.page.locator(loc).fill(SRNum);
    screenshot = this.captureScreenShot (action,obName)
    return screenshot;
}
async hoverElement(action, obName, loc)
{
    await expect(this.page.locator(loc)).toBeVisible();
    await expect(this.page.locator(loc)).toBeEnabled();
    await this.highlightElement(loc);
    await this.page.locator(loc).hover();
    await this.page.waitForTimeout(1000);
    screenshot = this.captureScreenShot (action,obName)
    return screenshot;
}
 
}
module.exports = {ActionLibrary};
