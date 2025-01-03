
const ExcelJS = require('exceljs');
const { test } = require('@playwright/test');
const { alert } = require('vscode-websocket-alerts');
const {ORHandler} = require('../HelperClasses/ORHandler')
const {ActionLibrary} = require('../HelperClasses/ActionLibrary');
const fs = require('fs');
const Officegen = require('officegen');
const Docx = Officegen('docx');
const scenarioSets = JSON.parse(JSON.stringify(require('../utils/ScenarioNumbers.json')));

for(const scenarioSet of scenarioSets)
{
if(scenarioSet.RunFlag === "Y")
{
test(`Scenario Executing - ${scenarioSet.ScenarioName}`,async ({browser})=>
{
    const context = await browser.newContext();
    let page = await context.newPage();
    const OR = new ORHandler(context, page);
    const actLib = new ActionLibrary(context, page);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("D:/Users/XY59004/OneDrive - Old Mutual/Desktop/Keywords.xlsx");
    const worksheet = workbook.getWorksheet('Keyword');
    let screenshot , operation
    let i = 0;
    let j = 0;
    let action = [];
    let locator = [];
    let rNum;
    worksheet.eachRow((row, rowNumber) =>
    {
    row.eachCell((cell, colNumber) =>
    {
        if(cell.value === scenarioSet.ScenarioName)
        {
            rNum = rowNumber;
        }   
    })
    })
    const scenario = worksheet.getRow(rNum).values;
    for(let a=2; a<scenario.length; a++)
    {
        const keyword = scenario[a].split("_");
        action[i] = keyword[0];
        locator[j] = keyword[1];
        i = i+1;
        j = j+1;
    }
let ORSheet;
for(let k = 0; k<=action.length;k++)
{
    switch (action[k]) {
        case "Navigate":
            try{
                let screenshotFilePath4;
                if(locator[k]==="ServiceDashBoardOpen")
                {
                    const newTabOpen = await actLib.tabChange(OR,action[k])
                    page = newTabOpen.page1
                    ORSheet = newTabOpen.ORSheet
                    screenshotFilePath4 = ORSheet.screenshot;
                }
                else if(locator[k]==="ServiceDashBoardClose")
                {
                    const newTab = await actLib.tabChangeGL(OR,action[k]);
                    page = newTab.page1;
                    ORSheet = newTab.ORSheet;
                    screenshotFilePath4 = ORSheet.screenshot;
                }
                else if(locator[k]==="CustomerAdviceRecord")
                {
                    const carPage = await actLib.carPageAction(action[k],locator[k])
                    for(let i=0;i<carPage.screenshot.length();i++)
                    {
                        let operation = await Docx.createP();
                        await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation.addImage(carPage.screenshot[i], {cx: 600, cy: 250})
                    }
                    
                }
                else
                {
                    ORSheet = await OR.getORSheet(action[k], locator[k]);
                    screenshotFilePath4 = ORSheet.screenshot;
                } 
                let operation5 = await Docx.createP();
                await operation5.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation5.addImage(screenshotFilePath4, {cx: 600, cy: 250})
            }
            catch(error)
            {
                alert(`Exception - ${action[k]}_${locator[k]}`);
                await page.waitForTimeout(10000);
                screenshot = await actLib.captureScreenShot(action[k],locator[k])
                let operation = await Docx.createP();
                await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation.addImage(screenshot, {cx: 600, cy: 250})
            }
            break;
            case "GOTO":
                try
                {
                    screenshot = await actLib.openURL(action[k],locator[k]);
                    let operation = await Docx.createP();
                    await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation.addImage(newTab.screenshotFilePath, {cx: 600, cy: 250})
                }
                catch(error)
                {
                    alert(`Exception - ${action[k]}_${locator[k]}`);
                    await page.waitForTimeout(10000);
                    screenshot = await actLib.captureScreenShot(action[k],locator[k])
                    let operation = await Docx.createP();
                    await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation.addImage(screenshot, {cx: 600, cy: 250})
                }
            break;
        case "Enter":
            let data;
            try{
            data = locator[k].split("-");
            let screenshotFilePath;
            for(let l = 0; l<=ORSheet.obName.length; l++)
            {
                if(ORSheet.obName[l]===data[0])
                {
                    if(locator[k] === "Password")
                    {
                    screenshotFilePath = await actLib.enterText(action[k],ORSheet.obName[l],ORSheet.obRef[l], "Password@01");
                    }
                    else
                    {  
                    screenshotFilePath = await actLib.enterText(action[k],ORSheet.obName[l],ORSheet.obRef[l],data[1]);
                    }
                    break;
                }
            }
            let operation = await Docx.createP();
            await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
            await operation.addImage(screenshotFilePath, {cx: 600, cy: 250})
            }
            catch(error)
            {
                alert(`Exception - ${action[k]}_${locator[k]}`);
                await page.waitForTimeout(10000);
                screenshot = await actLib.captureScreenShot(action[k],locator[k])
                let operation = await Docx.createP();
                await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation.addImage(screenshot, {cx: 600, cy: 250})

            }
            break;
            case "Select":
                let data1;
                try{
                data1 = locator[k].split("-");
                let screenshotFilePath1;
                for(let l = 0; l<=ORSheet.obName.length; l++)
                {
                    if(ORSheet.obName[l]===data1[0])
                    {  
                        screenshotFilePath1 = await actLib.selectText(action[k],ORSheet.obName[l],ORSheet.obRef[l],data1[1]);
                        break;
                    }
                }
                let operation1 = await Docx.createP();
                await operation1.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation1.addImage(screenshotFilePath1, {cx: 600, cy: 250})
                }
                catch(error)
                {
                    alert(`Exception - ${action[k]}_${locator[k]}`);
                    await page.waitForTimeout(10000);
                    screenshot = await actLib.captureScreenShot(action[k],locator[k])
                    let operation = await Docx.createP();
                    await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation.addImage(screenshot, {cx: 600, cy: 250})
                }
                break;
        case "Click":
            try{
            for(let m = 0; m<=ORSheet.obName.length; m++)
            {
                if(ORSheet.obName[m]===locator[k])
                {
                    /*if(locator[k] === "ProtectSavingsIncome" || locator[k] === "GreenLight")
                    {
                        const newTab = await actLib.onClickTabClose(action[k], ORSheet.obName[m],ORSheet.obRef[m]);
                        page = newTab.page1;
                        await OR.newTab(newTab.page1);
                        let operation2 = await Docx.createP();
                        await operation2.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation2.addImage(newTab.screenshotFilePath, {cx: 600, cy: 250})
                    }*/
                    if(locator[k] === "UploadDoc")
                    {
                        const newTab = await actLib.onClickUploadDoc(action[k], ORSheet.obName[m],ORSheet.obRef[m]);
                        let operation2 = await Docx.createP();
                        await operation2.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation2.addImage(newTab, {cx: 600, cy: 250})
                    }
                    else
                    {
                        const newTab = await actLib.onClick(action[k], ORSheet.obName[m], ORSheet.obRef[m]);
                        let operation2 = await Docx.createP();
                        await operation2.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation2.addImage(newTab, {cx: 600, cy: 250})
                    }
                    break;
                }
            }
            }
            catch(error)
            {
                alert(`Exception - ${action[k]}_${locator[k]}`);
                await page.waitForTimeout(10000);
                screenshot = await actLib.captureScreenShot(action[k],locator[k])
                let operation = await Docx.createP();
                await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation.addImage(screenshot, {cx: 600, cy: 250})
            }
            break;
        case "Check":
            try{
            let screenshotFilePath3;
            for(let n=0;n<ORSheet.obName.length;n++)
            {
                if(ORSheet.obName[n] === locator[k])
                {
                    screenshotFilePath3 = await actLib.checkElement(action[k], ORSheet.obName[n], ORSheet.obRef[n]);
                }
            }
            let operation3 = await Docx.createP();
            await operation3.addText(`Screenshot_${action[k]}_${locator[k]}`);
            await operation3.addImage(screenshotFilePath3, {cx: 600, cy: 250})
            }
            catch(error)
            {
                alert(`Exception - ${action[k]}_${locator[k]}`);
                await page.waitForTimeout(10000);
                screenshot = await actLib.captureScreenShot(action[k],locator[k])
                let operation = await Docx.createP();
                await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation.addImage(screenshot, {cx: 600, cy: 250})
            }
            break;
            case "Capture":
                try{
                    let screenshotFilePath2;
                    for(let n=0;n<ORSheet.obName.length;n++)
                    {
                        if(ORSheet.obName[n] === locator[k])
                        {  
                            screenshotFilePath2 = await actLib.captureSr(action[k],ORSheet.obName[n],ORSheet.obRef[n]);
                        }
                    }
                    let operation4 = await Docx.createP();
                    await operation4.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation4.addImage(screenshotFilePath2, {cx: 600, cy: 250})
                    }
                    catch(error)
                    {
                        alert(`Exception - ${action[k]}_${locator[k]}`);
                        await page.waitForTimeout(10000);
                        screenshot = await actLib.captureScreenShot(action[k],locator[k])
                        let operation = await Docx.createP();
                        await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation.addImage(screenshot, {cx: 600, cy: 250})
                    }
                break;
                case "SRNumber":
                let Reference;
                try{
                    for(let n=0;n<ORSheet.obName.length;n++)
                    {
                        if(ORSheet.obName[n] === locator[k])
                        {  
                            Reference = await actLib.enterData(action[k],ORSheet.obName[n],ORSheet.obRef[n],ServiceRequestNum);                            
                        }
                    }
                    let operation5 = await docx.createP();
                    await operation5.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation5.addImage(Reference.screenshotFilePath, {cx: 600, cy: 250});        
                }
                catch{
                    alert(`Exception - ${action[k]}_${locator[k]}`);
                    await page.waitForTimeout(10000);
                }
                break;
                case "Capture":
                    try{
                    for(let n=0;n<ORSheet.obName.length;n++)
                    {
                        if(ORSheet.obName[n] === locator[k])
                        {  
                            CaptureSR = await actLib.captureSr(action[k],ORSheet.obName[n],ORSheet.obRef[n]);
                            ServiceRequestNum= CaptureSR.SRNum;
                            console.log(CaptureSR.SRNum);
                        }
                    }
                    let operation4 = await docx.createP();
                    await operation4.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation4.addImage(CaptureSR.screenshotFilePath, {cx: 600, cy: 250});}
                    catch{
                        alert(`Exception - ${action[k]}_${locator[k]}`);
                        await page.waitForTimeout(10000);
                    }
                    break;
            default:
            break;
    }
}
        const out = fs.createWriteStream(`D:/Users/XY59004/OneDrive - Old Mutual/Desktop/PlaywrightFrameworkUpdated/test-results/${scenarioSet.ScenarioName}.docx`);
        Docx.generate(out);
        await page.close();
})
}
}
