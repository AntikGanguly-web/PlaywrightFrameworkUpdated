
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
    let screenshot , operation , ORSheet, ServiceRequestNum , screenshotFilePath , CaptureSR
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
for(let k = 0; k<=action.length;k++)
{
    switch (action[k]) {
        case "Navigate":
            try{
                if(locator[k]==="ServiceDashBoardOpen")
                {
                    const newTabOpen = await actLib.tabChange(OR,action[k])
                    page = newTabOpen.page1
                    ORSheet = newTabOpen.ORSheet
                    screenshotFilePath = ORSheet.screenshot;
                }
                else if(locator[k]==="ServiceDashBoardClose")
                {
                    const newTab = await actLib.tabChangeGL(OR,action[k]);
                    page = newTab.page1;
                    ORSheet = newTab.ORSheet;
                    screenshotFilePath = ORSheet.screenshot;
                }
                else if(locator[k]==="CollectionHistory")
                {
                    const newTab = await actLib.tabChangeCollectionHistory(OR,action[k]);
                    page = newTab.page1;
                    ORSheet = newTab.ORSheet;
                    screenshotFilePath = ORSheet.screenshot;
                }
                if(locator[k]==="ConservationDashBoardOpen")
                {
                    const newTabOpen = await actLib.tabChangeConservation(OR,action[k])
                    page = newTabOpen.page1
                    ORSheet = newTabOpen.ORSheet
                    screenshotFilePath = ORSheet.screenshot;
                }
                else if(locator[k]==="ConservationDashBoardClose")
                {
                    const newTab = await actLib.tabChangeConservationGL(OR,action[k]);
                    page = newTab.page1;
                    ORSheet = newTab.ORSheet;
                    screenshotFilePath = ORSheet.screenshot;
                }
                else if(locator[k]==="CustomerAdviceRecord")
                {
                    const carPage = await actLib.carPageAction(action[k],locator[k])
                    for(let i=0;i<carPage.screenshot.length();i++)
                    {
                        operation = await Docx.createP();
                        await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation.addImage(carPage.screenshot[i], {cx: 600, cy: 250})
                    }
                    
                }
                else
                {
                    ORSheet = await OR.getORSheet(action[k], locator[k]);
                    screenshotFilePath = ORSheet.screenshot;
                } 
                operation = await Docx.createP();
                await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation.addImage(screenshotFilePath, {cx: 600, cy: 250})
            }
            catch(error)
            {
                alert(`Exception - ${action[k]}_${locator[k]}`);
                await page.waitForTimeout(10000);
                screenshot = await actLib.captureScreenShot(action[k],locator[k])
                operation = await Docx.createP();
                await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation.addImage(screenshot, {cx: 600, cy: 250})
            }
            break;
            case "GOTO":
                try
                {
                    screenshot = await actLib.openURL(action[k],locator[k]);
                    operation = await Docx.createP();
                    await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation.addImage(screenshot, {cx: 600, cy: 250})
                }
                catch(error)
                {
                    alert(`Exception - ${action[k]}_${locator[k]}`);
                    await page.waitForTimeout(10000);
                    screenshot = await actLib.captureScreenShot(action[k],locator[k])
                    operation = await Docx.createP();
                    await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation.addImage(screenshot, {cx: 600, cy: 250})
                }
            break;
        case "Enter":
            let data;
            try{
            data = locator[k].split("-");
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
            operation = await Docx.createP();
            await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
            await operation.addImage(screenshotFilePath, {cx: 600, cy: 250})
            }
            catch(error)
            {
                alert(`Exception - ${action[k]}_${locator[k]}`);
                await page.waitForTimeout(10000);
                screenshot = await actLib.captureScreenShot(action[k],locator[k])
                operation = await Docx.createP();
                await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation.addImage(screenshot, {cx: 600, cy: 250})

            }
            break;
            case "Select":
                let data1;
                try{
                data1 = locator[k].split("-");
                for(let l = 0; l<=ORSheet.obName.length; l++)
                {
                    if(ORSheet.obName[l]===data1[0])
                    {  
                        screenshotFilePath = await actLib.selectText(action[k],ORSheet.obName[l],ORSheet.obRef[l],data1[1]);
                        break;
                    }
                }
                operation = await Docx.createP();
                await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation.addImage(screenshotFilePath, {cx: 600, cy: 250})
               }
                catch(error)
                {
                    alert(`Exception - ${action[k]}_${locator[k]}`);
                    await page.waitForTimeout(10000);
                    screenshot = await actLib.captureScreenShot(action[k],locator[k])
                    operation = await Docx.createP();
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
                    if(locator[k] === "UploadDoc"||locator[k] === "UploadDocAnother")
                    {
                        const newWindow = await actLib.onClickUploadDoc(action[k], ORSheet.obName[m],ORSheet.obRef[m]);
                        operation = await Docx.createP();
                        await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation.addImage(newWindow, {cx: 600, cy: 250})
                    }
                    if(locator[k] === "IEAddFile")
                    {
                        const newWindow = await actLib.onClickIEAddFile(action[k], ORSheet.obName[m],ORSheet.obRef[m]);
                        operation = await Docx.createP();
                        await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation.addImage(newWindow, {cx: 600, cy: 250})
                    }
                    else
                    {
                        const clickSS = await actLib.onClick(action[k], ORSheet.obName[m], ORSheet.obRef[m]);
                        operation = await Docx.createP();
                        await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation.addImage(clickSS, {cx: 600, cy: 250})
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
                operation = await Docx.createP();
                await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation.addImage(screenshot, {cx: 600, cy: 250})
            }
            break;
        case "Check":
            try{
            for(let n=0;n<ORSheet.obName.length;n++)
            {
                if(ORSheet.obName[n] === locator[k])
                {
                    screenshotFilePath = await actLib.checkElement(action[k], ORSheet.obName[n], ORSheet.obRef[n]);
                }
            }
            operation = await Docx.createP();
            await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
            await operation.addImage(screenshotFilePath, {cx: 600, cy: 250})
            }
            catch(error)
            {
                alert(`Exception - ${action[k]}_${locator[k]}`);
                await page.waitForTimeout(10000);
                screenshot = await actLib.captureScreenShot(action[k],locator[k])
                operation = await Docx.createP();
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
                    operation = await Docx.createP();
                    await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation.addImage(Reference.screenshotFilePath, {cx: 600, cy: 250});        
                }
                catch{
                    alert(`Exception - ${action[k]}_${locator[k]}`);
                    await page.waitForTimeout(10000);
                    screenshot = await actLib.captureScreenShot(action[k],locator[k])
                    operation = await Docx.createP();
                    await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation.addImage(screenshot, {cx: 600, cy: 250})
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
                            }
                        }
                        operation = await Docx.createP();
                        await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation.addImage(CaptureSR.screenshotFilePath, {cx: 600, cy: 250});
                    }
                    catch{
                        alert(`Exception - ${action[k]}_${locator[k]}`);
                        await page.waitForTimeout(10000);
                        screenshot = await actLib.captureScreenShot(action[k],locator[k])
                        operation = await Docx.createP();
                        await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                        await operation.addImage(screenshot, {cx: 600, cy: 250})
                    }
                break;
                case "Hover":
                try{
                let mouseHover;
                for(let n=0;n<ORSheet.obName.length;n++)
                {
                    if(ORSheet.obName[n] === locator[k])
                    {
                        mouseHover = await actLib.hoverElement(action[k], ORSheet.obName[n], ORSheet.obRef[n]);
                    }
                }
                operation = await docx.createP();
                await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                await operation.addImage(mouseHover, {cx: 600, cy: 250})
                }
                catch(error)
                {
                    alert(`Exception - ${action[k]}_${locator[k]}`);
                    await page.waitForTimeout(10000);
                    screenshot = await actLib.captureScreenShot(action[k],locator[k])
                    operation = await Docx.createP();
                    await operation.addText(`Screenshot_${action[k]}_${locator[k]}`);
                    await operation.addImage(screenshot, {cx: 600, cy: 250})
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

