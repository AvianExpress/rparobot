const puppeteer = require('puppeteer');
const fs=require('fs');
const xlsx=require('xlsx');
const clipboardy=require('clipboardy');




let mockup=[];

(async () =>{

    let vb=xlsx.readFile("test.xlsx",{cellDates:true});
    let vs = vb.Sheets['Втормет Красноярск 27.04'];
    let range = xlsx.utils.decode_range(vs['!ref']);
    let i=5;
    let address = {c:2, r:i};
    let cell_ref = xlsx.utils.encode_cell(address);
    let Sheet1C = vs[cell_ref];
    if (Sheet1C && (String(Sheet1C.v).length === 8)){
    let Sheet1Value = String(Sheet1C.v);
    clipboardy.writeSync(Sheet1Value);
    console.log(clipboardy.readSync());
    

    const browser = await puppeteer.launch({
        headless:false,
        slowmo:200,
        ignoreHTTPSErrors:true,
        args: ['--start-maximized'],
        executablePath: 'C:/Program Files/Google/Chrome/Application/chrome'
        
    });
    const page=await browser.newPage();
    
    const retry = (fn, ms)=> new Promise (resolve => {
        fn()
            .then(resolve)
            .catch(() => {
                setTimeout(()=>{
                    console.log('retrying...');
                    retry(fn, ms).then(resolve);
                }, ms);
            })
        });
    const retryHard = (fn, ms)=> new Promise (resolve => {
    fn()
        .then(resolve)
        .catch(() => {
            setTimeout(()=>{
                page.reload();
                console.log('retrying hard...');
                retryHard(fn, ms).then(resolve);
            }, ms);
        })
    });
    await page.setViewport({
        width: 1920,
        height:1080
    });
    // page.on("pageerror", function(err) {  
    //     theTempValue = err.toString();
    //     console.log("Page error: " + theTempValue); 
    // });
    await page.setDefaultNavigationTimeout(100000);


    await retryHard(() => page.goto('http://10.248.35.9:8092/WebShell/',{timeout:10000}), 10000);
    // await page.goto('http://10.248.35.9:8092/WebShell/');
    // page.on('console', msg => {
    //     console.log(msg.text());
    // });
   
    await page.waitForTimeout(2000);
    await retryHard(()=> page.waitForSelector('.dialog-auth-user-name',{timeout:5000}), 5000); 
    await page.type('.dialog-auth-user-name', 'Осина_ЮВ');
    await page.type('.dialog-auth-user-password', 'Etran222');
    await page.click('.button'); 
    await page.waitForSelector(':nth-child(16) > .shell-menu-item-img',{timeout:0});
    // await page.click(':nth-child(16) > .shell-menu-item-img');
    await retryHard(() => page.goto('http://10.248.35.9:8092/WebShell/%D0%97%D0%B0%D1%8F%D0%B2%D0%BA%D0%B0_%D0%BD%D0%B0_%D0%B3%D1%80%D1%83%D0%B7%D0%BE%D0%BF%D0%B5%D1%80%D0%B5%D0%B2%D0%BE%D0%B7%D0%BA%D1%83',{timeout:5000}), 5000);
    // await page.goto('http://10.248.35.9:8092/WebShell/%D0%97%D0%B0%D1%8F%D0%B2%D0%BA%D0%B0_%D0%BD%D0%B0_%D0%B3%D1%80%D1%83%D0%B7%D0%BE%D0%BF%D0%B5%D1%80%D0%B5%D0%B2%D0%BE%D0%B7%D0%BA%D1%83');
    await page.waitForTimeout(4000);
    // await page.reload();
    await retryHard(() => page.waitForSelector('.js-button-find-document',{timeout:10000}), 10000);
    await page.click('.js-button-find-document');
    await page.click(':nth-child(2) > .doc-search-dialog__radio-caption');
    await page.keyboard.down('Control',{delay: 250});
    await page.keyboard.down('V');
    await page.waitForSelector('.dialog__body > .dialog-data > .doc-search-dialog__container > .doc-search-dialog__buttons > .button:nth-child(1)');
    await page.click('.dialog__body > .dialog-data > .doc-search-dialog__container > .doc-search-dialog__buttons > .button:nth-child(1)');
    await  retry(() => page.waitForSelector('.ui-grid-summary-page > .ui-grid-row > .first-cell > .cell-content'), 5000);
    const num = await page.$eval('.ui-grid-summary-page > .ui-grid-row > .first-cell > .cell-content', ele => ele.textContent);
    console.log(num);
    num1=num;
    num1++;
    num1++;
    console.log(num1);
    await page.click(':nth-child('+ num1 +') > .first-cell > .cell-content', {clickCount: 2 });
    await page.waitForSelector('.xm-object > .nav > .xm-tabsheet-caption:nth-child(10) > .nav-link > .xm-tabsheet-caption-text',{timeout:0});
    await page.click('.xm-object > .nav > .xm-tabsheet-caption:nth-child(10) > .nav-link > .xm-tabsheet-caption-text', {timeout: 0});
    await page.waitForTimeout(4000);
    await page.waitForSelector('.xm-container-data > :nth-child(5) > .btn-caption');
    await page.click('.xm-container-data > :nth-child(5) > .btn-caption');
    
    // await page.waitForSelector('.xm-container-data > .xm-object > .xm-container-data > .xm-object:nth-child(5) > .btn-caption');
    // await page.click('.xm-container-data > .xm-object > .xm-container-data > .xm-object:nth-child(5) > .btn-caption');
    
    // await page.waitForSelector('.dialog-data > .xm-object > .xm-container-data > .xm-grid-container-wrapper > .xm-grid-container > .ui-grid > .ui-grid-content > .ui-grid-body > .ui-grid-body-scrollable > .ui-grid-page > :nth-child(3) > :nth-child(4) > .cell-content');
    // await page.waitForTimeout(4000);
    // await page.click('.dialog-data > .xm-object > .xm-container-data > .xm-grid-container-wrapper > .xm-grid-container > .ui-grid > .ui-grid-content > .ui-grid-body > .ui-grid-body-scrollable > .ui-grid-page > :nth-child(3) > :nth-child(4) > .cell-content');
    // await page.type('.dialog-data > .xm-object > .xm-container-data > .xm-grid-container-wrapper > .xm-grid-container > .ui-grid', 'ArrowUp');
    await page.waitForSelector('.xm-object:nth-child(1) > .xm-container-data:nth-child(2) > .xm-grid-container-wrapper:nth-child(1) > .xm-grid-container:nth-child(2) > .ui-grid:nth-child(2) > .ui-grid-content:nth-child(6) > .ui-grid-body:nth-child(3) .ui-grid-row:nth-child(3) > .ui-grid-cell:nth-child(4) > .cell-content:nth-child(1)');
    await page.click('.xm-object:nth-child(1) > .xm-container-data:nth-child(2) > .xm-grid-container-wrapper:nth-child(1) > .xm-grid-container:nth-child(2) > .ui-grid:nth-child(2) > .ui-grid-content:nth-child(6) > .ui-grid-body:nth-child(3) .ui-grid-row:nth-child(3) > .ui-grid-cell:nth-child(4) > .cell-content:nth-child(1)');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.down('Control');
    await page.keyboard.press('C');
    await page.keyboard.up('Control');
    for (i=0; mockup[i-1]!=clipboardy.readSync(); i++){
    mockup.push(clipboardy.readSync());
    // await page.type('.dialog-data > .xm-object > .xm-container-data > .xm-grid-container-wrapper > .xm-grid-container > .ui-grid', 'ArrowDownArrowDown');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.down('Control');
    await page.keyboard.press('C');
    await page.keyboard.up('Control');
    }
    console.log(mockup);
    await page.waitForTimeout(10000);
    browser.close();
    
}
})();