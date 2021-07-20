const puppeteer = require('puppeteer');
const fs=require('fs');
const xlsx=require('xlsx');
const clipboardy=require('clipboardy');




(async () =>{
    let vb=xlsx.readFile("test.xlsx",{cellDates:true});
    //У них разные имена книг
    let vs = vb.Sheets['Втормет Красноярск 27.04'];
    let range = xlsx.utils.decode_range(vs['!ref']);
    let i =0;
    let address = {c:0, r:0};
    let addressNСarriage = {c:0, r:0};
    let addressGU = {c:0, r:0};
    let cell_ref = xlsx.utils.encode_cell(address);
    for (i=0, l = range.e.c; i<l; i+=1){
        address = {c:i, r:0};
        cell_ref = xlsx.utils.encode_cell(address);
        let Sheet1C = vs[cell_ref];
        if (Sheet1C.v === "№ ГУ-12"){
            addressGU = address;
        }
        if (Sheet1C.v === "Номера вагонов"){
            addressNСarriage = address;
        }
    }
    const browser = await puppeteer.launch({
        headless:false,
        slowmo:20,
        ignoreHTTPSErrors:true,
        args: ['--start-maximized'],
        executablePath: 'C:/Program Files/Google/Chrome/Application/chrome'
    });
    const page=await browser.newPage();
    await page.setViewport({
        width: 1920,
        height:1080
    });
    await page.goto('http://10.248.35.9:8092/WebShell/', {
        waitUntil: 'load', timeout: 0}); 
    await page.waitForSelector('.dialog-auth-user-name') 
    await page.type('.dialog-auth-user-name', 'Осина_ЮВ');
    await page.type('.dialog-auth-user-password', 'Etran222');
    //await page.waitForTimeout(2000);
    await page.click('.button', {
        networkIdleTimeout: 5000, waitUntil: 'networkidle', waitUntil: 'load', timeout: 0});
    //await page.waitForTimeout(4000);
    const page2=await browser.newPage();
    await page2.setViewport({
        width: 1920,
        height:1080
    });
    await page2.goto('http://10.248.35.9:8092/WebShell/%D0%97%D0%B0%D1%8F%D0%B2%D0%BA%D0%B0_%D0%BD%D0%B0_%D0%B3%D1%80%D1%83%D0%B7%D0%BE%D0%BF%D0%B5%D1%80%D0%B5%D0%B2%D0%BE%D0%B7%D0%BA%D1%83', {
        waitUntil: 'networkidle', waitUntil: 'load', timeout: 0});
//            await page2.waitForTimeout(4000);
    // await page2.reload();
    // await page2.waitForNavigation();

            //await page2.waitForTimeout(4000);
    //await page2.waitForTimeout(4000);
    for (i=0, l = range.e.r; i<l; i+=1)
    {
        addressGU.r =i;
        cell_ref = xlsx.utils.encode_cell(addressGU);
        Sheet1C = vs[cell_ref];
        if (Sheet1C && (String(Sheet1C.v).length === 8)){
            console.log(Sheet1C.v);
            clipboardy.writeSync(String(Sheet1C.v));
            await page2.waitForSelector('.js-button-find-document',{timeout: 60000}).then(()=>page2.click('.js-button-find-document'));
//            await page2.waitForTimeout(4000);
            await page2.click(':nth-child(2) > .doc-search-dialog__radio-caption',{waitUntil: 'load'});
            await page2.focus(':nth-child(2) > .doc-search-dialog__radio-caption');
            await page2.evaluate( ()=>document.getElementById("InputID").value="")
            //await page2.keyboard.down('Delete',{delay: 1000}); надо удалить предыдущее значение
            //await page2.keyboard.up('Delete');
            await page2.type(':nth-child(2) > .doc-search-dialog__input', clipboardy.readSync());
//            await page2.waitForTimeout(4000);
            //await page2.keyboard.down('Control',{delay: 250});  
            //await page2.keyboard.down('V');
            await page2.waitForSelector('.js-dialog-button',{timeout: 60000}).then(()=>page2.click( '.js-dialog-button'));
            const data = await page2.evaluate(() => {
                const tds = Array.from(document.querySelectorAll(`.js-grid-cell `))
                return tds.map(td => td.innerText)
            });
            await page2.waitForTimeout(2000);
            console.log(data);
            //const page3=await browser.newPage();
            //await page3.setViewport({
            //    width: 1920,
            //    height:1080
            //});
            //await page3.goto(`http://10.248.35.9:8092/WebShell/%D0%97%D0%B0%D1%8F%D0%B2%D0%BA%D0%B0_%D0%BD%D0%B0_%D0%B3%D1%80%D1%83%D0%B7%D0%BE%D0%BF%D0%B5%D1%80%D0%B5%D0%B2%D0%BE%D0%B7%D0%BA%D1%83?docId=${data[0]}`, {
            //    waitUntil: 'networkidle', waitUntil: 'load', timeout: 0});
            //await page2.close();            
                //lock lock-active
                //js http://10.248.35.9:8092/WebShell/%D0%97%D0%B0%D1%8F%D0%B2%D0%BA%D0%B0_%D0%BD%D0%B0_%D0%B3%D1%80%D1%83%D0%B7%D0%BE%D0%BF%D0%B5%D1%80%D0%B5%D0%B2%D0%BE%D0%B7%D0%BA%D1%83?docId=1133079974
                    //await browser.close();
            
        }
    }

})();