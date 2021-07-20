const puppeteer = require('puppeteer');
const fs=require('fs');
const xlsx=require('xlsx');
const clipboardy=require('clipboardy');


(async () =>{

let vb=xlsx.readFile("test.xlsx",{cellDates:true});
//У них разные имена книг!!!!!!!
let listCounter=0;
// let vs = vb.Sheets[vb.SheetNames[listCounter]];
while (vb.SheetNames[listCounter] != undefined){
  console.log(vb.SheetNames[listCounter]);
  let str = vb.SheetNames[listCounter].toLowerCase();
  if ( str.indexOf('втормет')!=-1){
      console.log('Таки втормет ' + listCounter );


let vs = vb.Sheets[vb.SheetNames[listCounter]];
let vsjson=xlsx.utils.sheet_to_json(vs);
// console.log(vsjson);
let range = xlsx.utils.decode_range(vs['!ref']);
let i =0;
let address = {c:0, r:0};
let addressNСarriage = {c:0, r:0};
let addressGU = {c:0, r:0};
let cell_ref = xlsx.utils.encode_cell(address);

//Ищем колонки с вагонами и номерами ГУ-шек
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


let Sheet1Number;
let Sheet1Vagon;
// const mockup=[61783031,54143979,66821562,56593478];
let mockup=[];
let eq=false;
let equ=[];



//Начинаем перебор по json
//Установить i<vsjson.length когда закончятся тесты!!!!
for (i=0; i<20; i++){
    Sheet1Number= vsjson[i]['№ ГУ-12'];
    Sheet1Vagon=vsjson[i]['Номера вагонов'];
    //Если мы нашли ГУ, то:
      if (Sheet1Number && (String(Sheet1Number).length === 8))
      {
        clipboardy.writeSync(String(Sheet1Number));
        console.log(clipboardy.readSync())
        //Получение данных с ЭТРАН-а
        //TODO: протестить и пофиксить последний этап, там что-то нехорошее
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
        await page.setDefaultNavigationTimeout(0);
        //__________________________________
    
    
    
    
    await retryHard(() => page.goto('http://10.248.35.9:8092/WebShell/',{timeout:10000}), 10000);
    await page.waitForTimeout(2000);
    await retryHard(()=> page.waitForSelector('.dialog-auth-user-name',{timeout:5000}), 5000); 
    await page.type('.dialog-auth-user-name', 'Осина_ЮВ');
    await page.type('.dialog-auth-user-password', 'Etran222');
    await page.click('.button'); 
    await page.waitForSelector(':nth-child(16) > .shell-menu-item-img',{timeout:0});
    await retryHard(() => page.goto('http://10.248.35.9:8092/WebShell/%D0%97%D0%B0%D1%8F%D0%B2%D0%BA%D0%B0_%D0%BD%D0%B0_%D0%B3%D1%80%D1%83%D0%B7%D0%BE%D0%BF%D0%B5%D1%80%D0%B5%D0%B2%D0%BE%D0%B7%D0%BA%D1%83',{timeout:5000}), 5000);
    await page.waitForTimeout(4000);
    await retryHard(() => page.waitForSelector('.js-button-find-document',{timeout:10000}), 10000);
    await page.click('.js-button-find-document');
    await page.click(':nth-child(2) > .doc-search-dialog__radio-caption');
    await page.keyboard.down('Control',{delay: 250});
    await page.keyboard.down('V');
    await page.waitForSelector('.dialog__body > .dialog-data > .doc-search-dialog__container > .doc-search-dialog__buttons > .button:nth-child(1)');
    await page.click('.dialog__body > .dialog-data > .doc-search-dialog__container > .doc-search-dialog__buttons > .button:nth-child(1)');
    await page.waitForSelector('.ui-grid-summary-page > .ui-grid-row > .first-cell > .cell-content',{timeout:0});
    const num = await page.$eval('.ui-grid-summary-page > .ui-grid-row > .first-cell > .cell-content', ele => ele.textContent);
    console.log(num);
    num1=num;
    num1++;
    num1++;
    console.log(num1);
    await page.waitForSelector(':nth-child('+ num1 +') > .first-cell > .cell-content',{timeout:0});
    await page.click(':nth-child('+ num1 +') > .first-cell > .cell-content', {clickCount: 2 });
    await page.waitForSelector('.xm-object > .nav > .xm-tabsheet-caption:nth-child(10) > .nav-link > .xm-tabsheet-caption-text',{timeout:0});
    await page.click('.xm-object > .nav > .xm-tabsheet-caption:nth-child(10) > .nav-link > .xm-tabsheet-caption-text', {timeout: 0});
    await page.waitForTimeout(4000);
    await page.waitForSelector('.xm-container-data > :nth-child(5) > .btn-caption');
    await page.click('.xm-container-data > :nth-child(5) > .btn-caption');
    await page.waitForSelector('.xm-object:nth-child(1) > .xm-container-data:nth-child(2) > .xm-grid-container-wrapper:nth-child(1) > .xm-grid-container:nth-child(2) > .ui-grid:nth-child(2) > .ui-grid-content:nth-child(6) > .ui-grid-body:nth-child(3) .ui-grid-row:nth-child(3) > .ui-grid-cell:nth-child(4) > .cell-content:nth-child(1)',{timeout:0});
    await page.click('.xm-object:nth-child(1) > .xm-container-data:nth-child(2) > .xm-grid-container-wrapper:nth-child(1) > .xm-grid-container:nth-child(2) > .ui-grid:nth-child(2) > .ui-grid-content:nth-child(6) > .ui-grid-body:nth-child(3) .ui-grid-row:nth-child(3) > .ui-grid-cell:nth-child(4) > .cell-content:nth-child(1)');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.down('Control');
    await page.keyboard.press('C');
    await page.keyboard.up('Control');
    for (z=0; mockup[z-1]!=clipboardy.readSync(); z++){
    mockup.push(clipboardy.readSync());
    await page.keyboard.press('ArrowDown');
    await page.keyboard.down('Control');
    await page.keyboard.press('C');
    await page.keyboard.up('Control');
    }
    await page.waitForTimeout(10000);
    browser.close();






//___________________________________________________________________
        while(Sheet1Vagon===undefined)
        {
        i++;
        Sheet1Vagon=vsjson[i]['Номера вагонов'];
        }       
        while (Sheet1Vagon){
          console.log('Под эту вот ГУ вагон с номером '+ Sheet1Vagon);
          for (j=0; j<mockup.length; j++){
            if (mockup[j]==Sheet1Vagon){
              eq=true;
              equ.push(j);
            }
          }
          if (eq===false){
            vsjson[i]['Номера вагонов']='!' + Sheet1Vagon + '!';
            console.log('Данный вагон ' +Sheet1Vagon+' более не присутствует в ГУ. Нужно перекрасить его в красный цвет');
          }
          eq=false;
          i++
          Sheet1Vagon=vsjson[i]['Номера вагонов'];

        }
        let add=true;
        console.log(equ);
        for (j=0; j<mockup.length; j++)
        {
          for (q=0; q<equ.length; q++){
            if (j === equ[q]) {
              add=false;
            }
          }
          if (add===true){
          console.log('Индекс из мокапа, который надо добавить '+ j);
          console.log('Элемент, который стоит добавить: '+ mockup[j])
          vsjson.splice(i-1, 0, {"Номера вагонов":  'NEW: '+ mockup[j]});
          }
          add=true;
        }
        equ=[];
        mockup=[]
      }
    }
  }
  let wb = xlsx.utils.book_new();
  let vs2=xlsx.utils.json_to_sheet(vsjson);
  xlsx.utils.book_append_sheet(wb, vs2);
  listCounter++;
  let vs = vb.Sheets[vb.SheetNames[listCounter]];
}
xlsx.writeFile(wb, 'result.xlsx');  
})();

