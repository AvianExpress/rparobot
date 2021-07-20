const puppeteer = require('puppeteer');
const fs=require('fs');
const xlsx=require('xlsx');
const clipboardy=require('clipboardy');
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

async function sleep(time){
  const prm = new Promise((res,rej)=>setTimeout(()=>{ res()},time));
  await prm;
}

(async () =>{

  //Объявляем две функции-одна будет красить в красный уже отсутствующие, другая добавлять и красить в зелёный
  function paintRed(column, row,){
    var spawn = require("child_process").spawn,child;
     let str = ["AHAHAH", column, row, 1234];
     child = spawn("powershell.exe",["C:\\Users\\sherbova_as\\robot\\excel2.ps1", str]);
     //child.stdin.write();
     child.stdout.on("data",function(data){
         console.log("Powershell Data: " + data);
     });
     child.stderr.on("data",function(data){
         console.log("Powershell Errors: " + data);
     });
     child.on("exit",function(){
         console.log();
         console.log("Powershell  Paint Script finished");
     });
     child.stdin.end();
    }
  function addVagon(column, row, vagonNumber){
        var spawn = require("child_process").spawn,child;
         let str = ["AHAHAH", column, row, vagonNumber];
         child = spawn("powershell.exe",["C:\\Users\\sherbova_as\\robot\\excel.ps1", str]);
         //child.stdin.write();
         child.stdout.on("data",function(data){
             console.log("Powershell Data: " + data);
         });
         child.stderr.on("data",function(data){
             console.log("Powershell Errors: " + data);
         });
         child.on("exit",function(){
             console.log();
             console.log("Powershell Add Script finished");
         });
         child.stdin.end();
        }

let vb=xlsx.readFile("test.xlsx",{cellDates:true});
//TODO: сделать выбор файла
//У них разные имена книг!!!!!!!
let listCounter=0;
// let vs = vb.Sheets[vb.SheetNames[listCounter]];
while (vb.SheetNames[listCounter] != undefined){
  console.log(vb.SheetNames[listCounter]);
  let str = vb.SheetNames[listCounter].toLowerCase();
  if ( str.indexOf('втормет')!=-1){
      console.log('Таки втормет ' + listCounter );

let count=0;    
let vs = vb.Sheets[vb.SheetNames[listCounter]];
let range = xlsx.utils.decode_range(vs['!ref']);
let i =0;
let address = {c:0, r:0};
let addressNСarriage = {c:0, r:0};
let addressGU = {c:0, r:0};
let cell_ref = xlsx.utils.encode_cell(address);
for (i=0; i < range.e.c+2;  i++){
    address = {c:i, r:0}; //r=0!!!
    cell_ref = xlsx.utils.encode_cell(address);
    console.log(cell_ref);
    let Sheet1C = vs[cell_ref];
    if (Sheet1C!=undefined){
    if (Sheet1C.v === "№ ГУ-12"){
        addressGU = address;
    }
    if (Sheet1C.v === "Номера вагонов"){
        addressNСarriage = address;
    }
  }
}


let cell_ref_GU={c: addressGU, r:0};
let cell_ref_Ncarr={c:addressNСarriage, r:0};
let cell_ref_GUD;
let cell_ref_NcarrD;
// let value;
let Sheet1Number;
let Sheet1Vagon;
let mockup=[];
let eq=false;
let equ=[];


//Начинаем перебор
//Установить i<range.e.r когда закончятся тесты!!!!
for (i=0; i<range.e.r; i++){
  cell_ref_GU={c: addressGU.c, r:i};
  cell_ref_GUD = xlsx.utils.encode_cell(cell_ref_GU)
  cell_ref_Ncarr={c: addressNСarriage.c, r:i};
  cell_ref_NcarrD = xlsx.utils.encode_cell(cell_ref_Ncarr)
  Sheet1Number= vs[cell_ref_GUD];
  Sheet1Vagon=vs[cell_ref_NcarrD];
    if (Sheet1Number && (String(Sheet1Number.v).length == 8))
    {
        clipboardy.writeSync('');
        clipboardy.writeSync(String(Sheet1Number.v));
        console.log(clipboardy.readSync())
        //Получение данных с ЭТРАН-а
        //TODO: протестить и пофиксить последний этап, там что-то нехорошее
        const browser = await puppeteer.launch({
            headless:false,
            slowmo:200,
            ignoreHTTPSErrors:true,
            args: ['--start-maximized'],
            executablePath: 'C:/Program Files/Google/Chrome/Application/chrome',
            devtools: true
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
    
    
    
    // await sleep(3);
    console.log(`await retryHard(() => page.goto('http://10.248.35.9:8092/WebShell/',{timeout:10000}), 10000);`)
    await retryHard(() => page.goto('http://10.248.35.9:8092/WebShell/',{timeout:10000}), 10000);
    console.log(`await page.waitForTimeout(2000);`)
    await page.waitForTimeout(2000);
    console.log(`await retryHard(()=> page.waitForSelector('.dialog-auth-user-name',{timeout:5000}), 5000); `)
    await retryHard(()=> page.waitForSelector('.dialog-auth-user-name',{timeout:5000}), 5000); 
    console.log(`await page.type('.dialog-auth-user-name', 'Осина_ЮВ');`)
    await page.type('.dialog-auth-user-name', 'Осина_ЮВ');
    console.log(`await page.type('.dialog-auth-user-password', 'Etran222');`)
    await page.type('.dialog-auth-user-password', 'Etran222');
    console.log(` await page.click('.button'); `)
    await page.click('.button'); 
    console.log(`await page.waitForSelector(':nth-child(16) > .shell-menu-item-img',{timeout:0});`)
    await page.waitForSelector(':nth-child(16) > .shell-menu-item-img',{timeout:0});
    console.log(`await retryHard(() => page.goto('http://10.248.35.9:8092/WebShell/%D0%97%D0%B0%D1%8F%D0%B2%D0%BA%D0%B0_%D0%BD%D0%B0_%D0%B3%D1%80%D1%83%D0%B7%D0%BE%D0%BF%D0%B5%D1%80%D0%B5%D0%B2%D0%BE%D0%B7%D0%BA%D1%83',{timeout:5000}), 5000);`)
    await retryHard(() => page.goto('http://10.248.35.9:8092/WebShell/%D0%97%D0%B0%D1%8F%D0%B2%D0%BA%D0%B0_%D0%BD%D0%B0_%D0%B3%D1%80%D1%83%D0%B7%D0%BE%D0%BF%D0%B5%D1%80%D0%B5%D0%B2%D0%BE%D0%B7%D0%BA%D1%83',{timeout:5000}), 5000);
    console.log(`await page.waitForTimeout(4000);`)
    await page.waitForTimeout(4000);
    console.log(`await retryHard(() => page.waitForSelector('.js-button-find-document',{timeout:10000}), 10000);`)
    await retryHard(() => page.waitForSelector('.js-button-find-document',{timeout:10000}), 10000);
    console.log(`await page.click('.js-button-find-document');`)
    await page.click('.js-button-find-document');
    console.log(` await page.click(':nth-child(2) > .doc-search-dialog__radio-caption');`)
    await page.click(':nth-child(2) > .doc-search-dialog__radio-caption');
    console.log(`await page.keyboard.down('Control',{delay: 250});`)
    await page.keyboard.down('Control',{delay: 250});
    console.log(`await page.keyboard.down('V');`)
    await page.keyboard.down('V');
    console.log(`await page.waitForSelector('.dialog__body > .dialog-data > .doc-search-dialog__container > .doc-search-dialog__buttons > .button:nth-child(1)');`)
    await page.waitForSelector('.dialog__body > .dialog-data > .doc-search-dialog__container > .doc-search-dialog__buttons > .button:nth-child(1)');
    console.log(`await page.click('.dialog__body > .dialog-data > .doc-search-dialog__container > .doc-search-dialog__buttons > .button:nth-child(1)');`)
    await page.click('.dialog__body > .dialog-data > .doc-search-dialog__container > .doc-search-dialog__buttons > .button:nth-child(1)');
   
    let num
    try{
    await page.waitForSelector('.ui-grid-summary-page > .ui-grid-row > .first-cell > .cell-content',{timeout:5000});
    num = await page.$eval('.ui-grid-summary-page > .ui-grid-row > .first-cell > .cell-content', ele => ele.textContent);
    }
    catch (e)
    {
      console.log('Отсутствует ГУ');
      num=undefined;
    }
    
    console.log(num);
    if (num!=undefined){
    num1=Number(num);
    num1 +=2;
    // num1++;
    // num1++;
    console.log(num1);
    try{
    await page.waitForSelector(':nth-child('+ num1 +') > .first-cell > .cell-content',{timeout:500});
    await sleep(300);
    await page.click(':nth-child('+ num1 +') > .first-cell > .cell-content', {clickCount: 2 });
    }
    catch (e){
      console.log('С этой ГУ что-то не так, у неё нету версий. Согласована с первого раза?');
    }
    finally{
   console.log(`await page.waitForSelector('.xm-object > .nav > .xm-tabsheet-caption:nth-child(10) > .nav-link > .xm-tabsheet-caption-text',{timeout:0});`)
      await page.waitForSelector('.xm-object > .nav > .xm-tabsheet-caption:nth-child(10) > .nav-link > .xm-tabsheet-caption-text',{timeout:0});
   console.log(`await page.click('.xm-object > .nav > .xm-tabsheet-caption:nth-child(10) > .nav-link > .xm-tabsheet-caption-text', {timeout: 0});`)
    await page.click('.xm-object > .nav > .xm-tabsheet-caption:nth-child(10) > .nav-link > .xm-tabsheet-caption-text', {timeout: 0});
   console.log(`await page.waitForTimeout(1000);`);

   process.exit(1);

    await page.waitForTimeout(1000);
   console.log(`await page.waitForSelector('.xm-container-data > :nth-child(5) > .btn-caption');`)
    await page.waitForSelector('.xm-container-data > :nth-child(5) > .btn-caption');
   console.log(`await page.click('.xm-container-data > :nth-child(5) > .btn-caption');`)
    await page.click('.xm-container-data > :nth-child(5) > .btn-caption');
   console.log(`await page.waitForSelector('.xm-object:nth-child(1) > .xm-container-data:nth-child(2) > .xm-grid-container-wrapper:nth-child(1) > .xm-grid-container:nth-child(2) > .ui-grid:nth-child(2) > .ui-grid-content:nth-child(6) > .ui-grid-body:nth-child(3) .ui-grid-row:nth-child(3) > .ui-grid-cell:nth-child(4) > .cell-content:nth-child(1)',{timeout:0});`)
    await page.waitForSelector('.xm-object:nth-child(1) > .xm-container-data:nth-child(2) > .xm-grid-container-wrapper:nth-child(1) > .xm-grid-container:nth-child(2) > .ui-grid:nth-child(2) > .ui-grid-content:nth-child(6) > .ui-grid-body:nth-child(3) .ui-grid-row:nth-child(3) > .ui-grid-cell:nth-child(4) > .cell-content:nth-child(1)',{timeout:0});
   console.log(`await page.click('.xm-object:nth-child(1) > .xm-container-data:nth-child(2) > .xm-grid-container-wrapper:nth-child(1) > .xm-grid-container:nth-child(2) > .ui-grid:nth-child(2) > .ui-grid-content:nth-child(6) > .ui-grid-body:nth-child(3) .ui-grid-row:nth-child(3) > .ui-grid-cell:nth-child(4) > .cell-content:nth-child(1)');`)
    await page.click('.xm-object:nth-child(1) > .xm-container-data:nth-child(2) > .xm-grid-container-wrapper:nth-child(1) > .xm-grid-container:nth-child(2) > .ui-grid:nth-child(2) > .ui-grid-content:nth-child(6) > .ui-grid-body:nth-child(3) .ui-grid-row:nth-child(3) > .ui-grid-cell:nth-child(4) > .cell-content:nth-child(1)');
   console.log(`await page.keyboard.press('ArrowDown');`)
    await page.keyboard.press('ArrowDown');
   console.log(`await page.keyboard.press('ArrowUp');`)
    await page.keyboard.press('ArrowUp');
   console.log(`await page.keyboard.down('Control');`)
    await page.keyboard.down('Control');
   console.log(`await page.keyboard.press('C');`)
    await page.keyboard.press('C');
   console.log(`await page.keyboard.up('Control');`)
    await page.keyboard.up('Control');
    for (z=0; mockup[z-1]!=clipboardy.readSync(); z++){
    mockup.push(clipboardy.readSync());
    console.log(`await page.keyboard.press('ArrowDown');`);
    await page.keyboard.press('ArrowDown');
    console.log(`await page.keyboard.down('Control');`);
    await page.keyboard.down('Control');
    console.log(`await page.keyboard.press('C');`);
    await page.keyboard.press('C');
    console.log(`await page.keyboard.up('Control');`);
    await page.keyboard.up('Control');
    }
  }
  }
    await page.waitForTimeout(1000);
    console.log(`await page.waitForTimeout(1000);`);
    browser.close();
    eq=false;
    equ=[];
    cell_ref_Ncarr={c: addressNСarriage.c, r: i};
    cell_ref_NcarrD = xlsx.utils.encode_cell(cell_ref_Ncarr)
    Sheet1Vagon=vs[cell_ref_NcarrD];






//___________________________________________________________________
        while(!Sheet1Vagon)
        {
          i++
        cell_ref_Ncarr={c: addressNСarriage.c, r: i};
        cell_ref_NcarrD = xlsx.utils.encode_cell(cell_ref_Ncarr)
        Sheet1Vagon=vs[cell_ref_NcarrD];
        }       
        console.log('Первый вагон ' + Sheet1Vagon.v);
        console.log('На строке '+ i);
        console.log('Столбец '+ cell_ref_Ncarr.c);
        while (Sheet1Vagon){
          console.log('Под эту вот ГУ вагон с номером '+ Sheet1Vagon.v);
          for (j=0; j<mockup.length; j++){
            if (mockup[j]==Sheet1Vagon.v){
              eq=true;
              equ.push(j);
            }
          }
          if (eq==false){ 
             console.log(cell_ref_Ncarr.c+1+count);
             console.log(cell_ref_Ncarr.r+1+count);
             paintRed(cell_ref_Ncarr.c+1, i+1+count);
             await delay(5000);
             console.log('Данный вагон ' +Sheet1Vagon.v+' более не присутствует в ГУ. Нужно перекрасить его в красный цвет');
          }
          eq=false;
          i++
          cell_ref_Ncarr={c: addressNСarriage.c, r:i};
          cell_ref_NcarrD = xlsx.utils.encode_cell(cell_ref_Ncarr);
          Sheet1Vagon=vs[cell_ref_NcarrD];

        }
        await delay(5000);
        let add=true;
        console.log(equ);
        for (j=0; j<mockup.length; j++)
        {
          for (q=0; q<equ.length; q++){
            if (j == equ[q]) {
              add=false;
            }
          }
          if (add===true){
          console.log('Индекс из мокапа, который надо добавить '+ j);
          console.log('Элемент, который стоит добавить: '+ mockup[j])
          console.log(addressNСarriage.c);
          console.log(i);
          addVagon(addressNСarriage.c+1, i+1+count, mockup[j]);
          count++;
          await delay(5000);
          // paintSmth(addressNСarriage.c+1, i);
        
          }
          add=true;
        }
        console.log(mockup);
        mockup=[];
        equ=[];
        i--;
        console.log(i);
      }
     
    }
    
  }
  listCounter++;
  
  // let wb = xlsx.utils.book_new();
  // let vs2=xlsx.utils.json_to_sheet(vsjson);
  // xlsx.utils.book_append_sheet(wb, vs2);
  // listCounter++;
  // let vs = vb.Sheets[vb.SheetNames[listCounter]];
}
// xlsx.writeFile(wb, 'result.xlsx');    
})();
