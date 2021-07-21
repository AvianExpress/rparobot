const puppeteer = require('puppeteer');
const fs=require('fs');
const xlsx=require('xlsx');
const clipboardy=require('clipboardy');
const args = process.argv.slice(2);
const pathToEx=args[0];
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
// let  access = fs.createWriteStream(dir + '/node.access.log', { flags: 'a' })
//       , error = fs.createWriteStream(dir + '/node.error.log', { flags: 'a' });
const util = require('util');
let start1= new Date();
let logname = ('/debug'+ start1.getHours()+start1.getMinutes()+start1.getSeconds()+ '.log');
let log_file = fs.createWriteStream(__dirname + logname, {flags : 'w'});
let log_stdout = process.stdout;
let log_filed = fs.createWriteStream(__dirname + '/debug'+ Date.now()+'.log', {flags : 'w'});
console.log = function(d, param1=undefined) {
  if (param1!=undefined){
    start1=new Date();
  param1.write('['+start1.getHours() +':'+start1.getMinutes()+':'+start1.getSeconds()+'] '+util.format(d) + '\n');
  }
  log_stdout.write('['+start1.getHours() +':'+start1.getMinutes()+':'+start1.getSeconds()+'] '+util.format(d) + '\n');
 
};

(async () =>{

  //Объявляем две функции-одна будет красить в красный уже отсутствующие, другая добавлять и красить в зелёный
  function paintRed(path, column, row,){
    var spawn = require("child_process").spawn,child;
     let str = [path, column, row, 1234];
     child = spawn("powershell.exe",["C:\\Users\\sherbova_as\\robot\\excel2.ps1", str]);
     //child.stdin.write();
     child.stdout.on("data",function(data){
         console.log("Powershell Data: " + data, log_file);
     });
     child.stderr.on("data",function(data){
         console.log("Powershell Errors: " + data, log_file);
     });
     child.on("exit",function(){
         console.log();
         console.log("Powershell  Paint Script finished", log_file);
     });
     child.stdin.end();
    }
  function addVagon(path,column, row, vagonNumber){
        var spawn = require("child_process").spawn,child;
         let str = [path, column, row, vagonNumber];
         child = spawn("powershell.exe",["C:\\Users\\sherbova_as\\robot\\excel.ps1", str]);
         //child.stdin.write();
         child.stdout.on("data",function(data){
             console.log("Powershell Data: " + data, log_file);
         });
         child.stderr.on("data",function(data){
             console.log("Powershell Errors: " + data, log_file);
         });
         child.on("exit",function(){
             console.log();
             console.log("Powershell Add Script finished", log_file);
         });
         child.stdin.end();
        }

let vb=xlsx.readFile("test.xlsx",{cellDates:true});
//У них разные имена книг!!!!!!!
let listCounter=0;
// let vs = vb.Sheets[vb.SheetNames[listCounter]];
while (vb.SheetNames[listCounter] != undefined){
  console.log(vb.SheetNames[listCounter]);
  let str = vb.SheetNames[listCounter].toLowerCase();
  if ( str.indexOf('втормет')!=-1){
      console.log(' Лист со вторметом найден: '  + listCounter , log_file);

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
        console.log('Колонка с номерами ГУ найдена: '+ cell_ref.c, log_file);
    }
    if (Sheet1C.v === "Номера вагонов"){
        addressNСarriage = address;
        console.log('Колонка с номерами вагонов найдена: '+ cell_ref.c, log_file);
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
        console.log('Номер ГУ найден: '+clipboardy.readSync(), log_file)
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
        //Функция для того, чтобы пробовать ещё раз, без перезагрузки
        const retry = (fn, ms)=> new Promise (resolve => {
            fn()
                .then(resolve)
                .catch(() => {
                    setTimeout(()=>{
                        console.log('retrying...', log_filed);
                        retry(fn, ms).then(resolve);
                    }, ms);
                })
            });
        //Тоже ретрай, но с перезагрузкой
        const retryHard = (fn, ms)=> new Promise (resolve => {
        fn()
            .then(resolve)
            .catch(() => {
                setTimeout(()=>{
                    page.reload();
                    console.log('retrying hard...', log_filed);
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
        //Блок снятия данных с ЭТРАН-а
    
    
    //Идём на главную
    await retryHard(() => page.goto('http://10.248.35.9:8092/WebShell/',{timeout:10000}), 10000);
    await page.waitForTimeout(2000);
    //Ждём пока подгрузится и логинимся
    await retryHard(()=> page.waitForSelector('.dialog-auth-user-name',{timeout:5000}), 5000); 
    await page.type('.dialog-auth-user-name', 'Осина_ЮВ');
    await page.type('.dialog-auth-user-password', 'Etran222');
    await page.click('.button'); 
    //Идём в ГУ-шки
    await retryHard(() => page.waitForSelector(':nth-child(16) > .shell-menu-item-img',{timeout:5000}), 5000);

    await retryHard(() => page.goto('http://10.248.35.9:8092/WebShell/%D0%97%D0%B0%D1%8F%D0%B2%D0%BA%D0%B0_%D0%BD%D0%B0_%D0%B3%D1%80%D1%83%D0%B7%D0%BE%D0%BF%D0%B5%D1%80%D0%B5%D0%B2%D0%BE%D0%B7%D0%BA%D1%83',{timeout:5000}), 5000);
    await page.waitForTimeout(4000);
    //Ждём пока подгрузится форма поиска
    await retryHard(() => page.waitForSelector('.js-button-find-document',{timeout:10000}), 10000);
    await page.click('.js-button-find-document');
    await page.click(':nth-child(2) > .doc-search-dialog__radio-caption');
    //Из буфера обмена вкидываем номер ГУ
    //!!!ВАЖНО!!! в буфере не должно быть левых данных, иначе магия не удастся
    await page.keyboard.down('Control');
    await page.keyboard.press('V');
    await page.keyboard.up('Control');
    //И ждём, пока не появится кнопка поиска
    await page.waitForSelector('.dialog__body > .dialog-data > .doc-search-dialog__container > .doc-search-dialog__buttons > .button:nth-child(1)');
    await page.click('.dialog__body > .dialog-data > .doc-search-dialog__container > .doc-search-dialog__buttons > .button:nth-child(1)');
    let num
    //Здесь мы кликаем получаем количество итераций и кликаем по последней
    try{
    await page.waitForSelector('.ui-grid-summary-page > .ui-grid-row > .first-cell',{timeout:5000});
    num = await page.$eval('.ui-grid-summary-page > .ui-grid-row > .first-cell > .cell-content', ele => ele.textContent);
    }
    catch (e)
    {
      console.log('ГУ с таким номером отсутствует в базе ЭТРАН', log_file);
      num=undefined;
      i++;
    }
    
    console.log(num, log_filed);
    if (num!=undefined){
    num1=num;
    num1=num1+2;
    console.log(num1, log_filed);
    try{
      //Если он будет выделываться и не кликать в таблицу-раскомментить две нижние строчки, а те, что под ними, закомментить
    // await page.WaitForSelector(':nth-child('+ num1 +') > .first-cell');
    // await page.click(':nth-child('+ num1 +') > .first-cell', {clickCount: 2 });
    await page.waitForSelector(':nth-child('+ num1 +') > .first-cell',{timeout:0});
    await page.click(':nth-child('+ num1 +') > .first-cell > .cell-content', {clickCount: 2 });
    }
    catch (e){
      console.log('С этой ГУ что-то не так, у неё нету версий. Согласована с первого раза?', log_file);
    }
    finally{
      //Здесь уже работа с самим документом, идём в накладные и проходимся по вагонам
    await page.waitForSelector('.xm-object > .nav > .xm-tabsheet-caption:nth-child(10) > .nav-link > .xm-tabsheet-caption-text',{timeout:0});
    await page.click('.xm-object > .nav > .xm-tabsheet-caption:nth-child(10) > .nav-link > .xm-tabsheet-caption-text', {timeout: 0});
    await page.waitForTimeout(1000);
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
  }
  
    await page.waitForTimeout(1000);
    browser.close();
    eq=false;
    equ=[];
    cell_ref_Ncarr={c: addressNСarriage.c, r: i};
    cell_ref_NcarrD = xlsx.utils.encode_cell(cell_ref_Ncarr)
    Sheet1Vagon=vs[cell_ref_NcarrD];






//___________________________________________________________________
//Блок работы с экселем
//Пока у нас вагоны не начались, мы скипаем
        while(!Sheet1Vagon)
        {
          i++
        cell_ref_Ncarr={c: addressNСarriage.c, r: i};
        cell_ref_NcarrD = xlsx.utils.encode_cell(cell_ref_Ncarr)
        Sheet1Vagon=vs[cell_ref_NcarrD];
        }       
        // console.log('Первый вагон ' + Sheet1Vagon.v);
        // console.log('На строке '+ i);
        // console.log('Столбец '+ cell_ref_Ncarr.c);
        //Когда вагон найден, мы его обрабатываем
        while (Sheet1Vagon){
          console.log('Под эту ГУ в таблице записан вагон с номером '+ Sheet1Vagon.v, log_file);
          for (j=0; j<mockup.length; j++){
            //Мы сравниваем полученный из ЭТРАН-а массив с вагоном, и если его в массиве нету-его не должно быть и в ГУ
            //Одновременно, если такой вагон есть, его индекс мы запишем в другой массив, который обозначает те вагоны,
            //Которые есть и которые добавлять не надо
            if (mockup[j]==Sheet1Vagon.v){
              eq=true;
              equ.push(j);
            }
          }
          if (eq==false){ 
            //Блок закрашивания ненужных вагонов в красный
             console.log(cell_ref_Ncarr.c+1+count);
             console.log(cell_ref_Ncarr.r+1+count);
             paintRed(pathToEx,cell_ref_Ncarr.c+1, i+1+count);
             await delay(5000);
             console.log('Данный вагон ' +Sheet1Vagon.v+' более не присутствует в ГУ', log_file);
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
        //В этом цикле мы смотрим индексы вагонов из ЭТРАН-а, которые НЕ нужно вставлять, и вставляем все остальные
        for (j=0; j<mockup.length; j++)
        {
          for (q=0; q<equ.length; q++){
            if (j == equ[q]) {
              add=false;
            }
          }
          if (add===true)
          //Блок вставки недостающего вагона
          {
          console.log('Индекс из ЭТРАН-а, который надо добавить '+ j, log_filed);
          console.log('Элемент из базы ЭТРАН-А на добавление: '+ mockup[j], log_file)
          console.log(addressNСarriage.c);
          console.log(i, log_filed);
          addVagon(pathToEx,addressNСarriage.c+1, i+1+count, mockup[j]);
          count++;
          await delay(5000);
          // paintSmth(addressNСarriage.c+1, i);
        
          }
          add=true;
        }
      }
        await browser.close();
        console.log(mockup, log_filed);
        mockup=[];
        equ=[];
        i--;
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

