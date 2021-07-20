


const puppeteer = require('puppeteer');
const fs=require('fs');
const xlsx=require('xlsx');
const clipboardy=require('clipboardy');  
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
//_______
(async () =>{

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
       console.log("Powershell Script finished");
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
           console.log("Powershell Script finished");
       });
       child.stdin.end();
      }
//_______





let vb=xlsx.readFile("testtest.xlsx",{cellDates:true});
    //У них разные имена книг
    let vs = vb.Sheets[vb.SheetNames[0]];
    let range = xlsx.utils.decode_range(vs['!ref']);
    let i =0;
    let address = {c:0, r:0};
    let addressNСarriage = {c:0, r:0};
    let addressGU = {c:0, r:0};
    let cell_ref = xlsx.utils.encode_cell(address);
    for (i=0; i < range.e.c+2;  i++){
        console.log(range.e.c);
        address = {c:i, r:0}; //r=0!!!
        cell_ref = xlsx.utils.encode_cell(address);
        console.log(cell_ref);
        let Sheet1C = vs[cell_ref];
        if (Sheet1C!=undefined){
          console.log(Sheet1C.v);
        if (Sheet1C.v === "№ ГУ-12"){
            addressGU = address;
            console.log(addressGU.c + ' это ГУ-12');
        }
        if (Sheet1C.v === "Номера вагонов"){
            addressNСarriage = address;
            console.log(addressNСarriage.c + ' это адреснкарриаге');
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
    const mockup=[61783031,54143979,55821562,56593478];
    let eq=false;
    let equ=[];

    for (i=0; i<10; i++){
    cell_ref_GU={c: addressGU.c, r:i};
    cell_ref_GUD = xlsx.utils.encode_cell(cell_ref_GU)
    cell_ref_Ncarr={c: addressNСarriage.c, r:i};
    cell_ref_NcarrD = xlsx.utils.encode_cell(cell_ref_Ncarr)
    Sheet1Number= vs[cell_ref_GUD];
    Sheet1Vagon=vs[cell_ref_NcarrD];
      if (Sheet1Number && (String(Sheet1Number.v).length >= 8))
      {
        console.log('ГУ номер '+ Sheet1Number.v);
        console.log('Строка номер '+ cell_ref_GU.r);
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
          if (eq===false){
            
            
             paintRed(addressNСarriage.c+1, i+1);
             await delay(2000);
             console.log('Данный вагон ' +Sheet1Vagon.v+' более не присутствует в ГУ. Нужно перекрасить его в красный цвет');
         
          }
          eq=false;
          i++
          cell_ref_Ncarr={c: addressNСarriage.c, r:i};
          cell_ref_NcarrD = xlsx.utils.encode_cell(cell_ref_Ncarr);
          Sheet1Vagon=vs[cell_ref_NcarrD];

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
          console.log(addressNСarriage.c);
          console.log(i);
          addVagon(addressNСarriage.c+1, i+1, mockup[j]);
          await delay(2000);
          // paintSmth(addressNСarriage.c+1, i);
        
          }
          add=true;
        }
        equ=[];
      }
    }
   
  })();
    // (async () =>{
    // })()