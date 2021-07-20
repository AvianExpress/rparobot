//работает
function paintRed(column, row){
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

paintRed(3,1)


// const excel = "Threeeeeee";
// const { exec } = require('child_process');
// exec(`
// $ExcelObj = New-Object -comobject Excel.Application;
// $ExcelWorkbook = $ExcelObj.Workbooks.Open("C:\\Users\\sherbova_as\\robot\\test1.xlsx");
// $ExcelObj.visible=$true;
// $ExcelWorkSheet = $ExcelWorkbook.Sheets.Item(1);
// $ExcelWorkSheet.Columns.Item(5).Rows.Item(6)="${excel}";
// `, 
// {'shell':'powershell.exe'}, (error, stdout, stderr)=> { 
//    console.log(error)
// })

// var edge = require('edge');
// var hello = edge.func('ps', function () {/*
// "PowerShell welcomes $inputFromJS on $(Get-Date)"
// */});
// hello('Node.js', function (error, result) {
//     if (error) throw error;
//     console.log(result[0]);
// });