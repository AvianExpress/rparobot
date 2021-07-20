$ExcelObj = New-Object -comobject Excel.Application;
$ExcelWorkbook = $ExcelObj.Workbooks.Open("C:\Users\sherbova_as\robot\test.xlsx");
$ExcelObj.visible=$true;
$ExcelWorkSheet = $ExcelWorkbook.Sheets.Item(1);
$b = $args[0][1];
$c = $args[0][2];
$ExcelWorkSheet.Columns.Item($b).Rows.Item($c).Font.ColorIndex = 3; 
# $eRow = $ExcelWorkSheet.cells.item($b,$c).entireRow;
# $active = $eRow.activate();
# $active = $eRow.insert();
# $ExcelWorkSheet.Columns.Item($b).Rows.Item($c)=$args[0][3];
$ExcelWorkbook.Close($true)
$ExcelObj.Quit();
