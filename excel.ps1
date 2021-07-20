$ExcelObj = New-Object -comobject Excel.Application;
$ExcelPath=$args[0][0];
$ExcelWorkbook = $ExcelObj.Workbooks.Open($ExcelPath);
$ExcelObj.visible=$true;
$ExcelWorkSheet = $ExcelWorkbook.Sheets.Item(1);
$b = $args[0][1];
$c = $args[0][2];
$xlShiftDown = -4121;
$eRow = $ExcelWorkSheet.cells.item($c,$b).entireRow;
[void]$eRow.Insert($xlShiftDown)
$ExcelWorkSheet.Columns.Item($b).Rows.Item($c)=$args[0][3];
$ExcelWorkSheet.Columns.Item($b).Rows.Item($c).Font.ColorIndex = 4; 
$ExcelWorkbook.Close($true)
$ExcelObj.Quit();
# Write-Host $args[0][3]
# [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')