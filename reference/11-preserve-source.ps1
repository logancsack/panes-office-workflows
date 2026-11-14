$ErrorActionPreference='Stop'
$d='C:\work\Panes-office-demo';$xl=New-Object -ComObject Excel.Application;$xl.Visible=$true;$xl.DisplayAlerts=$false;$xl.WindowState=-4137
$wb=$xl.Workbooks.Open($d+'\October close.xlsx')
$rows=Get-Content -Raw -Encoding UTF8 ($d+'\raw-exact.json')|ConvertFrom-Json
$data=New-Object 'object[,]' $rows.Count,6
for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt 6;$c++){$data[$r,$c]=[string]$rows[$r][$c]}}
$wb.Worksheets.Item('Raw export').Range('A2:F484').Value2=$data
$wb.RefreshAll();$xl.CalculateUntilAsyncQueriesDone();$wb.Worksheets.Item('Pivot detail').PivotTables('ExpensePivot').RefreshTable()|Out-Null;$xl.CalculateFullRebuild()
$wb.Worksheets.Item('Report').Activate();$wb.Worksheets.Item('Report').Range('B6').Select();$wb.Save()
Write-Output ('Exact raw fields retained; actual='+$wb.Worksheets.Item('Report').Range('C9').Value2+'; pivot tie='+$wb.Worksheets.Item('Checks').Range('B10').Value2)
