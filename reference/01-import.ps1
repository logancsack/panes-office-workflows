$ErrorActionPreference='Stop'
$d='C:\work\Panes-office-demo'
$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$xl.Visible=$true;$xl.DisplayAlerts=$false;$xl.WindowState=-4137
$wb=$xl.ActiveWorkbook
if ($wb.Path -ne '') {throw 'Expected our unsaved empty workbook'}
$raw=$wb.Worksheets.Item(1);$raw.Name='Raw export'
function LoadCsv($sheet,$path,$headers,$numericColumn=-1){
 $rows=@(Import-Csv $path);$data=New-Object 'object[,]' ($rows.Count+1),$headers.Count
 for($c=0;$c -lt $headers.Count;$c++){$data[0,$c]=[string]$headers[$c]}
 for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $headers.Count;$c++){if($c -eq $numericColumn){$data[($r+1),$c]=[double]$rows[$r].($headers[$c])}else{$data[($r+1),$c]=[string]$rows[$r].($headers[$c])}}}
 $range=$sheet.Range('A1').Resize(($rows.Count+1),$headers.Count)
 if($numericColumn -lt 0){$range.NumberFormat='@'}
 $range.Value2=$data
 return ,$range
}
$r=LoadCsv $raw ($d+'\GL-export.csv') @('DocID','PostingDate','CostCenter','Account','Amount','Vendor')
$exact=Get-Content -Raw -Encoding UTF8 ($d+'\raw-exact.json')|ConvertFrom-Json
$values=New-Object 'object[,]' $exact.Count,6
for($ri=0;$ri -lt $exact.Count;$ri++){for($ci=0;$ci -lt 6;$ci++){$values[$ri,$ci]=[string]$exact[$ri][$ci]}}
$raw.Range('A2:F484').Value2=$values
$t=$raw.ListObjects.Add(1,$r,[Type]::Missing,1);$t.Name='RawExport';$t.TableStyle='TableStyleMedium2';$raw.Columns.AutoFit()|Out-Null
$raw.Activate();$xl.ActiveWindow.Zoom=110
$raw.Range('A1').Select()|Out-Null
$budget=$wb.Worksheets.Add();$budget.Name='Budget'
$r=LoadCsv $budget ($d+'\Budget.csv') @('Period','Department','Category','Budget') 3
$t=$budget.ListObjects.Add(1,$r,[Type]::Missing,1);$t.Name='Plan';$t.TableStyle='TableStyleMedium2';$budget.Columns.AutoFit()|Out-Null;$budget.Columns.Item(4).NumberFormat='#,##0;(#,##0);"–"'
$readme=$wb.Worksheets.Add();$readme.Name='Read me'
$notes=@('October 2026 operating expense review','Synthetic demonstration data; USD; positive expenses.','Scope: all three departments; January–October 2026.','Source: GL-export.csv and Budget.csv; raw export is preserved.','Cleanup: trim IDs and codes, normalize case, parse dates/currency, remove exact duplicate DocIDs.','Expense variance = actual less budget. Positive = unfavorable; negative = favorable.','Monthly percentage = variance / budget. Total percentages use total budget, not averaged percentages.','YTD includes January through the selected month. No annualized estimates.','Power Query produces Actuals; the model uses SUMIFS; the native PivotTable provides a detail cross-check.','PowerPoint uses workbook figures and editable charts. Commentary describes arithmetic contributions, not unverified business causes.','Change source tables and refresh to update the workbook. The slides are a saved October reporting snapshot.')
for($i=0;$i -lt $notes.Count;$i++){$readme.Cells.Item(($i+1),1).Value2=[string]$notes[$i]}
$readme.Columns.Item(1).ColumnWidth=110;$readme.UsedRange.WrapText=$true;$readme.UsedRange.Rows.AutoFit()|Out-Null
$readme.Range('A1').Font.Size=22;$readme.Range('A1').Font.Bold=$true
$raw.Activate();$raw.Range('A1').Select()|Out-Null
$wb.SaveAs(($d+'\October close.xlsx'),51)
Write-Output ('Raw rows '+($raw.ListObjects.Item(1).ListRows.Count)+'; budget rows '+($budget.ListObjects.Item(1).ListRows.Count))
