$ErrorActionPreference='Stop'
$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');$wb=$xl.ActiveWorkbook
$s=$wb.Worksheets.Add();$s.Name='Report';$s.Cells.Font.Name='Aptos';$s.Cells.Font.Size=14
$s.Range('A1:F1').Merge();$s.Range('A1').Value2='October close';$s.Range('A1').Font.Size=28;$s.Range('A1').Font.Bold=$true
$s.Range('A2:F2').Merge();$s.Range('A2').Value2='Operating expense | USD | Positive variance is unfavorable';$s.Range('A2').Font.Size=12;$s.Range('A2').Font.Color=7368816
$s.Range('A3').Value2='Reporting month';$s.Range('B3').NumberFormat='@';$s.Range('B3').Value2='2026-10';$s.Range('B3').Font.Color=12611584
$s.Columns.Item(1).ColumnWidth=23;$s.Range('B:D').ColumnWidth=17;$s.Columns.Item(5).ColumnWidth=12;$s.Columns.Item(6).ColumnWidth=17
$headers=@('Department','Budget','Actual','$ variance','% variance','Status')
foreach($r in @(5,13,22)){for($i=0;$i -lt 6;$i++){$s.Cells.Item($r,($i+1)).Value2=[string]$headers[$i]};$s.Range('A'+$r+':F'+$r).Interior.Color=6171950;$s.Range('A'+$r+':F'+$r).Font.Color=16777215;$s.Range('A'+$r+':F'+$r).Font.Bold=$true;$s.Range('A'+$r+':F'+$r).Font.Size=12}
foreach($base in @(6,14)){for($i=0;$i -lt 3;$i++){$s.Cells.Item(($base+$i),1).Value2=[string]@('Sales','Engineering','G&A')[$i]};$s.Cells.Item(($base+3),1).Value2='Total';$s.Range('A'+($base+3)+':F'+($base+3)).Font.Bold=$true;$s.Range('A'+($base+3)+':F'+($base+3)).Interior.Color=15921906}
$s.Range('A11:F11').Merge();$s.Range('A11').Value2='Year to date | January–October';$s.Range('A11').Font.Size=20;$s.Range('A11').Font.Bold=$true
$s.Range('A20:F20').Merge();$s.Range('A20').Value2='October variance by category';$s.Range('A20').Font.Size=20;$s.Range('A20').Font.Bold=$true
$s.Range('A22').Value2='Category'
$cats=@('Payroll','Cloud','Travel','Software')
for($i=0;$i -lt 4;$i++){$r=$i+23;$s.Cells.Item($r,1).Value2=[string]$cats[$i];$s.Cells.Item($r,2).Formula='=SUMIFS(Plan[Budget],Plan[Category],$A'+$r+',Plan[Period],$B$3)';$s.Cells.Item($r,3).Formula='=SUMIFS(Actuals[Amount],Actuals[Category],$A'+$r+',Actuals[Period],$B$3)';$s.Cells.Item($r,4).Formula='=C'+$r+'-B'+$r;$s.Cells.Item($r,5).Formula='=IF(B'+$r+'=0,NA(),D'+$r+'/ABS(B'+$r+'))';$s.Cells.Item($r,6).Formula='=IF(D'+$r+'>0,"Unfavorable",IF(D'+$r+'<0,"Favorable","On plan"))'}
$s.Range('B6:D26').NumberFormat='#,##0;(#,##0);"–"';$s.Range('E6:E26').NumberFormat='0.0%;(0.0%);"–"'
$s.Rows.RowHeight=29;$s.Rows.Item(1).RowHeight=44;$s.Rows.Item(2).RowHeight=24;$s.Rows.Item(3).RowHeight=32;$s.Range('A5:F26').VerticalAlignment=-4108
foreach($addr in @('D6:D9','D14:D17','D23:D26')){$fc=$s.Range($addr).FormatConditions.Add(1,5,'0');$fc.Font.Color=192;$fc=$s.Range($addr).FormatConditions.Add(1,6,'0');$fc.Font.Color=32768}
$s.Activate();$xl.ActiveWindow.DisplayGridlines=$false;$xl.ActiveWindow.Zoom=110;$s.Range('B6').Select()|Out-Null
$wb.Save()
Write-Output 'Report layout ready for formulas'
