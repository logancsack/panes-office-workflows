$ErrorActionPreference='Stop'
$d='C:\work\Panes-monthly-close';$src=$d+'\Alder Mechanical - Finance\FY2026\10 Oct'
$w=[Runtime.InteropServices.Marshal]::BindToMoniker('C:\work\Panes-monthly-close\October outputs\October 2026 close.xlsx');$xl=$w.Application
$s=$w.Worksheets.Item('Mapping');$t=$s.ListObjects.Item('Map');$new=$t.ListRows.Add();$new.Range.Cells.Item(1,1).NumberFormat='@'
@('6265','Third-party warranty callouts','Warranty','PL','Oct 2026 new GL / subcontractor warranty callouts') | ForEach-Object -Begin {$j=1} -Process {$new.Range.Cells.Item(1,$j).Value2=[string]$_;$j++}
$s.Activate();$xl.Goto($new.Range,$true);$xl.ActiveWindow.Zoom=95
$xl.CalculateFullRebuild()
# Append October activity groups. Current-month amounts remain formulas against the mapped GL.
$raw=@(Import-Csv ($src+'\01 ERP exports\GL detail 2026-10.csv'))
$map=@{};for($i=2;$i -le $s.UsedRange.Rows.Count;$i++){$map[$s.Cells.Item($i,1).Text]=@($s.Cells.Item($i,3).Text,$s.Cells.Item($i,4).Text)}
$groups=@{}
foreach($r in $raw){$m=$map[$r.GLAccount];if($m -and $m[1] -eq 'PL' -and $r.Source -ne 'Year close'){$key=$r.Entity+'|'+$r.Department+'|'+$m[0];$groups[$key]=@($r.Entity,$r.Department,$m[0])}}
$h=$w.Worksheets.Item('Monthly History');$t=$h.ListObjects.Item('Facts')
foreach($key in ($groups.Keys|Sort-Object)){
  $v=$groups[$key];$row=$t.ListRows.Add();$i=$row.Range.Row
  $row.Range.Cells.Item(1,1).Value2=[double]([datetime]'2026-10-31').ToOADate()
  for($j=0;$j -lt 3;$j++){$row.Range.Cells.Item(1,$j+2).Value2=[string]$v[$j]}
  $criteria='GL[Entity],B'+$i+',GL[Department],C'+$i+',GL[Line],D'+$i+',GL[Source],"<>Year close",GL[Source],"<>Opening"'
  $h.Cells.Item($i,5).Formula='=IF(OR(D'+$i+'="Contract revenue",D'+$i+'="Demand revenue",D'+$i+'="IC recovery"),-1,1)*(SUMIFS(GL[Debit],'+$criteria+')-SUMIFS(GL[Credit],'+$criteria+'))'
}
# Roll schedules from September closing to October opening, using the imported ledger for movements.
$gl=$w.Worksheets.Item('GL Import').ListObjects.Item('GL')
function Amount($entity,$account,$side){$sum=0.0;foreach($r in $raw){if($r.Entity -eq $entity -and $r.GLAccount -eq $account){$sum+=[double]$r.$side}};return $sum}
$s=$w.Worksheets.Item('Prepaids')
for($i=2;$i -le 4;$i++){$e=$s.Cells.Item($i,1).Text;$s.Cells.Item($i,3).Value2=[double]$s.Cells.Item($i,6).Value2;$s.Cells.Item($i,4).Value2=[double](Amount $e '1300' 'Debit')}
$s=$w.Worksheets.Item('Fixed Assets')
for($i=2;$i -le 4;$i++){$e=$s.Cells.Item($i,1).Text;$priorCost=[double]$s.Cells.Item($i,5).Value2;$priorDep=[double]$s.Cells.Item($i,8).Value2;$s.Cells.Item($i,3).Value2=$priorCost;$s.Cells.Item($i,6).Value2=$priorDep;$s.Cells.Item($i,4).Value2=[double]((Amount $e '1400' 'Debit')-(Amount $e '1400' 'Credit'))}
$s=$w.Worksheets.Item('Accruals')
for($i=2;$i -le 7;$i++){$e=$s.Cells.Item($i,1).Text;$a=$s.Cells.Item($i,2).Text;$s.Cells.Item($i,4).Value2=[double]$s.Cells.Item($i,7).Value2;$s.Cells.Item($i,5).Value2=[double](Amount $e $a 'Credit');$s.Cells.Item($i,6).Value2=[double](Amount $e $a 'Debit')}
# Preserve trailing windows and add the October YTD column.
$s=$w.Worksheets.Item('Trailing 12')
for($j=2;$j -le 13;$j++){$dt=([datetime]'2025-11-01').AddMonths($j-2);$monthend=$dt.AddMonths(1).AddDays(-1);$s.Cells.Item(1,$j).Value2=[double]$monthend.ToOADate()}
$s=$w.Worksheets.Item('YTD');$s.Columns.Item(11).Insert() | Out-Null;$s.Cells.Item(1,11).Value2=[double]([datetime]'2026-10-31').ToOADate();$s.Cells.Item(1,11).NumberFormat='mmm-yy'
for($i=2;$i -le 15;$i++){$s.Cells.Item($i,11).Formula='=SUMIFS(Facts[Amount],Facts[Line],$A'+$i+',Facts[Period],K$1)';$s.Cells.Item($i,12).Formula='=SUM(B'+$i+':K'+$i+')'}
$w.Worksheets.Item('Checks').Range('B15').Formula='=SUMIFS(TB[Ending],TB[Type],"PL")+SUM(YTD!L2:L4)-SUM(YTD!L5:L15)'
$payroll=Import-Csv ($src+'\02 Subledgers\Payroll summary 2026-10.csv');$head=0;foreach($r in $payroll){$head+=[int]$r.'Ending FTE'};$w.Worksheets.Item('KPI').Range('B5').Value2=[double]$head
# Replace the journal log with October close/treasury entries.
$s=$w.Worksheets.Item('Journal Log');$s.Range('A2:E2000').ClearContents();$i=2
foreach($r in $raw){if($r.Source -in @('Close JE','Consolidation','Year close','Treasury')){$s.Cells.Item($i,1).Value2=[string]$r.JournalID;$s.Cells.Item($i,2).Value2=[string]$r.Memo;$s.Cells.Item($i,3).Value2=[double]$r.Debit;$s.Cells.Item($i,4).Value2=[double]$r.Credit;$s.Cells.Item($i,5).Value2='Controller';$i++}}
$xl.CalculateFullRebuild();$w.Worksheets.Item('PL Group').Activate();$xl.ActiveWindow.Zoom=85;$w.Save()
Write-Output 'Mapped new warranty account; history, schedules, T12, YTD and KPIs refreshed.'

Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class CloseFocus{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);}'
$w.Activate();[CloseFocus]::SetForegroundWindow([IntPtr]$xl.Hwnd) | Out-Null
