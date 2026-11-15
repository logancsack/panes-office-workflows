$ErrorActionPreference='Stop'
$d='C:\work\Panes-monthly-close';$src=$d+'\Alder Mechanical - Finance\FY2026\10 Oct\01 ERP exports'
$w=[Runtime.InteropServices.Marshal]::BindToMoniker('C:\work\Panes-monthly-close\October outputs\October 2026 close.xlsx');$xl=$w.Application
function Set-Table($sheet,$table,$rows,$headers){
  $s=$w.Worksheets.Item($sheet);$t=$s.ListObjects.Item($table)
  $t.DataBodyRange.ClearContents();$t.Resize($s.Range('A1').Resize($rows.Count+1,$headers.Count))
  for($j=0;$j -lt $headers.Count;$j++){if($headers[$j] -in @('Account','GLAccount','Entity')){$s.Range('A2').Offset(0,$j).Resize($rows.Count,1).NumberFormat='@'}}
  $data=New-Object 'object[,]' $rows.Count,$headers.Count
  for($i=0;$i -lt $rows.Count;$i++){for($j=0;$j -lt $headers.Count;$j++){
    $val=$rows[$i].($headers[$j]);$data[$i,$j]=[string]$val
    if($headers[$j] -in @('Debit','Credit','Opening','Debits','Credits','Ending')){$data[$i,$j]=[double]$val}
  }}
  $s.Range('A2').Resize($rows.Count,$headers.Count).Value2=$data
  return
}
$raw=@(Import-Csv ($src+'\GL detail 2026-10.csv'))
Set-Table 'GL Import' 'GL' $raw @('PostingDate','JournalID','Entity','Department','GLAccount','Description','Debit','Credit','Memo','Source','Line','Type')
$s=$w.Worksheets.Item('GL Import')
$s.Range('K2:K'+($raw.Count+1)).Formula='=IFERROR(INDEX(Map[Line],MATCH(E2,Map[Account],0)),"UNMAPPED")'
$s.Range('L2:L'+($raw.Count+1)).Formula='=IFERROR(INDEX(Map[Type],MATCH(E2,Map[Account],0)),"UNMAPPED")'
$raw=@(Import-Csv ($src+'\Trial balance 2026-10.csv'))
Set-Table 'TB Import' 'TB' $raw @('Entity','Department','Account','Description','Opening','Debits','Credits','Ending','Roll check','Line','Type')
$s=$w.Worksheets.Item('TB Import')
$s.Range('I2:I'+($raw.Count+1)).Formula='=E2+F2-G2-H2'
$s.Range('J2:J'+($raw.Count+1)).Formula='=IFERROR(INDEX(Map[Line],MATCH(C2,Map[Account],0)),"UNMAPPED")'
$s.Range('K2:K'+($raw.Count+1)).Formula='=IFERROR(INDEX(Map[Type],MATCH(C2,Map[Account],0)),"UNMAPPED")'
$xl.CalculateFullRebuild()
$w.Worksheets.Item('Checks').Activate();$xl.ActiveWindow.Zoom=100
Write-Output ('Imported October. Unmapped TB accounts: '+$w.Worksheets.Item('Checks').Range('B5').Value2)
$w.Save()

Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class CloseFocus{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);}'
$w.Activate();[CloseFocus]::SetForegroundWindow([IntPtr]$xl.Hwnd) | Out-Null
