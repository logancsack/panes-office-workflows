$ErrorActionPreference='Stop'
$d='C:\work\Panes-monthly-close';$drive=$d+'\Alder Mechanical - Finance'
$sept=Get-ChildItem ($drive+'\FY2026\09 Sep\04 Close and review') -Filter '*.xlsx' | Select-Object -First 1
$out=$d+'\October outputs';New-Item -ItemType Directory -Path $out -Force | Out-Null
$xl=New-Object -ComObject Excel.Application;$xl.Visible=$true;$xl.DisplayAlerts=$false;$xl.WindowState=-4137
$w=$xl.Workbooks.Open($sept.FullName,0,$false)
$w.Worksheets.Item('PL Group').Activate();$xl.ActiveWindow.Zoom=85
$w.SaveAs(($out+'\October 2026 close.xlsx'),51)
# Freeze September's mapped GL analysis into its historical snapshot before replacing the current import.
$h=$w.Worksheets.Item('Monthly History');$range=$h.ListObjects.Item('Facts').DataBodyRange;$range.Value2=$range.Value2
$w.Worksheets.Item('Control').Range('B2').Value2=[double]([datetime]'2026-10-31').ToOADate()
$w.Worksheets.Item('Control').Range('B5').Value2='In progress'
$w.Worksheets.Item('Control').Range('B6').Value2='FY2026/10 Oct'
$w.Worksheets.Item('Control').Range('B7').Value2='October roll-forward'
$w.Worksheets.Item('Control').Activate();$xl.ActiveWindow.Zoom=100
$w.Save()
Write-Output 'September preserved; October workbook rolled forward.'

Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class CloseFocus{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);}'
$w.Activate();[CloseFocus]::SetForegroundWindow([IntPtr]$xl.Hwnd) | Out-Null
