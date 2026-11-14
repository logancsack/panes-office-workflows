$ErrorActionPreference='Stop'
$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');$wb=$xl.ActiveWorkbook
$charts=$wb.Worksheets.Item('Charts');$charts.Activate();$xl.ActiveWindow.DisplayGridlines=$false;$xl.ActiveWindow.Zoom=90
$charts.ChartObjects('DepartmentChart').Activate()|Out-Null
$ws=New-Object -ComObject WScript.Shell;$ws.AppActivate('October close')|Out-Null
Write-Output 'Department chart selected in Excel'
