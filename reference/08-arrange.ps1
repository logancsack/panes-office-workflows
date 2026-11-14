$ErrorActionPreference='Stop'
$p=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application');$deck=$p.ActivePresentation
$s=$deck.Slides.Item(1);$chart=$s.Shapes.Item($s.Shapes.Count)
if ($chart.HasChart -ne -1){throw 'Clipboard did not produce an editable chart'}
$chart.Name='DepartmentChart';$chart.LockAspectRatio=0;$chart.Left=36;$chart.Top=155;$chart.Width=585;$chart.Height=330
$s.Shapes.Item('Title').TextFrame.TextRange.Text='October costs finished $41.5k above plan'
$p.ActiveWindow.Selection.Unselect();$deck.Save()
$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');$wb=$xl.ActiveWorkbook;$charts=$wb.Worksheets.Item('Charts');$charts.Activate();$charts.Range('A34').Select();$charts.ChartObjects('VarianceChart').Activate()|Out-Null
$w=New-Object -ComObject WScript.Shell;$w.AppActivate('October close')|Out-Null
Write-Output 'First editable chart placed; second chart selected'
