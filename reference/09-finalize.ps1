$ErrorActionPreference='Stop'
$d='C:\work\Panes-office-demo'
$p=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application');$deck=$p.ActivePresentation;$s=$deck.Slides.Item(2);$ch=$s.Shapes.Item($s.Shapes.Count)
if ($ch.HasChart -ne -1){throw 'Second clipboard chart is not editable'}
$ch.Name='VarianceChart';$ch.LockAspectRatio=0;$ch.Left=36;$ch.Top=155;$ch.Width=585;$ch.Height=330
$p.ActiveWindow.Selection.Unselect();$deck.Save()
foreach($i in @(1,2)){$deck.Slides.Item($i).Export(($d+'\slide-'+$i+'.png'),'PNG',1600,900)}
$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');$wb=$xl.ActiveWorkbook
$read=$wb.Worksheets.Item('Read me');$read.UsedRange.Replace('Januaryâ€“October','January-October',2)|Out-Null
$xl.CalculateFullRebuild();$wb.Save()
$wb.Worksheets.Item('Report').Activate();$xl.ActiveWindow.ScrollRow=1;$xl.ActiveWindow.ScrollColumn=1;$wb.Worksheets.Item('Report').Range('A1').Select();$xl.ActiveWindow.Zoom=100
$w=New-Object -ComObject WScript.Shell;$w.AppActivate('October close')|Out-Null
Write-Output ('Excel '+$xl.Version+'; PowerPoint '+$p.Version+'; native charts '+$deck.Slides.Item(1).Shapes.Item('DepartmentChart').HasChart+','+$deck.Slides.Item(2).Shapes.Item('VarianceChart').HasChart)
