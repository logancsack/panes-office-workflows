$ErrorActionPreference='Stop'
$p=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application');$deck=$p.ActivePresentation;$s=$deck.Slides.Item(2);$ch=$s.Shapes.Item($s.Shapes.Count)
if($ch.HasChart -ne -1 -or $ch.Chart.ChartData.IsLinked){throw 'Expected a native embedded chart'}
$ch.Name='VarianceChart';$ch.LockAspectRatio=0;$ch.Left=36;$ch.Top=155;$ch.Width=585;$ch.Height=330;$p.ActiveWindow.Selection.Unselect();$deck.Save()
foreach($i in @(1,2)){$deck.Slides.Item($i).Export(('C:\work\Panes-office-demo\slide-'+$i+'.png'),'PNG',1600,900)}
Write-Output 'Two native embedded charts saved'
