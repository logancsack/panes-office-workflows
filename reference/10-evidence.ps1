$ErrorActionPreference='Stop'
$d='C:\work\Panes-office-demo\review';$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');$wb=$xl.ActiveWorkbook;$p=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application');$deck=$p.ActivePresentation
# Embed both native charts so the delivered deck works without a local source path.
foreach($s in @($deck.Slides.Item(1),$deck.Slides.Item(2))){foreach($sh in $s.Shapes){if($sh.HasChart -eq -1){if($sh.Chart.ChartData.IsLinked){throw 'Expected embedded chart'};Write-Output ($sh.Name+' linked='+$sh.Chart.ChartData.IsLinked)}}}
# Present the report first when the workbook is opened.
$wb.Worksheets.Item('Report').Move($wb.Worksheets.Item(1))
$wb.Worksheets.Item('Read me').Move([Type]::Missing,$wb.Worksheets.Item('Report'))
$wb.Worksheets.Item('Report').Activate();$xl.ActiveWindow.ScrollRow=1;$xl.ActiveWindow.ScrollColumn=1
$xl.CalculateFullRebuild();$wb.Save();$deck.Save()
$evidence=@{excelVersion=$xl.Version;powerPointVersion=$p.Version;report=@{};checks=@{};queries=@();charts=@()}
$r=$wb.Worksheets.Item('Report')
foreach($row in @(6,7,8,9,14,15,16,17,23,24,25,26)){foreach($col in @(1,2,3,4,5,6)){$cell=$r.Cells.Item($row,$col);$evidence.report[$cell.Address($false,$false)]=@{value=$cell.Value2;formula=$cell.Formula}}}
foreach($row in 3..10){$evidence.checks['B'+$row]=$wb.Worksheets.Item('Checks').Cells.Item($row,2).Value2}
foreach($q in $wb.Queries){$evidence.queries+=@{name=$q.Name;formula=$q.Formula}}
foreach($s in @($deck.Slides.Item(1),$deck.Slides.Item(2))){foreach($sh in $s.Shapes){if($sh.HasChart -eq -1){$entry=@{slide=$s.SlideIndex;name=$sh.Name;linked=$sh.Chart.ChartData.IsLinked;series=@()};$sc=$sh.Chart.SeriesCollection();for($j=1;$j -le $sc.Count;$j++){$ser=$sc.Item($j);$entry.series+=@{name=$ser.Name;values=@($ser.Values);categories=@($ser.XValues)}};$evidence.charts+=$entry}}}
$evidence|ConvertTo-Json -Depth 12|Set-Content -Encoding UTF8 ($d+'\office-evidence.json')
Write-Output 'Workbook and deck saved; chart data and formula evidence exported'
