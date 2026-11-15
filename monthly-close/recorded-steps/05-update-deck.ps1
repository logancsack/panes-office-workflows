$ErrorActionPreference='Stop'
$d='C:\work\Panes-monthly-close';$out=$d+'\October outputs';$drive=$d+'\Alder Mechanical - Finance'
$w=[Runtime.InteropServices.Marshal]::BindToMoniker('C:\work\Panes-monthly-close\October outputs\October 2026 close.xlsx');$xl=$w.Application
$p=New-Object -ComObject PowerPoint.Application;$p.Visible=-1;$p.WindowState=3
$old=Get-ChildItem ($drive+'\FY2026\09 Sep\04 Close and review') -Filter '*.pptx' | Select-Object -First 1
$deck=$p.Presentations.Open($old.FullName,0,0,-1);$deck.SaveAs(($out+'\October 2026 review.pptx'))
foreach($sl in $deck.Slides){foreach($sh in $sl.Shapes){if($sh.HasTextFrame -and $sh.TextFrame.HasText){$sh.TextFrame.TextRange.Text=$sh.TextFrame.TextRange.Text.Replace('September 2026','October 2026')}}}
$pl=$w.Worksheets.Item('PL Group');$rev=[double]$pl.Range('B16').Value2;$ebitda=[double]$pl.Range('B17').Value2;$ni=[double]$pl.Range('B18').Value2;$cash=[double]$w.Worksheets.Item('Balance Sheet').Range('B2').Value2;$head=[int]$w.Worksheets.Item('KPI').Range('B5').Value2
function SetText($sl,$shape,$text){$deck.Slides.Item($sl).Shapes.Item($shape).TextFrame.TextRange.Text=[string]$text}
# Existing layout contains header/title/subtitle/footer then chart and commentary.
SetText 1 2 'October 2026: operating performance'
SetText 1 3 ('Revenue ${0:N2}m   |   EBITDA ${1:N0}k   |   EBITDA margin {2:P1}' -f ($rev/1e6),($ebitda/1000),($ebitda/$rev))
SetText 1 6 'October is a shoulder-season close. Contract pricing carries the January 4% increase; Summit remains in the comparative year.'
# Shapes are found by type so the template can retain its layout and native chart objects.
$series=@();$margins=@();$cashpoints=@();$cats=@();$t12=$w.Worksheets.Item('Trailing 12')
for($j=2;$j -le 13;$j++){
  $dt=[datetime]::FromOADate([double]$t12.Cells.Item(1,$j).Value2);$cats+= $dt.ToString('MMM yy')
  $r=[double]$t12.Cells.Item(2,$j).Value2+[double]$t12.Cells.Item(3,$j).Value2+[double]$t12.Cells.Item(4,$j).Value2
  $cost=0.0;for($i=5;$i -le 11;$i++){$cost+=[double]$t12.Cells.Item($i,$j).Value2};$cost+=[double]$t12.Cells.Item(13,$j).Value2
  $e=$r-$cost;$series+=,@(($r/1000),($e/1000));$margins+=($e/$r)
  $monthFolder=$drive+'\FY'+$dt.Year+'\'+$dt.ToString('MM MMM')+'\01 ERP exports\Trial balance '+$dt.ToString('yyyy-MM')+'.csv'
  $tb=Import-Csv $monthFolder;$cv=0.0;foreach($row in $tb){if($row.Account -eq '1000'){$cv+=[double]$row.Ending}};$cashpoints+=($cv/1000)
}
function UpdateChart($slide,$kind){
  $chart=$null;foreach($sh in $deck.Slides.Item($slide).Shapes){if($sh.HasChart){$chart=$sh.Chart;break}}
  if(!$chart){throw 'Expected native chart on slide '+$slide}
  $chart.ChartData.Activate();$book=$chart.ChartData.Workbook;$sheet=$book.Worksheets.Item(1);$sheet.Range('A1:A13').NumberFormat='@'
  $sheet.Cells.Item(1,1).Value2='Month'
  if($kind -eq 'profit'){$sheet.Cells.Item(1,2).Value2='Revenue ($000)';$sheet.Cells.Item(1,3).Value2='EBITDA ($000)'}
  elseif($kind -eq 'margin'){$sheet.Cells.Item(1,2).Value2='EBITDA margin'}else{$sheet.Cells.Item(1,2).Value2='Closing cash ($000)'}
  for($i=0;$i -lt 12;$i++){
    $sheet.Cells.Item($i+2,1).Value2=[string]$cats[$i]
    if($kind -eq 'profit'){$sheet.Cells.Item($i+2,2).Value2=[double]$series[$i][0];$sheet.Cells.Item($i+2,3).Value2=[double]$series[$i][1]}
    elseif($kind -eq 'margin'){$sheet.Cells.Item($i+2,2).Value2=[double]$margins[$i]}else{$sheet.Cells.Item($i+2,2).Value2=[double]$cashpoints[$i]}
  }
  $last=if($kind -eq 'profit'){'C'}else{'B'}
  $chart.SetSourceData(('Sheet1!$A$1:$'+$last+'$13'))
  if($kind -eq 'profit'){$a=@();$b=@();for($q=0;$q -lt 12;$q++){$a+=[double]$series[$q][0];$b+=[double]$series[$q][1]};$chart.SeriesCollection(1).Values=$a;$chart.SeriesCollection(2).Values=$b}
  elseif($kind -eq 'margin'){$chart.SeriesCollection(1).Values=$margins}else{$chart.SeriesCollection(1).Values=$cashpoints}
  $seriesCount=if($kind -eq 'profit'){2}else{1};for($j=1;$j -le $seriesCount;$j++){$chart.SeriesCollection($j).XValues=[object[]]$cats}
  $book.Close($true);$chart.HasLegend=-1;$chart.Legend.Position=-4107;if($kind -eq 'margin'){$chart.Axes(2).TickLabels.NumberFormat='0%'};$chart.Refresh()
  if($chart.ChartData.IsLinked){throw 'External chart link found'}
}
UpdateChart 1 'profit'
$tbl=$null;foreach($sh in $deck.Slides.Item(2).Shapes){if($sh.HasTable){$tbl=$sh.Table;break}}
$vals=@(('$'+$rev.ToString('N0')),('$'+$ebitda.ToString('N0')),('$'+$ni.ToString('N0')),('$'+$cash.ToString('N0')),([string]$head))
for($i=0;$i -lt 5;$i++){$tbl.Cell($i+2,2).Shape.TextFrame.TextRange.Text=[string]$vals[$i]}
SetText 2 6 'New GL 6265 adds $18,750 of third-party warranty callouts. Field labor includes a $24,600 payroll accrual true-up. Both tie to the October ledger.'
SetText 3 5 'Mapped new GL 6265 to Warranty. The $18,750 callout invoices remain visible in the ledger and current-month analysis. Payroll JE-202610-017 accrues $24,600 of approved field overtime. The June warranty restatement stays in YTD and T12.'
UpdateChart 4 'margin'
SetText 4 6 ($head.ToString()+' employees at month end. Budget, prior year, T12 and YTD are refreshed in the close model. Seasonal service demand is compared with the same month last year.')
UpdateChart 5 'cash'
foreach($sh in $deck.Slides.Item(6).Shapes){if($sh.HasTable){$sh.Table.Cell(3,2).Shape.TextFrame.TextRange.Text='Review third-party warranty callouts and the $18,750 vendor batch';$sh.Table.Cell(4,2).Shape.TextFrame.TextRange.Text='Approve $24,600 overtime true-up and next payroll reversal'}}
foreach($sl in $deck.Slides){foreach($sh in $sl.Shapes){if($sh.HasTextFrame){$sh.TextFrame.WordWrap=-1}}}
$p.ActiveWindow.View.GotoSlide(1);$p.Activate()
$deck.Save();$w.Save();$deck.Close();$deck=$p.Presentations.Open(($out+'\October 2026 review.pptx'),0,0,-1);$p.Activate()
for($i=1;$i -le $deck.Slides.Count;$i++){$deck.Slides.Item($i).Export(($out+'\slide-'+$i+'.png'),'PNG',1600,900)}
Write-Output ('October presentation saved: '+$deck.Slides.Count+' slides, 3 editable embedded charts.')
