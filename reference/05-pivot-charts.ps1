$ErrorActionPreference='Stop'
$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');$wb=$xl.ActiveWorkbook;$s=$wb.Worksheets.Item('Report')
$s.Range('A11').Value2='Year to date | January-October'
foreach($r in @(5,13,22)){$s.Cells.Item($r,4).Value2='$ variance';$s.Cells.Item($r,5).Value2='% variance';$s.Range('A'+$r+':F'+$r).Font.Size=12}
$s.Range('A1').Formula='=TEXT(DATEVALUE($B$3&"-01"),"mmmm yyyy")&" close"'
$p=$wb.Worksheets.Add();$p.Name='Pivot detail';$p.Range('A1').Value2='October expense detail';$p.Range('A1').Font.Size=24;$p.Range('A1').Font.Bold=$true
$pc=$wb.PivotCaches().Create(1,'Actuals');$pt=$pc.CreatePivotTable($p.Range('A4'),'ExpensePivot')
$pt.PivotFields('Department').Orientation=1;$pt.PivotFields('Department').Position=1
$pt.PivotFields('Category').Orientation=1;$pt.PivotFields('Category').Position=2
$pt.PivotFields('Period').Orientation=3;$pt.PivotFields('Period').CurrentPage='2026-10'
$field=$pt.AddDataField($pt.PivotFields('Amount'),'Actual spend',-4157);$field.NumberFormat='#,##0;(#,##0);"–"'
$pt.RowAxisLayout(1);$pt.TableStyle2='PivotStyleMedium9';$p.UsedRange.Font.Name='Aptos';$p.UsedRange.Font.Size=14;$p.Columns.AutoFit()|Out-Null
$checks=$wb.Worksheets.Add();$checks.Name='Checks';$checks.Range('A1').Value2='Reconciliation';$checks.Range('A1').Font.Size=24;$checks.Range('A1').Font.Bold=$true
$labels=@('Raw document rows','Clean document rows','Duplicate rows removed','Monthly actual tie-out','Monthly budget tie-out','YTD actual tie-out','Category variance tie-out','Pivot total tie-out')
$formulas=@('=ROWS(RawExport[DocID])','=ROWS(Actuals[DocID])','=B3-B4','=SUMIFS(Actuals[Amount],Actuals[Period],Report!B3)-Report!C9','=SUMIFS(Plan[Budget],Plan[Period],Report!B3)-Report!B9','=SUMIFS(Actuals[Amount],Actuals[Period],">="&LEFT(Report!B3,4)&"-01",Actuals[Period],"<="&Report!B3)-Report!C17','=SUM(Report!D23:D26)-Report!D9','=GETPIVOTDATA("Actual spend",''Pivot detail''!A4)-Report!C9')
for($i=0;$i -lt $labels.Count;$i++){$r=$i+3;$checks.Cells.Item($r,1).Value2=[string]$labels[$i];$checks.Cells.Item($r,2).Formula=[string]$formulas[$i]}
$checks.Range('A12').Value2='All five financial tie-outs must be zero. Duplicate removal must be three.';$checks.Columns.Item(1).ColumnWidth=77;$checks.Columns.Item(2).ColumnWidth=18;$checks.Range('B3:B10').NumberFormat='#,##0.00;(#,##0.00);"–"';$checks.UsedRange.Font.Name='Aptos';$checks.Range('A3:B12').Font.Size=14
$charts=$wb.Worksheets.Add();$charts.Name='Charts';$charts.Range('A1').Value2='October | Budget vs actual';$charts.Range('A1').Font.Size=24;$charts.Range('A1').Font.Bold=$true
$co=$charts.ChartObjects().Add(12,60,800,440);$co.Name='DepartmentChart';$ch=$co.Chart;$ch.ChartType=51;$ch.SetSourceData($s.Range('A5:C8'));$ch.HasTitle=$true;$ch.ChartTitle.Text='Operating expense by department';$ch.ChartTitle.Font.Size=20;$ch.Legend.Position=-4107;$ch.Legend.Font.Size=13;$ch.Axes(1).TickLabels.Font.Size=14;$ch.Axes(2).TickLabels.Font.Size=12;$ch.Axes(2).TickLabels.NumberFormat='$0,"k"';$ch.Axes(2).MinimumScale=0
$ch.SeriesCollection(1).Format.Fill.ForeColor.RGB=12632256;$ch.SeriesCollection(2).Format.Fill.ForeColor.RGB=6171950
for($i=1;$i -le 2;$i++){$ch.SeriesCollection($i).ApplyDataLabels();$ch.SeriesCollection($i).DataLabels().NumberFormat='$0.0,"k"';$ch.SeriesCollection($i).DataLabels().Font.Size=12}
$co2=$charts.ChartObjects().Add(12,540,800,400);$co2.Name='VarianceChart';$ch=$co2.Chart;$ch.ChartType=57;$series=$ch.SeriesCollection().NewSeries();$series.Name='Over budget';$series.XValues=$s.Range('A23:A26');$series.Values=$s.Range('D23:D26');$series.Format.Fill.ForeColor.RGB=6171950;$ch.HasLegend=$false;$ch.HasTitle=$true;$ch.ChartTitle.Text='October overspend by category';$ch.ChartTitle.Font.Size=20;$ch.Axes(1).TickLabels.Font.Size=14;$ch.Axes(2).TickLabels.NumberFormat='$0,"k"';$ch.Axes(2).MinimumScale=0;$series.ApplyDataLabels();$series.DataLabels().NumberFormat='$0.0,"k"';$series.DataLabels().Font.Size=13
$xl.CalculateFullRebuild();$wb.Save()
$p.Activate();$xl.ActiveWindow.DisplayGridlines=$false;$xl.ActiveWindow.Zoom=115;$p.Range('A4').Select()|Out-Null
Write-Output ('Pivot October total '+$pt.GetPivotData('Actual spend').Value2)
Write-Output ('Monthly variance '+$s.Range('D9').Value2+' / '+$s.Range('E9').Value2)
