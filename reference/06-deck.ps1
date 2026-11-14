$ErrorActionPreference='Stop'
$d='C:\work\Panes-office-demo\review'
$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');$wb=$xl.ActiveWorkbook;$r=$wb.Worksheets.Item('Report')
$ppt=New-Object -ComObject PowerPoint.Application;$ppt.Visible=-1;$ppt.WindowState=3
$deck=$ppt.Presentations.Add();$deck.PageSetup.SlideWidth=960;$deck.PageSetup.SlideHeight=540
$s1=$deck.Slides.Add(1,12);$s2=$deck.Slides.Add(2,12)
function Text($s,$name,$text,$x,$y,$w,$h,$size,$bold=$false,$color=2434341){
 $sh=$s.Shapes.AddTextbox(1,[single]$x,[single]$y,[single]$w,[single]$h);$sh.Name=$name
 $sh.TextFrame.MarginLeft=0;$sh.TextFrame.MarginRight=0;$sh.TextFrame.MarginTop=0;$sh.TextFrame.MarginBottom=0
 $sh.TextFrame.TextRange.Text=[string]$text;$sh.TextFrame.TextRange.Font.Name='Aptos';$sh.TextFrame.TextRange.Font.Size=[single]$size;$sh.TextFrame.TextRange.Font.Bold=[int](-[int]$bold);$sh.TextFrame.TextRange.Font.Color.RGB=[int]$color
 return $sh
}
$actual=[double]$r.Range('C9').Value2;$budget=[double]$r.Range('B9').Value2;$variance=[double]$r.Range('D9').Value2;$percent=[double]$r.Range('E9').Value2;$cloud=[double]$r.Range('D24').Value2;$travel=[double]$r.Range('D25').Value2
Text $s1 'Title' ('October costs finished $'+($variance/1000).ToString('0.0')+'k above plan') 36 22 890 72 30 $true | Out-Null
Text $s1 'Subtitle' 'October 2026 | Operating expense | USD' 36 105 890 25 16 $false 7368816 | Out-Null
Text $s1 'ActualLabel' 'ACTUAL' 660 175 250 25 14 | Out-Null
Text $s1 'Actual' ('$'+($actual/1000).ToString('0.0')+'k') 660 205 250 50 38 $true | Out-Null
Text $s1 'Budget' ('Budget: $'+($budget/1000).ToString('0.0')+'k') 660 270 250 30 20 | Out-Null
Text $s1 'Variance' ('$'+($variance/1000).ToString('0.0')+'k over budget') 660 335 270 35 24 $true 192 | Out-Null
Text $s1 'VariancePercent' ($percent.ToString('0.0%')+' unfavorable') 660 377 250 28 20 $false 192 | Out-Null
Text $s2 'Title' ('Cloud contributes $'+($cloud/1000).ToString('0.0')+'k to the overspend') 36 22 890 72 30 $true | Out-Null
Text $s2 'Subtitle' 'October 2026 | Actual less budget | USD' 36 105 890 25 16 $false 7368816 | Out-Null
Text $s2 'Cloud' (($cloud/$variance).ToString('0.0%')+' of the variance is cloud spend.') 660 185 260 75 24 $true | Out-Null
Text $s2 'Travel' ('Travel adds $'+($travel/1000).ToString('0.0')+'k.') 660 285 260 50 22 | Out-Null
Text $s2 'Action' 'Review cloud consumption and travel bookings with cost owners.' 660 370 260 95 22 | Out-Null
$s1.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text='Source: October close.xlsx, Report!B9:E9 and DepartmentChart. Calculated in real Excel. Positive expense variance is unfavorable. Native editable chart is embedded from the workbook. Synthetic demonstration data.'
$s2.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text='Source: October close.xlsx, Report!A23:F26. Cloud share = Report!D24/Report!D9. This is an arithmetic attribution, not an asserted operational cause. Synthetic demonstration data.'
$deck.SaveAs(($d+'\October review.pptx'),24)
$deck.Windows.Item(1).ViewType=1;$deck.Windows.Item(1).View.GotoSlide(1)
Write-Output 'PowerPoint summary created; chart areas ready for desktop paste'
