$ErrorActionPreference='Stop'
$d='C:\work\Panes-monthly-close';$out=$d+'\October outputs'
$w=[Runtime.InteropServices.Marshal]::BindToMoniker($out+'\October 2026 close.xlsx');$xl=$w.Application
$xl.CalculateFullRebuild();$errors=@();$formulas=0
foreach($s in $w.Worksheets){
  try{$bad=$s.UsedRange.SpecialCells(-4123,16);foreach($c in $bad.Cells){$errors+=($s.Name+'!'+$c.Address()+': '+$c.Text)}}catch{}
  try{$formulas+=$s.UsedRange.SpecialCells(-4123).Count}catch{}
}
$checks=@();$s=$w.Worksheets.Item('Checks');for($i=2;$i -le 15;$i++){$checks+=@{name=$s.Cells.Item($i,1).Text;difference=$s.Cells.Item($i,2).Value2;result=$s.Cells.Item($i,3).Text}}
$p=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application');$deck=$p.Presentations.Item('October 2026 review.pptx');$charts=0;$links=0
foreach($sl in $deck.Slides){foreach($sh in $sl.Shapes){if($sh.HasChart){$charts++;if($sh.Chart.ChartData.IsLinked){$links++}}}}
@{workbook=@{opened=$true;version=$xl.Version;formulaErrors=$errors;formulaCount=$formulas;checks=$checks};presentation=@{opened=$true;version=$p.Version;slides=$deck.Slides.Count;editableCharts=$charts;linkedCharts=$links}}|ConvertTo-Json -Depth 8|Set-Content -Encoding UTF8 ($out+'\october-desktop-validation.json')
if($errors.Count -or $links -or ($checks|Where-Object {$_.result -ne 'PASS'})){throw 'Final output verification failed'}
$w.Worksheets.Item('PL Group').Activate();$xl.ActiveWindow.Zoom=85;$w.Save();$deck.Save();$p.ActiveWindow.View.GotoSlide(1);$p.Activate()
Write-Output 'October workbook and deck verified in desktop Office.'
