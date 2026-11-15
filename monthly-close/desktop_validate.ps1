$ErrorActionPreference='Stop'
$d='C:\work\Panes-monthly-close'
$xl=New-Object -ComObject Excel.Application
$xl.Visible=$true;$xl.DisplayAlerts=$false;$xl.AskToUpdateLinks=$false;$xl.WindowState=-4137
$p=New-Object -ComObject PowerPoint.Application
$p.Visible=-1
$results=@()
$files=Get-ChildItem ($d+'\Alder Mechanical - Finance') -Filter '*.xlsx' -Recurse | Sort-Object FullName
foreach($f in $files){
  $w=$xl.Workbooks.Open($f.FullName,0,$false)
  $xl.CalculateFullRebuild()
  $errors=@();$formulas=0
  foreach($s in $w.Worksheets){
    try {$bad=$s.UsedRange.SpecialCells(-4123,16);foreach($c in $bad.Cells){$errors+=($s.Name+'!'+$c.Address()+': '+$c.Text)}} catch {}
    try {$fc=$s.UsedRange.SpecialCells(-4123);$formulas+=$fc.Count} catch {}
  }
  $checks=@()
  $s=$w.Worksheets.Item('Checks')
  for($i=2;$i -le $s.UsedRange.Rows.Count;$i++){$checks+=@{name=$s.Cells.Item($i,1).Text;difference=$s.Cells.Item($i,2).Value2;result=$s.Cells.Item($i,3).Text}}
  $results+=@{file=$f.FullName.Substring($d.Length+1);application='Microsoft Excel';version=$xl.Version;opened=$true;formulaCount=$formulas;formulaErrors=$errors;checks=$checks;revenue=$w.Worksheets.Item('PL Group').Range('B16').Value2;netIncome=$w.Worksheets.Item('PL Group').Range('B18').Value2}
  $w.Save();$w.Close($false)
  $results|ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 ($d+'\desktop-validation.json')
}
foreach($f in (Get-ChildItem ($d+'\Alder Mechanical - Finance') -Filter '*.pptx' -Recurse | Sort-Object FullName)){
  $deck=$p.Presentations.Open($f.FullName,0,0,0)
  $charts=0;$linked=0
  foreach($sl in $deck.Slides){foreach($shape in $sl.Shapes){if($shape.HasChart){$charts++;if($shape.Chart.ChartData.IsLinked){$linked++}}}}
  $results+=@{file=$f.FullName.Substring($d.Length+1);application='Microsoft PowerPoint';version=$p.Version;opened=$true;slides=$deck.Slides.Count;editableCharts=$charts;linkedCharts=$linked}
  $deck.Save();$deck.Close()
  $results|ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 ($d+'\desktop-validation.json')
}
$xl.Quit();$p.Quit()
Write-Output ('Validated '+$results.Count+' Office files')
