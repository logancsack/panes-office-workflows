$ErrorActionPreference='Stop';$d='C:\work\Panes-monthly-close\Delivery verification'
$xl=New-Object -ComObject Excel.Application;$xl.DisplayAlerts=$false;$xl.Visible=$false;$xl.AskToUpdateLinks=$false
$p=New-Object -ComObject PowerPoint.Application;$p.Visible=-1;$results=@()
foreach($f in (Get-ChildItem $d -Filter '*.xlsx' -Recurse)){
 $before=(Get-FileHash $f.FullName -Algorithm SHA256).Hash
 $w=$xl.Workbooks.Open($f.FullName,0,$true);$xl.CalculateFullRebuild();$errors=@()
 foreach($s in $w.Worksheets){try{$bad=$s.UsedRange.SpecialCells(-4123,16);foreach($c in $bad.Cells){$errors+=($s.Name+'!'+$c.Address())}}catch{}}
 $checks=@();$s=$w.Worksheets.Item('Checks');for($i=2;$i -le 15;$i++){$checks+=@{name=$s.Cells.Item($i,1).Text;result=$s.Cells.Item($i,3).Text;difference=$s.Cells.Item($i,2).Value2}}
 $w.Close($false);$after=(Get-FileHash $f.FullName -Algorithm SHA256).Hash
 if($before -ne $after -or $errors.Count -or ($checks|Where-Object {$_.result -ne 'PASS'})){throw ('Final workbook verification failed: '+$f.Name)}
 $results+=@{file=$f.FullName.Substring($d.Length+1);application='Microsoft Excel';version=$xl.Version;opened=$true;formulaErrors=$errors;checks=$checks;sha256=$after.ToLower();unchangedByVerification=$true}
}
foreach($f in (Get-ChildItem $d -Filter '*.pptx' -Recurse)){
 $before=(Get-FileHash $f.FullName -Algorithm SHA256).Hash;$deck=$p.Presentations.Open($f.FullName,-1,0,0);$charts=0;$links=0
 foreach($sl in $deck.Slides){foreach($sh in $sl.Shapes){if($sh.HasChart){$charts++;if($sh.Chart.ChartData.IsLinked){$links++}}}}
 $slides=$deck.Slides.Count;$deck.Close();$after=(Get-FileHash $f.FullName -Algorithm SHA256).Hash
 if($before -ne $after -or $links -or $charts -ne 3 -or $slides -ne 6){throw ('Final deck verification failed: '+$f.Name)}
 $results+=@{file=$f.FullName.Substring($d.Length+1);application='Microsoft PowerPoint';version=$p.Version;opened=$true;editableCharts=$charts;externalLinks=$links;slides=$slides;sha256=$after.ToLower();unchangedByVerification=$true}
}
$xl.Quit();$p.Quit();$results|ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 ($d+'\verified-deliverables.json')
Write-Output ('Final byte-identical deliverables opened and verified in desktop Office: '+$results.Count)
