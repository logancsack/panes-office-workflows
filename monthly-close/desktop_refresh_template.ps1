$ErrorActionPreference='Stop';$d='C:\work\Panes-monthly-close';$p=New-Object -ComObject PowerPoint.Application;$p.Visible=-1;$n=0
foreach($f in (Get-ChildItem ($d+'\Alder Mechanical - Finance\FY2026') -Filter '*.pptx' -Recurse)){
 if($f.Name -notmatch '2026-(0[2-9]|1[0-2])'){continue}
 $deck=$p.Presentations.Open($f.FullName,0,0,0)
 foreach($sl in $deck.Slides){foreach($sh in $sl.Shapes){
  if($sh.HasChart){$chart=$sh.Chart;for($i=1;$i -le $chart.SeriesCollection().Count;$i++){$ser=$chart.SeriesCollection($i);$ser.Format.Fill.Solid();$ser.Format.Fill.ForeColor.RGB=if($i -eq 1){4798744}else{8093440};$ser.Format.Line.ForeColor.RGB=if($i -eq 1){4798744}else{8093440}}}
  if($sh.HasTable){for($j=1;$j -le $sh.Table.Columns.Count;$j++){$cell=$sh.Table.Cell(1,$j);$cell.Shape.Fill.ForeColor.RGB=4798744;$cell.Shape.TextFrame.TextRange.Font.Color.RGB=16777215}}
 }}
 $deck.Save();$deck.Close();$n++
}
Write-Output ('February 2026 template refresh applied to '+$n+' saved monthly presentations')
