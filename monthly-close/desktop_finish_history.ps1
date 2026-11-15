$ErrorActionPreference='Stop'
$d='C:\work\Panes-monthly-close';$base=$d+'\Alder Mechanical - Finance'
$x=New-Object -ComObject Excel.Application;$x.Visible=$false;$x.DisplayAlerts=$false;$x.AskToUpdateLinks=$false
$opened=@()
foreach($f in (Get-ChildItem $base -Filter '*.csv' -Recurse)){
 $w=$x.Workbooks.Open($f.FullName,0,$true)
 $opened+=@{file=$f.FullName.Substring($d.Length+1);application='Microsoft Excel';version=$x.Version;opened=$true;rows=$w.Worksheets.Item(1).UsedRange.Rows.Count;columns=$w.Worksheets.Item(1).UsedRange.Columns.Count}
 $w.Close($false)
}
$opened|ConvertTo-Json -Depth 4|Set-Content -Encoding UTF8 ($d+'\csv-open-validation.json')
foreach($f in (Get-ChildItem $base -Filter '*.xlsx' -Recurse)){
 $w=$x.Workbooks.Open($f.FullName,0,$false)
 $s=$w.Worksheets.Item('PL Group');$s.Range('E2:E18').NumberFormat='0.0%';$s.Columns.Item(1).ColumnWidth=26;$s.Range('B:H').ColumnWidth=15;$s.Range('A16:H18').Font.Bold=$true
 $s=$w.Worksheets.Item('KPI');$s.Range('C2:C9').WrapText=$true;$s.Rows('2:9').RowHeight=30
 $s=$w.Worksheets.Item('Checks');$s.Columns.Item(1).ColumnWidth=36;$s.Columns.Item(2).ColumnWidth=16;$s.Columns.Item(3).ColumnWidth=16
 $s.Range('C2:C15').Font.Color=32768
 $x.CalculateFullRebuild();$w.Save();$w.Close($false)
}
$x.Quit();Write-Output ('Opened all '+$opened.Count+' CSV source files in Excel; historical workbook display formats corrected.')
