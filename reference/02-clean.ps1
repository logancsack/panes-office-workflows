$ErrorActionPreference='Stop'
$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');$wb=$xl.ActiveWorkbook
$clean=$wb.Worksheets.Add();$clean.Name='Actuals'
$clean.Range('A1').Value2='Cleaned actuals';$clean.Range('A1').Font.Size=22;$clean.Range('A2').Value2='Typed amounts and dates. One row per unique document.'
$m=@'
let
 Source=Excel.CurrentWorkbook(){[Name="RawExport"]}[Content],
 Trimmed=Table.TransformColumns(Source,{{"DocID",each Text.Trim(Text.From(_)),type text},{"PostingDate",each Text.Trim(Text.From(_)),type text},{"CostCenter",each Text.Trim(Text.From(_)),type text},{"Account",each Text.Proper(Text.Trim(Text.From(_))),type text},{"Amount",each Text.Trim(Text.From(_)),type text}}),
 Typed=Table.TransformColumns(Trimmed,{{"PostingDate",each Date.FromText(_,[Culture="en-US"]),type date},{"Amount",each Number.FromText(Text.Remove(_,{"$",","," "}),"en-US"),Currency.Type}}),
 Unique=Table.Distinct(Typed,{"DocID"}),
 Department=Table.AddColumn(Unique,"Department",each Record.Field([#"100"="Sales",#"200"="Engineering",#"300"="G&A"],[CostCenter]),type text),
 Period=Table.AddColumn(Department,"Period",each Date.ToText([PostingDate],"yyyy-MM"),type text),
 Renamed=Table.RenameColumns(Period,{{"Account","Category"}}),
 Output=Table.SelectColumns(Renamed,{"DocID","PostingDate","Period","Department","Category","Amount"})
in Output
'@
$q=$wb.Queries.Add('Clean_Actuals',$m)
$conn='OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Clean_Actuals;Extended Properties=""'
$lo=$clean.ListObjects.Add(0,$conn,[Type]::Missing,1,$clean.Range('A4'))
$lo.Name='Actuals';$lo.TableStyle='TableStyleMedium2'
$lo.QueryTable.CommandType=2;$lo.QueryTable.CommandText='SELECT * FROM [Clean_Actuals]';$lo.QueryTable.BackgroundQuery=$false
$lo.QueryTable.Refresh($false)|Out-Null
$xl.CalculateUntilAsyncQueriesDone()
$clean.Columns.AutoFit()|Out-Null;$clean.Columns.Item(2).NumberFormat='mmm d, yyyy';$clean.Columns.Item(6).NumberFormat='#,##0.00;(#,##0.00);"–"'
$clean.Activate();$clean.Range('A5').Select()|Out-Null;$xl.ActiveWindow.FreezePanes=$true;$xl.ActiveWindow.Zoom=110
$wb.Save()
Write-Output ('Cleaned rows '+$lo.ListRows.Count+'; sum '+$xl.WorksheetFunction.Sum($lo.ListColumns.Item('Amount').DataBodyRange))
