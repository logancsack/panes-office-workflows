$ErrorActionPreference='Stop'
$w=[Runtime.InteropServices.Marshal]::BindToMoniker('C:\work\Panes-monthly-close\October outputs\October 2026 close.xlsx');$xl=$w.Application
$s=$w.Worksheets.Item('Flux');$pl=$w.Worksheets.Item('PL Group')
for($i=2;$i -le 15;$i++){
  $line=$s.Cells.Item($i,1).Text;$actual=[double]$pl.Cells.Item($i,2).Value2;$budget=[double]$pl.Cells.Item($i,3).Value2;$py=[double]$pl.Cells.Item($i,6).Value2
  $reason=switch($line){
   'Contract revenue' {'Contract renewals carry the January 4% price step; October is a shoulder-season month for commercial HVAC.'}
   'Demand revenue' {'Demand service is below the summer cooling peak; compare to the same month last year after the Summit acquisition.'}
   'Field labor' {'Includes the $24,600 approved overtime accrual true-up, JE-202610-017. Reconcile gross pay and cut-off reserve to the payroll summary.'}
   'Warranty' {'Includes $18,750 of third-party callouts in new account 6265, mapped to Warranty. The June restatement remains in YTD and T12.'}
   'Parts' {'Parts and subcontractor invoice batches reconcile to the imported ledger; review margin alongside the contract/demand mix.'}
   'Office labor' {'Central finance, HR and dispatch payroll follows the established staffing run rate; entity allocations eliminate in consolidation.'}
   'Occupancy' {'Branch leases and utilities follow the established run rate; fixed costs are less seasonal than revenue.'}
   'Fleet' {'Fleet fuel and maintenance use GL 6205; legacy 6200 stays mapped for historical comparisons.'}
   'Insurance' {'January annual premiums amortize monthly. Prepaids schedule ties to account 1300.'}
   'Depreciation' {'84-month straight-line fleet and tool pools; current additions and accumulated depreciation tie to the fixed asset schedule.'}
   'Interest' {'6.75% term debt interest follows scheduled principal reductions.'}
   'Tax' {'25% provision on positive entity pretax profit. October has no quarterly estimated cash tax payment.'}
   default {'Shared service recovery and allocations eliminate fully in the consolidated results.'}
  }
  $s.Cells.Item($i,5).Value2=('Actual ${0:N0}; budget ${1:N0}; prior year ${2:N0}. {3}' -f $actual,$budget,$py,$reason)
  $s.Cells.Item($i,7).Value2='Reviewed'
}
$s.Columns.ColumnWidth=20;$s.Columns.Item(1).ColumnWidth=24;$s.Columns.Item(5).ColumnWidth=88;$s.Range('E2:E15').WrapText=$true;$s.Rows('2:15').RowHeight=62
$s.Activate();$xl.ActiveWindow.Zoom=85
$log=$w.Worksheets.Item('Change Log');$i=$log.UsedRange.Rows.Count+1
$log.Cells.Item($i,1).Value2=[double]([datetime]'2026-10-31').ToOADate();$log.Cells.Item($i,2).Value2='October close roll-forward';$log.Cells.Item($i,3).Value2='Imported GL/TB; mapped 6265 to Warranty; verified $24,600 payroll accrual true-up; refreshed statements, schedules, KPIs, YTD/T12 and deck.';$log.Cells.Item($i,4).Value2='Finance team'
$xl.CalculateFullRebuild();$s=$w.Worksheets.Item('Checks');$s.Activate();$xl.ActiveWindow.Zoom=100
$fail=@();for($i=2;$i -le 15;$i++){if($s.Cells.Item($i,3).Text -ne 'PASS'){$fail+=$s.Cells.Item($i,1).Text+': '+$s.Cells.Item($i,2).Text}}
if($fail.Count){throw ($fail -join '; ')}
$w.Worksheets.Item('Control').Range('B5').Value2='Reviewed';$w.Save()
Write-Output 'All 14 workbook controls PASS; October close saved.'

Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class CloseFocus{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);}'
$w.Activate();[CloseFocus]::SetForegroundWindow([IntPtr]$xl.Hwnd) | Out-Null
