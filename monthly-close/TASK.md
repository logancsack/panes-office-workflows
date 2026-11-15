# October 2026 monthly close

Work on a real Panes Windows computer with signed-in Microsoft desktop Excel and PowerPoint. Never simulate the apps or replace them with a browser or another spreadsheet engine. Keep the computer, profile, prior reports and source exports. Use a new task folder, and keep API credentials and viewer URLs private.

Alder Mechanical Services Group is a fictional regional commercial HVAC services group. All figures are synthetic. September 2023–September 2026 contains 37 closed monthly packs, plus earlier reporting history for prior-year analysis. October 2026 contains new ERP and subledger exports, with no completed October report.

Start with `FY2026/09 Sep/04 Close and review`. Preserve the September workbook and deck. Save new October files.

1. Roll the workbook to October 31. Freeze September's current-month formula facts into the historical snapshot before replacing the GL. Retain mappings, prior months, budgets, reforecasts, review notes and hidden legacy tabs.
2. Import October GL detail and trial balance. Preserve account codes as text so INDEX/MATCH finds the original reporting map. Every journal is double-entry; debit-positive TB and positive reporting revenue are intentional conventions.
3. The new account `6265 Third-party warranty callouts` is unmapped. Inspect the source invoice batch (`AP-WAR-2610-084`, $18,750), add it to Warranty / PL, then rebuild the current-month mapped activity formulas. Do not suppress the unmapped-account check.
4. Review payroll accrual true-up `JE-202610-017` for $24,600, already posted in the October ERP export. Do not post it a second time. Roll the payroll/warranty accrual, insurance prepaid and fleet depreciation schedules from September closing to October opening and movements.
5. Update T12 to November 2025–October 2026 and add October to YTD. Refresh department/entity P&L, balance sheet, cash flow, original budget vs actual, July reforecast, prior year, and KPI headcount from the payroll summary.
6. Write flux commentary grounded in ledger amounts and the documented company events. The June $42,000 warranty reserve restatement and $10,500 tax effect stay in historical results. Do not erase the superseded June pack or use its numbers for current comparisons.
7. All 14 Checks controls must PASS. Review the independent bank, AP, AR and payroll ties. Recalculate fully in Excel and scan every formula cell for errors.
8. Open September's six-slide deck in PowerPoint, save as October, and update the existing native chart data, financial table, date labels and commentary. Charts must remain editable, embedded and unlinked. Export all slides and inspect the result, including wrapping, chart categories, units and legends.
9. Download the final workbook, deck, slide images and app evidence. Run `validate_history.py --require-excel` and `validate_october.py`. Record the desktop from the source-folder review through the finished deck. Keep the unedited take and action timestamps; any shorter edit must document cuts/speed changes.

`recorded-steps/` contains the PowerShell COM stages used for this recording. They depend on the signed-in desktop and the `C:\work\Panes-monthly-close` task path, and are sequential stages rather than a background headless app. Run through Panes SDK `machine.exec(..., desktop=True)`, inspect before each stage, and never run an append/roll stage twice on the same workbook.

Expected October: revenue $1,988,001.06; EBITDA $90,142.13; net income $41,488.92; closing cash $3,426,168.98; ending headcount 141. These are reconciliation targets, not inputs to paste over formulas.

## Review gate

Publication approved November 15, 2026 after review of the replacement recording, sample files and page preview. Release both changes after their checks pass, then verify the live example and downloads.
