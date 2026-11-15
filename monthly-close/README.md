# A three-year finance close, continued in Excel and PowerPoint

An agent takes September's living close workbook and management deck through the October 2026 close on a real Panes Windows computer. The sample company, Alder Mechanical Services Group, is fictional. The files and figures are synthetic; the Microsoft desktop apps and recorded work are real.

The shared drive contains **37 monthly close packs** from September 2023 through September 2026. Each month has GL detail, a trial balance, AP/AR aging, payroll, bank reconciliations, an Excel close workbook and a six-slide management deck. Supporting history begins in 2022. Calendar fiscal years have original board budgets and July reforecasts.

The group includes an operating business, shared services and a January 2025 acquisition, with explicit intercompany elimination journals. Its history includes spring/summer staffing, heating/cooling seasonality, annual price and wage changes, obsolete inventory, department changes, a payroll COA migration, retained legacy tabs, reviewed versions, a revised June warranty reserve and a February 2026 deck refresh.

The living workbook retains native formulas, structured tables and named ranges. The current month's analysis flows from GL debit/credit rows through the mapping table. Prior closed periods remain frozen snapshots. Schedules roll opening balances, additions, provisions, amortization, depreciation, settlements and ending balances. The budget, prior year, YTD, T12, entity/department P&L, balance sheet and indirect cash flow share the same ledger.

## Inspect and reproduce

- [Agent task](TASK.md)
- [Monthly pack index](manifest.json)
- [Source validation](validation.json)
- [October output validation](october-validation.json)
- [Native desktop file-opening and calculation evidence](desktop-validation.json)
- [Completed October workbook](outputs/October%202026%20close.xlsx) and [deck](outputs/October%202026%20review.pptx)
- [Raw recording](media/october-close-unedited.mp4) and [edit provenance](media/provenance.json)

```sh
python -m pip install -r requirements.txt
python monthly-close/validate_delivery.py
```

Run those commands from the repository root. The combined validator also matches all 78 delivered Office files to hashes checked in real Office, after removing personal author metadata. The validators read cached results saved by real desktop Excel, reconstruct expected figures from delivered exports using decimal arithmetic, and inspect native PowerPoint chart caches and embedded data. They never substitute another calculation engine or rewrite the Office files.

`build_history.py` deterministically constructs the fictional past using balanced integer-cent double-entry journals, then rolls each saved workbook into the next. It produces source workbooks; native Excel recalculation and PowerPoint opening are separate required steps. Rebuilding replaces local fixtures, so use a scratch checkout. `desktop_validate.ps1` and `desktop_finish_history.ps1` document native Office validation. The recorded October stages use only September and the October source files, not a prepared October workbook.

This is a constructed history of a finance process. File dates, reviewed initials and company events are part of the fictional case; they are not a claim that a real company maintained these documents in those years.
