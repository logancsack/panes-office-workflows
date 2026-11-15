# Delivery verification

**PASS** — Microsoft Excel and PowerPoint desktop 16.0 on the kept Panes Windows computer.

- 37 historical closes: September 2023–September 2026.
- 39 workbooks and 39 presentations opened in native Office, including the superseded June pack and October outputs. Their exact delivered SHA-256 hashes match the read-only native verification.
- 241 CSV exports opened in Excel; 319 Office/source files in total.
- Every workbook has 14 passing financial controls; no formula errors.
- Independent October reconciliation checks 543 cells across monthly, prior-year, entity, department, T12, YTD and chart results.
- Six-slide October deck; three editable embedded charts; no external links.
- Uncut 71.2-second real desktop recording; no synthetic UI, speed changes or cuts.
- Sample-company author metadata; no personal name in the delivered Office files.

October revenue: $1,988,001.06. EBITDA: $90,142.13. Net income: $41,488.92. Cash: $3,426,168.98. Ending FTE: 141.

The validator derives expected figures from the delivered ledger exports, not the generator's expected-results file. Annual retained earnings, intercompany eliminations, bank reconciliations, AP/AR aging, payroll, opening-to-closing continuity and GL-to-TB controls pass across all periods. The August 2023 opening trial balance anchors the first monthly close.

Run `python monthly-close/validate_delivery.py` from the repository root. See `verified-deliverables.json` for exact native-checked hashes, `delivery-validation.json` for the summary and the individual validation files for detail.
