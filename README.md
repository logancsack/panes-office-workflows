# A real monthly close in Excel and PowerPoint

A three-year finance shared drive. An existing close model. A management deck that needs to go out again next month. Watch an agent carry the October 2026 close through real Microsoft Excel and PowerPoint on a [Panes](https://usepanes.com) Windows computer.

[Explore the monthly-close workflow](monthly-close/README.md) · [Agent task](monthly-close/TASK.md) · [October workbook](monthly-close/outputs/October%202026%20close.xlsx) · [Management deck](monthly-close/outputs/October%202026%20review.pptx)

Alder Mechanical Services Group is a fictional sample company. Its history contains **37 monthly packs**, September 2023–September 2026, and new October source exports. Each pack includes ERP GL detail and trial balance, AP/AR aging, payroll, bank recs, an Excel close workbook and a six-slide management deck. The group has multiple entities, an acquisition, intercompany eliminations, a department reorganization, a COA change, budgets, reforecasts, and a restated month.

The recorded October task rolls September forward, imports the new ledger, maps an unmapped warranty account, verifies a payroll accrual true-up, refreshes the financial statements and analysis, writes flux commentary, and updates the editable management deck. The raw footage, slide exports and independent validators are included. The file history is synthetic; the Panes computer, Office apps and recorded work are real.

## Verify the files

```sh
python -m pip install -r requirements.txt
python monthly-close/validate_history.py --require-excel
python monthly-close/validate_october.py
```

The validators read the real Office files without rewriting them. They reconstruct expected values from delivered CSVs and inspect the cached results saved by desktop Excel and the native PowerPoint chart data. [Native Office evidence](monthly-close/desktop-validation.json) records the historical file-opening and calculation checks. [CSV evidence](monthly-close/csv-open-validation.json) records source files opened in Excel.

## Run it on Panes

Follow the [MCP setup](integrations/mcp/README.md) or use the [Claude Desktop extension](integrations/desktop/README.md), list the available templates and choose the Office-installed computer. Use your kept signed-in Office computer and its existing files, or complete your own Office sign-in. Give the agent the [October close task](monthly-close/TASK.md). Run Office work in the signed-in desktop (`desktop=True`) and use the private viewer to watch or take over. Keep keys and customer files in your own secret store.

Building with LangChain? The [runnable LangChain connection](integrations/langchain/README.md) checks MCP tool discovery locally without an account or model call. Its execution example uses the original compact expense task below.

Using the OpenAI Agents SDK? The [runnable OpenAI connection](integrations/openai-agents/README.md) checks the real MCP transport and restricts the agent to one selected Windows computer. Discovery needs no account or model call; its execution example uses the original compact expense task.

The [recorded native stages](monthly-close/recorded-steps) document the COM work performed in real Excel and PowerPoint. They are stateful steps with file and desktop preconditions, not a substitute UI or a headless rendering of Office.

## Original compact example

The [compact agent task](COMPACT_TASK.md) matches the root upload helper and both framework clients. The original expense-export demo remains in `inputs/`, `outputs/`, `reference/`, `media/`, `prepare.py` and `validate.py`. It demonstrates native Power Query cleanup, SUMIFS, a PivotTable and two embedded PowerPoint charts. Its original validator still runs with `python validate.py`; the upload helper is documented by `python prepare.py --help`. The multi-year monthly close is the expanded example above. The original upload helper saves private handoffs in `local-results/viewer.json`, restricted to your local user on Windows and POSIX. Use `python prepare.py --machine YOUR_MACHINE_ID --execute --viewer-only` to recover a viewer link after completed uploads, without uploading again.

Panes gives an agent the Windows desktop, apps and files. Bring your agent through MCP or the SDKs. [Read the docs](https://usepanes.com/docs#mcp) or [open a workflow question](https://github.com/logancsack/panes-office-workflows/issues/new).
