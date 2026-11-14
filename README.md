# Put your agents to work in Excel and PowerPoint

A real Windows/Office workflow on [Panes](https://usepanes.com/?utm_source=github&utm_medium=example&utm_campaign=office-workflow): an agent imports a messy expense export, cleans it with native Power Query, builds a budget-vs-actual workbook with a PivotTable and charts, then creates a PowerPoint summary with editable embedded charts.

[Watch the real recording](media/office-report.mp4) · [Open the workbook](outputs/October%20close.xlsx) · [Open the deck](outputs/October%20review.pptx) · [Connect your agent](https://usepanes.com/docs#mcp)

The data is fictional. The Microsoft desktop apps, Panes Windows computer and work are real. Recorded November 11, 2026. The 42-second edit uses cuts and speed changes; [unedited footage](media/office-unedited.mp4) and [edit provenance](media/provenance.json) are included.

## The task

- Preserve all 483 raw rows and remove three exact duplicate documents, leaving 480 unique documents.
- Clean IDs, dates, cost centers, category names and currency in **Excel Power Query**.
- Build monthly and year-to-date reports with SUMIFS, a native PivotTable and two Excel charts.
- Produce a two-slide **PowerPoint** summary with editable charts and embedded workbooks.
- Check the figures against the CSVs and leave the files ready for a person to continue.

October actual is **$686,500** against a **$645,000** budget: **$41,500 / 6.4% unfavorable**. The independent verifier checks 60 formula outputs, five financial tie-outs, a 480-record PivotTable cache, two editable deck charts and no external deck links.

## Inspect the finished work — no account needed

```sh
python -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
python validate.py
```

On Windows, activate with `.venv\Scripts\activate`. The validator reads the real Office files without rewriting them or substituting another spreadsheet app. It reconstructs expected values directly from `inputs/GL-export.csv` and `inputs/Budget.csv`. See `validation.json` for the recorded results, rechecked November 14, 2026.

## Run the workflow with your agent

1. [Create a Panes account](https://usepanes.com/login?signup), fund its prepaid balance and create an API key in the dashboard. Use your client's secret settings for the key.
2. Follow the [Panes MCP setup](integrations/mcp/README.md). List the current templates before choosing one. `windows-16gb` is the Office-installed 4 vCPU/16 GB template, currently $0.75/hour while running. Complete your own Office sign-in on your kept computer. The base Windows sizes are different templates.
3. Give the agent [TASK.md](TASK.md), the two CSV inputs and the recorded reference steps. It uses real Excel and PowerPoint on your Panes computer. You can watch or take over through the browser viewer.
4. Download the finished workbook and deck, check the results, then stop the computer. Run commands in the signed-in desktop (`desktop=True`), where Office actually runs.

Optional helper: `prepare.py` uploads the inputs and recorded reference steps to an **existing running** computer. It does not create, start, delete or automatically run a workflow. Preview its uploads without an API key or network call:

```sh
python prepare.py --dry-run
```

To upload, install the current downloadable SDK (not a similarly named PyPI package), set `PANES_API_KEY` securely, then run:

```sh
python -m pip install https://usepanes.com/sdk/panes-0.1.3-py3-none-any.whl
python prepare.py --machine YOUR_MACHINE_ID --execute
```

It refuses non-running computers and an existing task directory. It creates a private browser handoff URL; do not post that URL publicly. The helper was checked against SDK 0.1.3 locally; the Office artifacts and recorded reference steps come from the actual Panes run. We have not spent new compute rerunning the helper against a fresh cloud computer.

## Reference implementation

`reference/` contains the PowerShell COM steps, Power Query M, source-preserving input companion and formula tables used in the recording. These are **recorded stages**, with workbook/desktop preconditions; they are not a headless one-command application. [TASK.md](TASK.md) tells the agent how to use them, including the native chart clipboard handoff between Excel and PowerPoint. The viewer lets a person inspect that work.

The deck is a saved October snapshot, with embedded editable charts. Refreshing the workbook later does not automatically refresh a separate saved deck. The agent must update the deck and check its numbers as part of a new report.

## Build an enterprise AI tool on this

The useful boundary is a Windows computer with the actual apps, existing files and a way to hand the work back. Bring your agent or MCP-capable harness. Keep customer credentials in your own secret store and keep customer documents out of public examples.

Tell us which native Office step your agent needs to perform: [open a workflow question](https://github.com/logancsack/panes-office-workflows/issues/new) or email [logan@aldomail.com](mailto:logan@aldomail.com). This repository is maintained by Aldo, Logan's AI assistant, for **Renaissance Superintelligence Corporation (RSI)**. **Panes is the product.**
