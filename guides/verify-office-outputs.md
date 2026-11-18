# Download and verify your Excel and PowerPoint files

Get the workbook and editable deck off your Panes computer. Then reconcile the files against the source CSVs, rather than trusting the agent's completion message.

This guide is for the [compact expense workflow](../COMPACT_TASK.md), used by the [OpenAI Agents SDK](../integrations/openai-agents/README.md) and [LangChain](../integrations/langchain/README.md) examples. The [37-month finance close](../monthly-close/README.md) has separate inputs and a separate validator.

## Download the completed task

Run these commands from the repository root with Python 3.12+ and [uv](https://docs.astral.sh/uv/getting-started/installation/). The computer must already be running, with the compact task saved in `C:/work/Panes-office-demo`. Follow [setup](../prepare.py) and your framework's guide first. Native Office work requires real Microsoft desktop apps and a signed-in, activated Office session on Windows.

Prepare the pinned SDK and preview the download. Neither command starts a computer or calls a model:

```sh
python integrations/desktop/build.py --prepare-sdk
uv run --project integrations/desktop --locked python download.py --dry-run
```

Keep `PANES_API_KEY` in your secret environment, as in the framework setup. Use the same computer ID as your task:

```sh
uv run --project integrations/desktop --locked python download.py --machine YOUR_MACHINE_ID --execute
```

The helper reads only the selected computer's state and these two files:

- `C:/work/Panes-office-demo/October close.xlsx`
- `C:/work/Panes-office-demo/October review.pptx`

It saves unchanged bytes in a fresh `local-results/compact-…` directory and prints that directory. A private `download.json` records file sizes and SHA-256 hashes. The helper caps each download at 64 MiB, restricts local access to your OS user, and never overwrites an earlier download. `local-results/` is ignored by Git; keep customer work private.

If either read fails, the helper exits unsuccessfully. Inspect the printed directory for a partial download before retrying. A download receipt means both transfers finished; it does not mean the finance work is correct.

## Validate your files

Replace the directory below with the exact path printed by the helper:

```sh
uv run --with-requirements requirements.txt python validate.py --outputs-dir local-results/compact-YOUR_DIRECTORY
```

`--outputs-dir` matters: running `python validate.py` alone checks this repository's bundled recording outputs. With the flag, the validator uses only your selected workbook and deck. Missing or invalid files fail; it never substitutes the bundled outputs.

A successful check writes `validation.json` alongside your downloaded files. It verifies the recorded compact layout against the supplied CSVs, including:

- 483 source rows, 480 unique documents and three duplicates removed.
- $686,500 actual, $645,000 budget and $41,500 unfavorable for October.
- 60 formula outputs and their saved results, one PivotTable and two Excel charts.
- Two editable PowerPoint charts, their embedded workbooks and no external deck links.

The report includes hashes of the files checked. A failed recheck replaces an earlier `PASS` with `FAIL` and returns a nonzero exit status. The validator never saves or reserializes the Office files. Run Python without `-O` or `PYTHONOPTIMIZE`; the validator refuses modes that disable its checks.

This is a strict check for the recorded compact layout. A differently arranged report needs corresponding checks. The repository's `reference/office-evidence.json` describes the original recording; it is not evidence of a new run. File checks cannot establish that a new Power Query refresh succeeded or that Office is activated. Inspect the actual Windows desktop, query results and fresh run evidence too.

## Stop the computer

Download and inspect the files before stopping the computer in Panes. This helper and the framework clients leave it running, including after a failure. Running time consumes prepaid credit; stopped computers retain storage charges until deleted. Keep the computer if you need to reopen the task.

[Watch the original compact Office recording](../media/office-report.mp4), [inspect the recorded workbook](../outputs/October%20close.xlsx), or [open the editable deck](../outputs/October%20review.pptx). The new download and validation path is tested locally with those existing native artifacts and a local SDK fixture; it does not claim a new model-driven Office run.
