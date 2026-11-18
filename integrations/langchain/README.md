# Run a LangChain agent in Excel and PowerPoint

Give a LangChain agent the Windows desktop where the work happens: native Excel Power Query, a PivotTable, charts, and a PowerPoint deck someone can keep editing. This example connects LangChain to the Panes MCP server and supplies the [original compact expense task](../../COMPACT_TASK.md).

[Watch the actual Windows run](../../media/office-report.mp4) · [Inspect the Excel workbook](../../outputs/October%20close.xlsx) · [Inspect the editable deck](../../outputs/October%20review.pptx)

The recording and Office files come from the November 11 Panes run. The data is fictional. This LangChain client was added November 15; its connection, preflight and tool restrictions are tested locally. A new model-driven Office run has not been performed with this client. The expanded [37-month close](../../monthly-close/README.md) is a separate task with its own files; the helper and default instructions here prepare the compact example.

## Check the connection without an account

Install [uv](https://docs.astral.sh/uv/getting-started/installation/), clone this repository, and run from `integrations/langchain`:

```sh
python ../desktop/build.py --prepare-sdk
uv run --locked office_agent.py
```

This downloads the hash-checked Panes SDK wheel from usepanes.com, then discovers its 18 MCP tools using the current `langchain.mcp.MCPAdapter`. The client and server have separate locked environments. The command uses an inert key and a loopback API address even if real credentials are in your shell. It makes no cloud or model calls.

The agent receives seven work tools: computer state, commands, file reads/writes, screenshots, the accessibility tree, and desktop input. It receives no create, resume, delete, clone, snapshot, stop or viewer tool. Middleware rejects a call targeting any computer other than the one selected by the operator, and rejects Office commands outside the signed-in desktop.

## Run the finance task on your computer

1. [Create a Panes account](https://usepanes.com/login?signup), fund its prepaid balance, and save its API key in your secret environment as `PANES_API_KEY`. Use an existing running computer with Excel and PowerPoint desktop installed and signed in. List current templates in Panes before choosing one; the Office-installed template is `windows-16gb`. Use a computer reserved for this example, with no unrelated work open.
2. Upload the compact example's CSVs and reference stages using the repository's [setup helper](../../prepare.py). From this directory, the command below uses the locked SDK environment. It refuses an existing task directory and saves the private viewer locally without printing its URL. To inspect its options without a key or network call, run `python ../../prepare.py --help`; use `python ../../prepare.py --dry-run` to preview the uploads.

```sh
uv run --project ../desktop --locked python ../../prepare.py --machine YOUR_MACHINE_ID --execute
```

3. Save your model provider key as `OPENAI_API_KEY`. Select an OpenAI vision model available to your account; the locked example includes LangChain's OpenAI provider. Then run:

```sh
uv run --locked office_agent.py --machine YOUR_MACHINE_ID --model openai:YOUR_VISION_MODEL --execute
```

`--execute` permits model calls and work on that running computer. The client first reads its state and refuses a stopped computer or an unresolved recovery warning. It passes [COMPACT_TASK.md](../../COMPACT_TASK.md) to the agent, which works in real desktop Office. The prepared task directory is `C:/work/Panes-office-demo`.

Open the private viewer from `local-results/viewer.json` to inspect the work. The agent should leave `October close.xlsx` and `October review.pptx` in that task directory. Reconcile the CSVs against the workbook and deck: $686,500 actual, $645,000 budget, $41,500 unfavorable, 480 unique documents, and editable embedded PowerPoint charts. The [independent verifier](../../validate.py) documents the expected checks for the recorded output layout; an agent's completion message alone is not verification.

[Download and verify your own files](../../guides/verify-office-outputs.md), then stop and keep the computer in Panes. This client leaves the selected computer running after completion or failure; it never changes its lifecycle. Running time uses prepaid Panes credit, and the model provider bills model calls separately. Set a provider spending limit before running.

## Bound the run and inspect failures

The default limits are 40 graph steps and 900 seconds. Set `--max-steps` and `--max-seconds` explicitly for another task; `--task-file` selects its instructions. These limits end the local agent, not a command already running on Windows or the computer's billing. Inspect the desktop and existing outputs before retrying an interrupted run.

Commands retain the power of the signed-in Windows user; the tool filter is a workflow restriction, not an operating-system security boundary. Keep customer documents and credentials on the private computer. The example disables hosted LangSmith tracing and prints a completion count rather than the document/tool transcript. No model/provider error body or private viewer link is printed by this client.

## Verify the integration

```sh
uv run --locked python -m unittest test_connection.py
```

The tests start the actual Panes stdio server and use LangChain 1.5.0 / FastMCP 4.2.0. A loopback HTTP fixture supplies computer state; deterministic test messages exercise LangChain's tool middleware without a paid model. The checks prove tool discovery, structured state parsing, refusal of stopped/recovery-warning computers, wrong-computer rejection, and signed-in desktop enforcement. They do not simulate or substitute an Office demo.

Maintained by Aldo, Logan's AI assistant, for **Renaissance Superintelligence Corporation (RSI)**, which operates **Panes**. No LangChain endorsement or partnership is implied.

Primary documentation checked November 15, 2026: [LangChain MCP](https://docs.langchain.com/oss/python/langchain/mcp), [migration from the retired adapter package](https://docs.langchain.com/oss/python/migrate/langchain-mcp-adapters), and [LangChain tool middleware](https://docs.langchain.com/oss/python/langchain/middleware/custom#wrap-style-hooks).
