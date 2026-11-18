# Put OpenAI agents to work in Excel and PowerPoint

Connect the OpenAI Agents SDK to a Panes Windows computer. Give the agent native Excel, Power Query, PivotTables, charts and PowerPoint desktop apps, then keep the editable workbook and deck.

This client uses the [original compact expense workflow](../../COMPACT_TASK.md). [Watch its real Panes Office recording](../../media/office-report.mp4), [open the workbook](../../outputs/October%20close.xlsx) and [inspect the editable deck](../../outputs/October%20review.pptx). The expanded [37-month finance close](../../monthly-close/README.md) is a separate example. The recordings show real Office; this new SDK client's transport and restrictions are tested locally, without a new model-driven Office run.

## Check the connection first

Install [uv](https://docs.astral.sh/uv/getting-started/installation/), clone this repository and run from `integrations/openai-agents`:

```sh
python ../desktop/build.py --prepare-sdk
uv run --locked office_agent.py
```

The first command downloads the hash-checked Panes SDK wheel from usepanes.com. The second starts its real stdio MCP server, verifies all 18 tool names and exposes seven work tools to the OpenAI agent: computer status, commands, file reads/writes, screenshots, accessibility and desktop input. Discovery uses an inert key and loopback API address, even if real keys are present in your environment. It makes no Panes cloud or model calls.

The OpenAI client and Panes server have separate locked Python environments. You do not need a Panes account or an OpenAI API key for discovery.

## Run the Office task

1. Use your [Panes account](https://usepanes.com/login) and existing prepaid balance. Save its API key as `PANES_API_KEY` through your secret environment. Select an existing running Windows computer with Excel and PowerPoint desktop installed and signed in. Use a computer reserved for this example, with no unrelated work open. Inspect current templates rather than treating a template name as proof of Office activation.
2. Upload the compact workflow inputs and native reference stages. From this directory:

```sh
uv run --project ../desktop --locked python ../../prepare.py --machine YOUR_MACHINE_ID --execute
```

The helper refuses an existing task directory and saves the private viewer in `local-results/viewer.json`, restricted to your local user. It does not print the viewer URL. Preview uploads with `python ../../prepare.py --dry-run`.

3. Set `OPENAI_API_KEY` in your secret environment and a spending limit in your OpenAI project. Select a vision model available to your API account. Then explicitly execute:

```sh
uv run --locked office_agent.py --machine YOUR_MACHINE_ID --model YOUR_VISION_MODEL --execute
```

The client reads the selected computer's actual state before making a model call. It refuses a stopped computer or unresolved recovery warning. The agent receives the compact [COMPACT_TASK.md](../../COMPACT_TASK.md), works in `C:/work/Panes-office-demo`, and runs Office commands in the signed-in Windows desktop.

Use the saved private viewer to inspect the work. Verify $686,500 actual, $645,000 budget, $41,500 unfavorable, 480 unique documents and editable embedded PowerPoint charts. The [independent validator](../../validate.py) documents the recorded output layout and checks; a model's completion message is not verification. [Download and verify your own files](../../guides/verify-office-outputs.md) with the helper and `--outputs-dir` command.

**Stop the computer in Panes when finished.** This client leaves it running after success, interruption or failure. It cannot create, start, stop, delete, clone or snapshot computers. Panes runtime uses prepaid balance; OpenAI bills model calls separately. Stopped storage continues until you delete the computer in Panes.

## Scope and interrupted work

Every call is checked before MCP transport: it must use a work tool and the exact computer ID selected on the command line. An Office command with `desktop=false` is rejected. Lifecycle, account, viewer and snapshot tools are absent from the agent's catalog and rejected if called directly. MCP retries are disabled, and the model is asked to use serial tool calls.

These checks restrict this client's computer/tool selection. Commands still have the signed-in Windows user's permissions, including network access; they are not an operating-system security boundary. Run fictional inputs first and keep unrelated credentials and documents off the example computer.

The defaults are 40 model turns and 900 seconds. `--max-turns`, `--max-seconds` and `--task-file` select another bound or task. A local timeout or Ctrl+C ends the client, not a command already running on Windows or the computer's billing. Inspect current files and desktop state before retrying. If Office requests sign-in, resolve it in the private desktop before continuing.

Hosted Agents SDK tracing is disabled. The CLI prints a completion count rather than the document/tool transcript, provider error body or private viewer URL. The model provider still receives the task and tool results needed to run it.

## Verify the integration

```sh
uv run --locked python -m unittest test_connection.py
```

The tests use OpenAI Agents SDK 0.24.0, the actual Panes MCP subprocess and a loopback HTTP fixture. Scripted model messages exercise the real Agents SDK Runner without hosted inference. They verify discovery, selection checks, rejected wrong-computer/non-desktop/lifecycle calls, and stopped/recovery-warning preflight. Fixture results are transport checks, not Office execution evidence.

This standalone API-key example is separate from the OAuth connection at [usepanes.com/openai](https://usepanes.com/openai). No ChatGPT directory approval, OpenAI endorsement or partnership is claimed.

Maintained by Aldo, Logan's AI assistant, for **Renaissance Superintelligence Corporation (RSI)**, which operates **Panes**.

Primary documentation checked November 15, 2026: [OpenAI Agents SDK MCP](https://openai.github.io/openai-agents-python/mcp/), [MCP server API](https://openai.github.io/openai-agents-python/ref/mcp/server/) and [tracing](https://openai.github.io/openai-agents-python/tracing/).
