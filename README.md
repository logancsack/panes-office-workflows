# Panes

**Windows computers for AI agents.**

Panes gives agents persistent Windows computers to run desktop apps, execute commands, move files and work through the screen. Connect through MCP, Python, TypeScript or the API. Watch the desktop and take over from your browser.

[Connect your agent](https://usepanes.com/docs/mcp) · [Quickstart](https://usepanes.com/docs/quickstart) · [API and SDKs](https://usepanes.com/docs) · [Pricing](https://usepanes.com/#pricing)

## What your agents can do

- **Office work:** refresh Excel models, revise Word documents and build editable PowerPoint presentations in the native desktop apps.
- **Desktop apps:** operate Windows software through commands, accessibility, screenshots, mouse and keyboard.
- **Browser tasks:** use Microsoft Edge to work in web apps and move results into desktop tools.
- **Testing:** run Windows programs, exercise desktop interfaces and inspect screenshots and output files.
- **Enterprise workflows:** keep apps, files and sign-ins between tasks; move work across business software on the same computer.

## Connect

Use the hosted MCP server with a compatible Streamable HTTP client:

```text
https://api.usepanes.com/mcp/agent
```

1. Follow the [MCP setup](https://usepanes.com/docs/mcp) to create an account and connect with a Panes API key. Discovery and email-code signup are available through the endpoint.
2. Ask your agent to list templates and current prices. Choose the apps and computer size for your task. For Office work, select an Office-ready template and complete your Office sign-in in the private desktop viewer.
3. Fund the account, start a computer, complete the task, download the results and stop the computer. Compute is prepaid and billed by the second.

For local stdio clients, use the [manual connection](integrations/mcp/README.md) or [desktop bundle](integrations/desktop/README.md). Build with the [Python or TypeScript SDK](https://usepanes.com/docs/sdks), [LangChain](integrations/langchain/README.md) or [OpenAI Agents SDK](integrations/openai-agents/README.md).

## Computer controls

| Capability | Tools in the local MCP connection |
|---|---|
| Discover computers and capacity | `account`, `templates`, `machines`, `machine_get` |
| Create and recover a create request | `machine_create`, `machine_request` |
| Start, stop and delete | `machine_start`, `machine_stop`, `machine_delete` |
| Run commands and transfer files | `exec`, `file_read`, `file_write` |
| Observe and control the desktop | `screenshot`, `accessibility`, `input`, `viewer` |
| Save and copy environments | `snapshot`, `fork` |

The [hosted MCP catalog](https://usepanes.com/docs/mcp) also covers account signup and funding. Use the private viewer for sign-in and human handoff. Keep API keys and viewer links in your secret store.

## See real work

[Watch the Office workflow](https://usepanes.com/examples/excel-to-powerpoint): an agent carries a monthly close through real Excel and PowerPoint on Panes, returning an editable workbook and management deck. The company and source data are fictional; the computer, apps and recorded work are real. [Inspect the task, files and validation](monthly-close/README.md).

This repository contains the runnable connections and example assets. The [compact task](COMPACT_TASK.md) matches `prepare.py` and the framework clients; `python prepare.py --help` documents uploads and private viewer recovery. Example verification commands and recorded native stages are documented in the [workflow guide](monthly-close/README.md).

Panes is a product of **Renaissance Superintelligence Corporation (RSI)**. [Documentation](https://usepanes.com/docs) · [Support](https://usepanes.com/support) · [Connection issues](https://github.com/logancsack/panes-office-workflows/issues).
