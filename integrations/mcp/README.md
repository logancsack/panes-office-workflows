# Connect an MCP-capable agent to Panes

Python 3.11 or later. Panes SDK 0.1.3 is distributed directly; do not install another project's similarly named PyPI package.

```sh
python -m venv .venv
.venv/bin/python -m pip install 'panes[mcp] @ https://usepanes.com/sdk/panes-0.1.3-py3-none-any.whl'
```

On Windows use `.venv\Scripts\python.exe`. Add this server to your MCP client's configuration, replacing the executable path with your own absolute path:

```json
{
  "mcpServers": {
    "panes": {
      "command": "/ABSOLUTE/PATH/.venv/bin/panes-mcp",
      "args": []
    }
  }
}
```

Set PANES_API_KEY using the client's secret/environment settings. Do not put it in the JSON you share, in source control or in the agent task. On Windows the executable ends `.venv\Scripts\panes-mcp.exe`.

The initial read-only checks are account and templates. They reveal balance, current sizes and prices before a billable create. Then give the agent TASK.md and your explicit execution budget. Use a unique request_key when creating a computer; inspect an uncertain request rather than creating twice. Desktop Office calls use exec with desktop=true. Screenshots, keyboard/mouse, accessibility, file transfer and the private viewer are available to the agent.

[Current Panes MCP docs](https://usepanes.com/docs#mcp). This example uses the MCP connection; it does not claim endorsement or an official partnership with any agent framework.
