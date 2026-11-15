"""Read-only preflight for the explicitly selected computer."""

import json


async def require_running(tools, machine: str) -> None:
    tool = next(tool for tool in tools if tool.name == "machine_get")
    # Request a ToolMessage so structured MCP artifacts are preserved.
    message = await tool.ainvoke({"type": "tool_call", "name": "machine_get", "id": "preflight", "args": {"id": machine}})
    if message.status == "error":
        raise ValueError("Could not read the selected computer.")
    artifact = message.artifact
    state = artifact.get("structured_content") if isinstance(artifact, dict) else None
    if state is None:
        blocks = message.content
        if isinstance(blocks, str):
            state = json.loads(blocks)
        else:
            state = json.loads(next(block["text"] for block in blocks if block.get("type") == "text"))
    if state.get("id") != machine or state.get("state") != "running":
        raise ValueError("Select an existing running computer; this example never starts one.")
    if state.get("stop_warning") or state.get("resume_warning"):
        raise ValueError("Resolve the computer's recovery warning before running the workflow.")
