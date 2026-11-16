"""Inspect a relocated bundle and discover its tools without cloud API calls."""

import asyncio
import hashlib
import json
import os
from pathlib import Path
import tempfile
import zipfile

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

from build import SDK_NAME, SDK_SHA256, ROOT


async def check():
    version = json.loads((ROOT / "manifest.json").read_text())["version"]
    artifact = ROOT / "dist" / f"panes-{version}.mcpb"
    expected = {"manifest.json", "pyproject.toml", "uv.lock", "src/server.py", "README.md", "LICENSE", "wheels/" + SDK_NAME}
    with tempfile.TemporaryDirectory(prefix="panes-bundle-check-") as directory:
        with zipfile.ZipFile(artifact) as bundle:
            assert set(bundle.namelist()) == expected, "Unexpected or missing bundled files"
            assert hashlib.sha256(bundle.read("wheels/" + SDK_NAME)).hexdigest() == SDK_SHA256
            manifest = json.loads(bundle.read("manifest.json"))
            assert manifest["user_config"]["api_key"]["sensitive"] is True
            assert manifest["user_config"]["api_key"]["required"] is True
            bundle.extractall(directory)
        server = StdioServerParameters(
            command="uv",
            args=["run", "--locked", "--directory", directory, "src/server.py"],
            env={
                "PATH": os.environ["PATH"],
                "PANES_API_KEY": "local-protocol-check-not-a-real-key",
                "PANES_BASE_URL": "http://127.0.0.1:9",
            },
        )
        async with stdio_client(server) as (read, write):
            async with ClientSession(read, write) as session:
                initialized = await session.initialize()
                assert initialized.server_info.name == "Panes"
                result = await session.list_tools()
                names = {tool.name for tool in result.tools}
                assert len(names) == 18, names
                assert {"account", "templates", "exec", "screenshot", "input", "viewer", "machine_stop"} <= names
                print(json.dumps({"bundle": artifact.name, "server": initialized.server_info.name, "tool_count": len(names), "tools": sorted(names), "cloud_tools_called": 0}))


if __name__ == "__main__":
    asyncio.run(asyncio.wait_for(check(), timeout=90))
