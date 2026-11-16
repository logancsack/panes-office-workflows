"""Generate registry metadata for the exact release bundle, without credentials."""

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def metadata():
    manifest = json.loads((ROOT / "manifest.json").read_text())
    version = manifest["version"]
    name = f"panes-{version}.mcpb"
    return {
        "$schema": "https://static.modelcontextprotocol.io/schemas/2025-12-11/server.schema.json",
        "name": "io.github.logancsack/panes",
        "title": "Panes",
        "description": "Windows computers for AI agents. Office, desktop apps, browser tasks, testing and file transfer.",
        "version": version,
        "websiteUrl": "https://usepanes.com/docs/mcp",
        "repository": {"url": "https://github.com/logancsack/panes-office-workflows", "source": "github"},
        "packages": [{
            "registryType": "mcpb",
            "identifier": f"https://github.com/logancsack/panes-office-workflows/releases/download/desktop-v{version}/{name}",
            "fileSha256": hashlib.sha256((ROOT / "dist" / name).read_bytes()).hexdigest(),
            "transport": {"type": "stdio"},
        }],
    }


if __name__ == "__main__":
    (ROOT / "dist" / "server.json").write_text(json.dumps(metadata(), indent=2) + "\n")
