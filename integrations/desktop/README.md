# Panes desktop connection

Load Panes into a desktop agent that supports **MCPB 0.4 with the UV runtime**. The host manages Python and dependencies. Your agent gets the same 18 tools as the [manual MCP connection](../mcp/README.md).

The bundle uses the unchanged public Panes SDK 0.1.3, included as a hash-checked wheel. It does not install a similarly named PyPI package. The locked MCP dependencies download on first setup; the host needs network access to package registries.

## Install

1. Download `panes-0.1.0.mcpb` from this repository's `desktop-v0.1.0` release once available. Check its SHA-256 against the release's checksum file.
2. Open the bundle in a compatible desktop client or use its extension installer. The client shows the publisher and collects your **Panes API key** as a sensitive setting. Create the key in your [Panes dashboard](https://usepanes.com/dashboard).
3. Ask the agent to list your account and templates first. Then give it [the Excel-to-PowerPoint task](../../TASK.md), the inputs and an explicit execution budget.

Installing the connection does not create or start a cloud computer. Running computers spend Panes prepaid credit; check current prices before creating one. Office work uses the Office-installed `windows-16gb` template and your own Office sign-in. Stop the computer when finished. Keep API keys and private desktop viewer links out of posts and public files.

This is a Panes connection maintained by Aldo for **Renaissance Superintelligence Corporation**. It is not an endorsement or partnership with a client vendor. Local dependency resolution, the stdio handshake and tool discovery are checked; native desktop installation and a new cloud Office run have not been performed for this bundle.

## Build and inspect

From this directory:

```sh
python build.py --prepare-sdk
uv sync --locked
npx --yes @anthropic-ai/mcpb@2.1.2 validate manifest.json
python build.py
uv run --locked smoke_test.py
python registry.py
```

`build.py` includes only its explicit allowlist. It never copies `.env`, a virtual environment, local credentials or results into the ZIP. It refuses a changed SDK hash. `uv.lock` pins the dependencies; the bundled wheel supplies `panes` directly.

For a local protocol check, use an inert API key and a loopback API base. No cloud API or billable tool calls are needed to initialize the server and list its tools.

The manual `Publish Panes desktop connection` workflow runs only from `main`. It checks the bundle again, publishes versioned release assets, and authenticates to the official MCP Registry using GitHub's short-lived OIDC identity. No new service account, stored registry token or paid service is needed. Re-running refuses to replace a release bundle with different bytes. Registry publication is verified separately after the workflow; a successful release alone does not prove a listing exists.

Specification checked November 14, 2026: [MCPB manifest and UV runtime](https://github.com/modelcontextprotocol/mcpb/blob/main/MANIFEST.md), [official UV example](https://github.com/modelcontextprotocol/mcpb/tree/main/examples/hello-world-uv). The manifest reference's latest change was March 12, 2026; this is current format documentation, not a reported marketing result.
