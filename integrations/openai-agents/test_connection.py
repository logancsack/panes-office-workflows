"""Actual SDK/MCP transport with a loopback fixture and scripted model; no Office run."""

import json
import importlib.util
import os
import re
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

from agents import Model, ModelResponse, RunConfig, Runner, Usage
from openai.types.responses import ResponseFunctionToolCall, ResponseOutputMessage, ResponseOutputText

from office_agent import ROOT, SelectedComputerServer, SetupError, WORK_TOOLS, office_agent, parse_args, require_running, run


class ScriptedModel(Model):
    """Deterministic messages that exercise the real Runner; no hosted inference."""

    def __init__(self):
        self.calls = 0

    async def get_response(self, *args, **kwargs):
        assert {tool.name for tool in kwargs["tools"]} == WORK_TOOLS
        assert kwargs["model_settings"].parallel_tool_calls is False
        self.calls += 1
        if self.calls == 1:
            cases = [
                ("wrong-computer", "machine_get", {"id": "m-other"}),
                ("selected-computer", "machine_get", {"id": "m-selected"}),
                ("wrong-desktop", "exec", {"id": "m-selected", "command": "Get-Date", "desktop": False}),
            ]
            output = [ResponseFunctionToolCall(type="function_call", call_id=call_id, name=name, arguments=json.dumps(arguments)) for call_id, name, arguments in cases]
        else:
            output = [ResponseOutputMessage(
                id="fixture-message", type="message", role="assistant", status="completed",
                content=[ResponseOutputText(type="output_text", text="Protocol test complete", annotations=[])],
            )]
        return ModelResponse(output=output, usage=Usage(), response_id=None)

    def stream_response(self, *args, **kwargs):
        raise NotImplementedError("This fixture is not a streaming or hosted model.")


class ConnectionTest(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.requests = []
        self.state = "running"
        self.warning = ""
        owner = self

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                owner.requests.append(("GET", self.path))
                if self.path != "/v1/machines/m-selected":
                    self.send_error(404)
                    return
                data = json.dumps({"id": "m-selected", "state": owner.state, "stop_warning": owner.warning}).encode()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def do_POST(self):
                owner.requests.append(("POST", self.path))
                self.send_error(405)

            def log_message(self, *args):
                pass

        self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.env = patch.dict(os.environ, {
            "PANES_API_KEY": "local-protocol-test-inert",
            "PANES_BASE_URL": f"http://127.0.0.1:{self.server.server_port}",
        })
        self.env.start()

    def tearDown(self):
        self.env.stop()
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(timeout=5)

    async def test_real_discovery_ignores_ambient_credentials_and_denies_calls(self):
        async with SelectedComputerServer(execute=False, machine="m-selected") as server:
            self.assertEqual({tool.name for tool in await server.list_tools()}, WORK_TOOLS)
            with self.assertRaisesRegex(ValueError, "selected computer"):
                await server.call_tool("machine_get", {"id": "m-selected"})
        self.assertEqual(self.requests, [])

    async def test_documented_execution_sends_task_matching_actual_uploads(self):
        with patch.dict(os.environ, {"OPENAI_API_KEY": "local-model-test-inert"}), patch(
            "sys.argv", ["office_agent.py", "--machine", "m-selected", "--model", "fixture-model", "--execute"]
        ), patch("office_agent.Runner.run", new_callable=AsyncMock) as runner:
            runner.return_value = SimpleNamespace(new_items=[])
            await run(parse_args())
        runner.assert_awaited_once()
        task = runner.await_args.args[1]
        self.assertEqual(task, (ROOT / "COMPACT_TASK.md").read_text(encoding="utf-8"))
        self.assertIn("expense export", task)
        self.assertNotIn("monthly-close/", task)
        self.assertEqual(self.requests, [("GET", "/v1/machines/m-selected")])

        # Reconcile instructions against the actual helper's upload manifest,
        # not a second hand-maintained fixture or the expanded close history.
        spec = importlib.util.spec_from_file_location("compact_uploads", ROOT / "prepare.py")
        helper = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(helper)
        prepared = {remote.rsplit("/", 1)[-1] for _, remote in helper.uploads()}
        required = set(re.findall(r"`([A-Za-z0-9][A-Za-z0-9 ._-]*\.(?:csv|ps1|tsv|json))`", task))
        self.assertEqual(required, prepared)

    async def test_real_runner_filters_scope_before_transport(self):
        async with SelectedComputerServer(execute=True, machine="m-selected") as server:
            self.assertEqual({tool.name for tool in await server.list_tools()}, WORK_TOOLS)
            await require_running(server)
            self.assertEqual(self.requests, [("GET", "/v1/machines/m-selected")])
            self.requests.clear()
            model = ScriptedModel()
            result = await Runner.run(
                office_agent(server, model), "Protocol fixture only", max_turns=3,
                run_config=RunConfig(tracing_disabled=True),
            )
            self.assertEqual(model.calls, 2)
            self.assertEqual(result.final_output, "Protocol test complete")
            self.assertEqual(self.requests, [("GET", "/v1/machines/m-selected")])
            results = {item.raw_item["call_id"]: item.output for item in result.new_items if item.type == "tool_call_output_item"}
            self.assertIn("selected computer", str(results["wrong-computer"]))
            self.assertIn("signed-in desktop", str(results["wrong-desktop"]))

            for name in ("machine_start", "machine_stop", "machine_delete", "viewer", "account"):
                with self.assertRaisesRegex(ValueError, "selected computer"):
                    await server.call_tool(name, {"id": "m-selected"})
            self.assertEqual(self.requests, [("GET", "/v1/machines/m-selected")])

            self.state = "stopped"
            with self.assertRaisesRegex(SetupError, "existing running"):
                await require_running(server)
            self.state = "running"
            self.warning = "Review recovery before use"
            with self.assertRaisesRegex(SetupError, "recovery warning"):
                await require_running(server)
            self.assertTrue(all(method == "GET" for method, _ in self.requests))


if __name__ == "__main__":
    unittest.main()
