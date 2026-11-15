"""Protocol/agent checks against a local HTTP fixture; no Office or model run."""

import json
import os
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from unittest.mock import patch

from langchain.agents import create_agent
from langchain_core.language_models import BaseChatModel
from langchain_core.messages import AIMessage, ToolMessage
from langchain_core.outputs import ChatGeneration, ChatResult

from office_agent import EXPECTED_TOOLS, WORK_TOOLS, adapter, scope_middleware
from panes_state import require_running


class ScriptedModel(BaseChatModel):
    """Deterministic unit-test messages, not a hosted model or Office demo."""

    @property
    def _llm_type(self):
        return "scripted-protocol-test"

    def bind_tools(self, tools, **kwargs):
        return self

    def _generate(self, messages, stop=None, run_manager=None, **kwargs):
        if any(isinstance(message, ToolMessage) for message in messages):
            message = AIMessage(content="Test complete")
        else:
            message = AIMessage(content="", tool_calls=[
                {"id": "wrong-computer", "name": "machine_get", "args": {"id": "m-other"}},
                {"id": "selected-computer", "name": "machine_get", "args": {"id": "m-selected"}},
                {"id": "wrong-desktop", "name": "exec", "args": {"id": "m-selected", "command": "Get-Date", "desktop": False}},
            ])
        return ChatResult(generations=[ChatGeneration(message=message)])


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
                payload = json.dumps({"id": "m-selected", "state": owner.state, "stop_warning": owner.warning}).encode()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)

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

    async def test_real_discovery_ignores_ambient_api_credentials(self):
        async with adapter(execute=False) as connection:
            tools = await connection.list_tools()
            self.assertEqual({tool.name for tool in tools}, EXPECTED_TOOLS)
            self.assertTrue(all(tool.args.get("id") for tool in tools if tool.name in WORK_TOOLS))
        self.assertEqual(self.requests, [])

    async def test_real_protocol_preflight_and_agent_scope(self):
        async with adapter(execute=True) as connection:
            tools = [tool for tool in await connection.list_tools() if tool.name in WORK_TOOLS]
            await require_running(tools, "m-selected")
            self.assertEqual(self.requests, [("GET", "/v1/machines/m-selected")])
            self.requests.clear()
            agent = create_agent(ScriptedModel(), tools, middleware=[scope_middleware("m-selected")])
            result = await agent.ainvoke({"messages": [{"role": "user", "content": "Protocol test only"}]})
            replies = {message.tool_call_id: message for message in result["messages"] if isinstance(message, ToolMessage)}
            self.assertEqual(replies["wrong-computer"].status, "error")
            self.assertEqual(replies["wrong-desktop"].status, "error")
            self.assertEqual(replies["selected-computer"].status, "success")
            self.assertEqual(self.requests, [("GET", "/v1/machines/m-selected")])

            self.state = "stopped"
            with self.assertRaisesRegex(ValueError, "existing running"):
                await require_running(tools, "m-selected")
            self.state = "running"
            self.warning = "Review recovery before use"
            with self.assertRaisesRegex(ValueError, "recovery warning"):
                await require_running(tools, "m-selected")
            self.assertTrue(all(method == "GET" for method, _ in self.requests))


if __name__ == "__main__":
    unittest.main()
