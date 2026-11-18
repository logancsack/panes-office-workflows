"""Exercise the pinned SDK over loopback; no key, VM, model or cloud calls."""
import contextlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import io
import json
import os
from pathlib import Path
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch
from urllib.parse import parse_qs, urlsplit

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
import download


class TransportTests(unittest.TestCase):
    def test_real_sdk_downloads_only_selected_task_files(self):
        requests = []
        payload = {f'{download.REMOTE}/{name}': (ROOT / 'outputs' / name).read_bytes() for name in download.FILES}

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *_):
                pass

            def do_GET(self):
                requests.append(self.path)
                route = urlsplit(self.path)
                if route.path == '/v1/machines/fixture-id':
                    data = json.dumps({'id': 'fixture-id', 'state': 'running'}).encode()
                elif route.path == '/v1/machines/fixture-id/files':
                    data = payload[parse_qs(route.query)['path'][0]]
                else:
                    self.send_error(404)
                    return
                self.send_response(200)
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)

        server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
        worker = threading.Thread(target=server.serve_forever)
        worker.start()
        try:
            with tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                with patch.object(download, 'ROOT', root), patch.dict(os.environ, {'PANES_API_KEY': 'local-fixture-only', 'PANES_BASE_URL': f'http://127.0.0.1:{server.server_port}'}), patch.object(sys, 'argv', ['download.py', '--machine', 'fixture-id', '--execute']), contextlib.redirect_stdout(io.StringIO()):
                    download.main()
                directory, = (root / 'local-results').iterdir()
                for name in download.FILES:
                    self.assertEqual((directory / name).read_bytes(), payload[f'{download.REMOTE}/{name}'])
                self.assertEqual(len(requests), 3)
                self.assertEqual(requests[0], '/v1/machines/fixture-id')
        finally:
            server.shutdown()
            server.server_close()
            worker.join()


if __name__ == '__main__':
    unittest.main()
