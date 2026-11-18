"""No cloud calls: selected-file downloads and independent failure checks."""
import contextlib
import hashlib
import io
import json
import os
from pathlib import Path
import shutil
import stat
import subprocess
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import download

ROOT = Path(__file__).parent


class DownloadTests(unittest.TestCase):
    def invoke(self, root, arguments, state='running', responses=None):
        machine = Mock(state=state)
        machine.files.read.side_effect = responses or [b'workbook', b'deck']
        client = Mock()
        client.machines.get.return_value = machine
        module = SimpleNamespace(Panes=Mock(return_value=client))
        output = io.StringIO()
        with patch.object(download, 'ROOT', root), patch.dict(sys.modules, {'panes': module}), patch.dict(os.environ, {'PANES_API_KEY': 'fixture-only'}), patch.object(sys, 'argv', ['download.py', *arguments]), contextlib.redirect_stdout(output):
            download.main()
        return machine, client, module, output.getvalue()

    def test_default_and_dry_run_do_not_import_sdk_or_write(self):
        for arguments in ([], ['--dry-run']):
            with tempfile.TemporaryDirectory() as temporary:
                machine, client, module, output = self.invoke(Path(temporary), arguments)
                module.Panes.assert_not_called()
                self.assertEqual(json.loads(output)['cloud_calls'], 0)
                self.assertEqual(list(Path(temporary).iterdir()), [])

    def test_downloads_exact_files_without_overwriting_or_lifecycle_calls(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            payload = [(ROOT / 'outputs' / name).read_bytes() for name in download.FILES]
            for attempt in range(2):
                machine, client, module, output = self.invoke(root, ['--machine', 'known-id', '--execute'], responses=payload)
                client.machines.get.assert_called_once_with('known-id')
                self.assertEqual([c.args[0] for c in machine.files.read.call_args_list], [f'{download.REMOTE}/{name}' for name in download.FILES])
                for call in machine.files.read.call_args_list:
                    self.assertEqual(call.kwargs, {'max_bytes': download.MAX_FILE_BYTES})
                self.assertEqual(len(client.mock_calls), 3)
                self.assertEqual(len(machine.mock_calls), 2)
            directories = list((root / 'local-results').iterdir())
            self.assertEqual(len(directories), 2)
            for directory in directories:
                receipt = json.loads((directory / 'download.json').read_text())
                for name, data, entry in zip(download.FILES, payload, receipt['files']):
                    self.assertEqual((directory / name).read_bytes(), data)
                    self.assertEqual(entry, {'name': name, 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()})
                if os.name != 'nt':
                    self.assertEqual(stat.S_IMODE(directory.stat().st_mode), 0o700)
                    for path in directory.iterdir():
                        self.assertEqual(stat.S_IMODE(path.stat().st_mode), 0o600)

    def test_stopped_computer_is_never_resumed_or_read(self):
        with tempfile.TemporaryDirectory() as temporary:
            with self.assertRaisesRegex(SystemExit, 'Select a running computer'):
                self.invoke(Path(temporary), ['--machine', 'known-id', '--execute'], state='stopped')
            self.assertEqual(list((Path(temporary) / 'local-results').iterdir()), [])

    def test_partial_failure_has_no_receipt_or_provider_error_leak(self):
        with tempfile.TemporaryDirectory() as temporary:
            with self.assertRaises(SystemExit) as raised:
                self.invoke(Path(temporary), ['--machine', 'known-id', '--execute'], responses=[b'partial', RuntimeError('PRIVATE_PROVIDER_DETAIL')])
            self.assertNotIn('PRIVATE_PROVIDER_DETAIL', str(raised.exception))
            directory, = (Path(temporary) / 'local-results').iterdir()
            self.assertEqual((directory / download.FILES[0]).read_bytes(), b'partial')
            self.assertFalse((directory / 'download.json').exists())

    def test_conflicting_flags_and_invalid_id_have_no_sdk_calls(self):
        for arguments in (['--machine', 'known-id'], ['--execute', '--dry-run'], ['--machine', '../elsewhere', '--execute'], ['--execute']):
            with tempfile.TemporaryDirectory() as temporary, contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit):
                    self.invoke(Path(temporary), arguments)
                self.assertEqual(list(Path(temporary).iterdir()), [])


class ValidationTests(unittest.TestCase):
    def run_validator(self, directory):
        return subprocess.run([sys.executable, str(ROOT / 'validate.py'), '--outputs-dir', str(directory)], text=True, capture_output=True)

    def test_selected_native_files_pass_without_reserializing(self):
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            for name in download.FILES:
                shutil.copyfile(ROOT / 'outputs' / name, directory / name)
            before = {name: (directory / name).read_bytes() for name in download.FILES}
            result = self.run_validator(directory)
            self.assertEqual(result.returncode, 0, result.stderr)
            report = json.loads((directory / 'validation.json').read_text())
            self.assertEqual(report['result'], 'PASS')
            self.assertEqual(report['outputs_directory'], str(directory.resolve()))
            self.assertEqual(report['formula_outputs_verified'], 60)
            for name, data in before.items():
                self.assertEqual((directory / name).read_bytes(), data)
                self.assertEqual(report['hashes'][name], hashlib.sha256(data).hexdigest())

    def test_missing_and_corrupt_selected_files_clear_old_pass_without_fallback(self):
        for content in (None, b'not-an-office-file'):
            with tempfile.TemporaryDirectory() as temporary:
                directory = Path(temporary)
                if content is not None:
                    (directory / download.FILES[0]).write_bytes(content)
                (directory / 'validation.json').write_text('{"result":"PASS"}')
                result = self.run_validator(directory)
                self.assertNotEqual(result.returncode, 0)
                self.assertEqual(json.loads((directory / 'validation.json').read_text())['result'], 'FAIL')
                self.assertNotIn('Traceback', result.stderr)


if __name__ == '__main__':
    unittest.main()
