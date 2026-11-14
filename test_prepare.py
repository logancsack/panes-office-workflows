"""Real OS handoff-permission checks, plus upload-free recovery behavior."""

import contextlib
import io
import json
import os
from pathlib import Path
import stat
import subprocess
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import prepare


class HandoffTests(unittest.TestCase):
    def test_private_output_and_atomic_replacement(self):
        with tempfile.TemporaryDirectory() as temporary:
            directory = prepare.prepare_private_directory(Path(temporary) / 'private handoff')
            machine = Mock()
            machine.viewer.return_value = 'https://example.invalid/private-bearer'
            target = directory / 'viewer.json'
            target.write_text('old')
            if os.name != 'nt':
                target.chmod(0o644)
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                actual = prepare.save_viewer(machine, directory)
            self.assertEqual(actual, target)
            self.assertEqual(json.loads(target.read_text()), {'url': machine.viewer.return_value})
            self.assertNotIn(machine.viewer.return_value, output.getvalue())
            self.assertEqual([p.name for p in directory.iterdir()], ['viewer.json'])
            if os.name == 'nt':
                script = '''$ErrorActionPreference='Stop'
$acl = [System.IO.File]::GetAccessControl($env:PANES_HANDOFF_FILE)
$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$parent = [System.IO.Directory]::GetAccessControl([System.IO.Path]::GetDirectoryName($env:PANES_HANDOFF_FILE))
if (-not $parent.AreAccessRulesProtected) { throw 'Unprotected directory' }
foreach ($rule in $acl.Access) {
  if ($rule.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value -ne $sid) { throw 'Unexpected reader' }
}
if ($acl.Access.Count -lt 1) { throw 'No user access' }
'''
                result = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', script], env={**os.environ, 'PANES_HANDOFF_FILE': str(target)}, capture_output=True, text=True)
                self.assertEqual(result.returncode, 0, result.stderr)
            else:
                self.assertEqual(stat.S_IMODE(directory.stat().st_mode), 0o700)
                self.assertEqual(stat.S_IMODE(target.stat().st_mode), 0o600)

    def test_viewer_only_recovers_without_uploading(self):
        with tempfile.TemporaryDirectory() as temporary:
            machine = Mock(state='running')
            machine.viewer.return_value = 'https://example.invalid/private-bearer'
            client = Mock()
            client.machines.get.return_value = machine
            module = SimpleNamespace(Panes=Mock(return_value=client))
            with patch.object(prepare, 'ROOT', Path(temporary)), patch.dict('sys.modules', {'panes': module}), patch('sys.argv', ['prepare.py', '--machine', 'known-id', '--execute', '--viewer-only']), contextlib.redirect_stdout(io.StringIO()) as output:
                prepare.main()
            machine.exec.assert_not_called()
            machine.files.write.assert_not_called()
            client.machines.create.assert_not_called()
            machine.start.assert_not_called()
            self.assertNotIn(machine.viewer.return_value, output.getvalue())
            self.assertTrue((Path(temporary) / 'local-results' / 'viewer.json').is_file())

    def test_permission_failure_precedes_remote_changes(self):
        module = SimpleNamespace(Panes=Mock())
        with patch.dict('sys.modules', {'panes': module}), patch.object(prepare, 'prepare_private_directory', side_effect=OSError('denied')), patch('sys.argv', ['prepare.py', '--machine', 'known-id', '--execute']):
            with self.assertRaises(OSError):
                prepare.main()
        module.Panes.assert_not_called()


if __name__ == '__main__':
    unittest.main()
