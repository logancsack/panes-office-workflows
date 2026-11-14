"""Upload the reference task to an existing running Panes computer; no implicit spend."""
import argparse
from pathlib import Path
import re
import json
import os
import subprocess
import tempfile
import stat

ROOT = Path(__file__).parent
REMOTE = 'C:/work/Panes-office-demo'


def prepare_private_directory(path):
    """Restrict the handoff directory before any remote files are changed."""
    reparse = path.exists() and getattr(path.lstat(), 'st_file_attributes', 0) & getattr(stat, 'FILE_ATTRIBUTE_REPARSE_POINT', 0)
    if path.is_symlink() or reparse:
        raise OSError('The private output directory must not be a link or junction.')
    path.mkdir(mode=0o700, exist_ok=True)
    if os.name == 'nt':
        script = '''$ErrorActionPreference = 'Stop'
$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
$acl = [System.Security.AccessControl.DirectorySecurity]::new()
$acl.SetAccessRuleProtection($true, $false)
$rule = [System.Security.AccessControl.FileSystemAccessRule]::new($sid, [System.Security.AccessControl.FileSystemRights]::FullControl, [System.Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit', [System.Security.AccessControl.PropagationFlags]::None, [System.Security.AccessControl.AccessControlType]::Allow)
$acl.AddAccessRule($rule)
Set-Acl -LiteralPath $env:PANES_HANDOFF_DIRECTORY -AclObject $acl
'''
        result = subprocess.run(
            ['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', script],
            env={**os.environ, 'PANES_HANDOFF_DIRECTORY': str(path.resolve())},
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode:
            # This runs before a viewer is requested; stderr has no bearer URL.
            raise OSError('Could not restrict the private output directory to your Windows user: ' + result.stderr.strip())
    else:
        path.chmod(0o700)
    return path


def save_viewer(machine, directory):
    # mkstemp creates a new owner-only POSIX file or inherits the restricted
    # Windows directory ACL. Replacing the target also discards old file ACLs.
    fd, temporary = tempfile.mkstemp(prefix='.viewer-', suffix='.json', dir=directory)
    try:
        with os.fdopen(fd, 'w') as file:
            json.dump({'url': machine.viewer()}, file)
        target = directory / 'viewer.json'
        os.replace(temporary, target)
        return target
    finally:
        Path(temporary).unlink(missing_ok=True)


def uploads():
    return [(p, f'{REMOTE}/{p.name}') for directory in ('inputs', 'reference')
            for p in sorted((ROOT / directory).iterdir()) if p.is_file()]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--machine', help='An existing running computer you own')
    parser.add_argument('--execute', action='store_true', help='Allow directory creation and uploads')
    parser.add_argument('--dry-run', action='store_true', help='List files, with no SDK or network calls')
    parser.add_argument('--viewer-only', action='store_true', help='Save a private viewer without creating a task directory or uploading again; requires --execute')
    args = parser.parse_args()
    if args.dry_run:
        for local, remote in uploads():
            print(f'{local.relative_to(ROOT)} -> {remote}')
        return
    if not args.execute or not args.machine or not re.fullmatch(r'[A-Za-z0-9_-]+', args.machine):
        parser.error('Preview with --dry-run, or give --machine ID --execute for your running computer.')
    private = prepare_private_directory(ROOT / 'local-results')
    from panes import Panes
    machine = Panes().machines.get(args.machine)
    if machine.state != 'running':
        raise SystemExit('Computer is not running. This helper does not create, start or resume computers.')
    if not args.viewer_only:
        result = machine.exec("$ErrorActionPreference='Stop'; if(Test-Path 'C:/work/Panes-office-demo'){throw 'Task directory already exists; inspect it before proceeding'}; New-Item -ItemType Directory -Path 'C:/work/Panes-office-demo' | Out-Null", desktop=True, timeout_s=30)
        if result.exit_code != 0 or result.timed_out:
            raise SystemExit('Could not create a fresh task directory. Inspect the desktop before retrying; --viewer-only can save a handoff without uploading again.')
        for local, remote in uploads():
            machine.files.write(remote, local.read_bytes())
        print('Inputs and recorded reference stages uploaded. Give TASK.md to your agent.')
    save_viewer(machine, private)
    print('Private viewer saved to local-results/viewer.json (not printed or committed).')
    print('Stop and keep the computer after the agent finishes. Running time uses prepaid credit.')


if __name__ == '__main__':
    main()
