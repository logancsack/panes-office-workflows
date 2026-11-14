"""Upload the reference task to an existing running Panes computer; no implicit spend."""
import argparse
from pathlib import Path
import re
import json
import os

ROOT = Path(__file__).parent
REMOTE = 'C:/work/Panes-office-demo'


def uploads():
    return [(p, f'{REMOTE}/{p.name}') for directory in ('inputs', 'reference')
            for p in sorted((ROOT / directory).iterdir()) if p.is_file()]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--machine', help='An existing running computer you own')
    parser.add_argument('--execute', action='store_true', help='Allow directory creation and uploads')
    parser.add_argument('--dry-run', action='store_true', help='List files, with no SDK or network calls')
    args = parser.parse_args()
    if args.dry_run:
        for local, remote in uploads():
            print(f'{local.relative_to(ROOT)} -> {remote}')
        return
    if not args.execute or not args.machine or not re.fullmatch(r'[A-Za-z0-9_-]+', args.machine):
        parser.error('Preview with --dry-run, or give --machine ID --execute for your running computer.')
    from panes import Panes
    machine = Panes().machines.get(args.machine)
    if machine.state != 'running':
        raise SystemExit('Computer is not running. This helper does not create, start or resume computers.')
    result = machine.exec("$ErrorActionPreference='Stop'; if(Test-Path 'C:/work/Panes-office-demo'){throw 'Task directory already exists; inspect it before proceeding'}; New-Item -ItemType Directory -Path 'C:/work/Panes-office-demo' | Out-Null", desktop=True, timeout_s=30)
    if result.exit_code != 0 or result.timed_out:
        raise SystemExit('Could not create a fresh task directory. Inspect the desktop before retrying.')
    for local, remote in uploads():
        machine.files.write(remote, local.read_bytes())
    print('Inputs and recorded reference stages uploaded. Give TASK.md to your agent.')
    private = ROOT / 'local-results'
    private.mkdir(mode=0o700, exist_ok=True)
    fd = os.open(private / 'viewer.json', os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    os.fchmod(fd, 0o600)
    with os.fdopen(fd, 'w') as file:
        json.dump({'url': machine.viewer()}, file)
    print('Private viewer saved to local-results/viewer.json (not printed or committed).')
    print('Stop and keep the computer after the agent finishes. Running time uses prepaid credit.')


if __name__ == '__main__':
    main()
