"""Download a compact Office run from one existing running Panes computer."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import tempfile

from prepare import ROOT, REMOTE, prepare_private_directory

FILES = ('October close.xlsx', 'October review.pptx')
MAX_FILE_BYTES = 64 * 1024 * 1024


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--machine', help='Existing running computer containing your completed compact task.')
    parser.add_argument('--execute', action='store_true', help='Read the selected computer and download two files; no lifecycle changes.')
    parser.add_argument('--dry-run', action='store_true', help='Show file paths without SDK, network or local writes.')
    args = parser.parse_args()
    if args.dry_run or not args.execute:
        if args.execute or args.machine:
            parser.error('Use --machine ID --execute, or --dry-run without a machine.')
        print(json.dumps({'mode': 'dry-run', 'remote_files': [f'{REMOTE}/{name}' for name in FILES], 'cloud_calls': 0}))
        return
    if not args.machine or not re.fullmatch(r'[A-Za-z0-9_-]+', args.machine):
        parser.error('--execute requires an existing --machine ID.')
    if not os.environ.get('PANES_API_KEY'):
        parser.error('Set PANES_API_KEY in your secret environment.')

    try:
        private = prepare_private_directory(ROOT / 'local-results')
        from panes import Panes
        machine = Panes().machines.get(args.machine)
        if machine.state != 'running':
            raise SystemExit('Select a running computer. This helper never creates, starts or resumes one.')
        directory = Path(tempfile.mkdtemp(prefix='compact-', dir=private))
        print(f'Download directory: {directory.relative_to(ROOT)}')
        receipt = {'machine_id': args.machine, 'files': []}
        for name in FILES:
            data = machine.files.read(f'{REMOTE}/{name}', max_bytes=MAX_FILE_BYTES)
            if not isinstance(data, bytes) or not data:
                raise ValueError('Expected a nonempty file.')
            # Exclusive, owner-only POSIX file; Windows inherits the private ACL.
            fd = os.open(directory / name, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with os.fdopen(fd, 'wb') as output:
                output.write(data)
            receipt['files'].append({'name': name, 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()})
        fd = os.open(directory / 'download.json', os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, 'w') as output:
            json.dump(receipt, output, indent=2)
            output.write('\n')
        print('Both files downloaded. Run validate.py with this --outputs-dir before trusting the result.')
        print('The computer is still running. Stop it in Panes after checking your work.')
    except SystemExit:
        raise
    except Exception:
        # Provider errors may include credentials or document details.
        raise SystemExit('Download failed. Inspect local-results before retrying; no remote file or lifecycle was changed.') from None


if __name__ == '__main__':
    main()
