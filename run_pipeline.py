"""Refresh the SharePoint line list; optionally run vaccination reconciliation."""
import argparse
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--secrets-env', type=Path, default=ROOT / 'config/secrets.env')
    parser.add_argument('--share-url', help='Override the existing AFRO line list sharing link')
    parser.add_argument('--reconcile', action='store_true', help='Run existing reconciliation after a successful refresh')
    args = parser.parse_args()
    command = [sys.executable, str(ROOT / 'src/extraction/fetch_afro_line_list.py'),
               '--secrets-env', str(args.secrets_env.resolve())]
    if args.share_url:
        command.extend(['--share-url', args.share_url])
    result = subprocess.run(command, cwd=ROOT)
    if result.returncode:
        return result.returncode
    if args.reconcile:
        return subprocess.run([sys.executable, str(ROOT / 'src/transformation/reconcile_vaccination.py')], cwd=ROOT).returncode
    return 0

if __name__ == '__main__':
    sys.exit(main())
