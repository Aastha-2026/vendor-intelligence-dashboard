"""Watch the OneDrive master workbook and publish changes to the live dashboard.

Polls the workbook for changes, and on each change: copies it into the repo,
regenerates vendors.json / vendors-data.js, then commits and pushes so GitHub
Pages redeploys. Typical time from Excel save to live site: 1-2 minutes.

Run it with:  tools\\watch_onedrive.cmd   (or: python tools/watch_onedrive.py)
Stop it with: Ctrl+C
"""

import os
import shutil
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOURCE = os.environ.get(
    'VENDOR_XLSX',
    os.path.join(os.path.expanduser('~'), 'OneDrive - Walmart Inc', 'Aastha',
                 'Vendor Intelligence Data', 'Vendor Bucketing-Master list-Final.xlsx'))
REPO_XLSX = os.path.join(ROOT, 'Vendor bucketing', 'Vendor Bucketing-Master list-Final.xlsx')
TRACKED = ['Vendor bucketing/Vendor Bucketing-Master list-Final.xlsx',
           'Vendor bucketing/vendors.json',
           'Vendor bucketing/vendors-data.js']

POLL_SECONDS = 10
LOG_FILE = os.path.join(ROOT, 'tools', 'sync.log')


def log(msg):
    line = '[%s] %s' % (time.strftime('%Y-%m-%d %H:%M:%S'), msg)
    print(line, flush=True)
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except OSError:
        pass


def git(*args, check=True):
    r = subprocess.run(['git'] + list(args), cwd=ROOT, capture_output=True, text=True)
    if check and r.returncode != 0:
        raise RuntimeError('git %s failed: %s' % (' '.join(args), r.stderr.strip()))
    return r.stdout.strip()


def signature(path):
    """Change fingerprint; None while the file is missing or locked by Excel."""
    try:
        s = os.stat(path)
        with open(path, 'rb'):
            pass
        return (s.st_mtime_ns, s.st_size)
    except (OSError, PermissionError):
        return None


def publish():
    shutil.copy2(SOURCE, REPO_XLSX)

    r = subprocess.run([sys.executable, os.path.join(ROOT, 'tools', 'build_vendors.py'), REPO_XLSX],
                       cwd=ROOT, capture_output=True, text=True)
    if r.returncode != 0:
        log('BUILD FAILED - nothing published:')
        log((r.stdout + r.stderr).strip())
        return
    log(r.stdout.strip().splitlines()[0])

    if not git('status', '--porcelain', '--', *TRACKED):
        log('Workbook changed but the vendor data is identical - nothing to publish.')
        return

    git('add', '--', *TRACKED)
    git('commit', '-m', 'data: sync vendor master from OneDrive')
    try:
        git('push', 'origin', 'HEAD')
    except RuntimeError as e:
        log('PUSH FAILED (commit is saved locally, push it manually): %s' % e)
        return
    log('Published - GitHub Pages will redeploy in about a minute.')


def main():
    if not os.path.exists(SOURCE):
        raise SystemExit('Workbook not found:\n  %s\nSet VENDOR_XLSX to override.' % SOURCE)

    log('Watching: %s' % SOURCE)
    log('Press Ctrl+C to stop.')
    last = signature(SOURCE)
    pending = None

    while True:
        time.sleep(POLL_SECONDS)
        sig = signature(SOURCE)
        if sig is None:
            continue            # file open in Excel or mid-sync
        if sig != last:
            if pending != sig:  # seen once; wait for it to stop changing
                pending = sig
                continue
            log('Change detected - publishing...')
            try:
                publish()
            except Exception as e:
                log('ERROR: %s' % e)
            last = signature(SOURCE) or sig
            pending = None
        else:
            pending = None


if __name__ == '__main__':
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log('Stopped.')
