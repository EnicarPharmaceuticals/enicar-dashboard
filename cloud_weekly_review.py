#!/usr/bin/env python3
"""
cloud_weekly_review.py  —  weekly data review in the cloud (GitHub Actions)
──────────────────────────────────────────────────────────────────────────
1. Downloads the latest .xlsx from Drive (same OAuth login as the dashboard).
2. Runs the customer-name and batch-number checks → Weekly_Data_Review.txt.
3. Emails the report to the team (send_weekly_report.py, always-team mode).

Secrets used (GitHub → Settings → Secrets):
  GOOGLE_OAUTH_TOKEN   – same one the dashboard already uses
  GMAIL_APP_PASSWORD   – Gmail app password for sending
Exits cleanly if a secret is missing, so the workflow never errors out.
"""

import os, sys, io, json, subprocess

HERE          = os.path.dirname(os.path.abspath(__file__))
DRIVE_FILE_ID = os.environ.get('DRIVE_FILE_ID', '').strip()
OAUTH_JSON    = os.environ.get('GOOGLE_OAUTH_TOKEN', '').strip()
XLSX_OUT      = os.path.join(HERE, 'Enicar_Dashboard_Template.xlsx')
REPORT        = os.path.join(HERE, 'Weekly_Data_Review.txt')


def download():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    creds = Credentials.from_authorized_user_info(
        json.loads(OAUTH_JSON), scopes=['https://www.googleapis.com/auth/drive'])
    creds.refresh(Request())
    svc = build('drive', 'v3', credentials=creds)
    fh = io.BytesIO()
    dl = MediaIoBaseDownload(fh, svc.files().get_media(fileId=DRIVE_FILE_ID))
    done = False
    while not done:
        _, done = dl.next_chunk()
    open(XLSX_OUT, 'wb').write(fh.getvalue())
    print(f'✓ Downloaded data ({len(fh.getvalue())//1024} KB)')


def run(script):
    env = dict(os.environ, DASHBOARD_ROOT=HERE)
    return subprocess.run([sys.executable, os.path.join(HERE, script)],
                          env=env, capture_output=True, text=True).stdout


def main():
    if not OAUTH_JSON or not DRIVE_FILE_ID:
        print('ℹ️  Missing GOOGLE_OAUTH_TOKEN / DRIVE_FILE_ID — skipping.'); return 0
    download()

    from datetime import datetime
    with open(REPORT, 'w', encoding='utf-8') as f:
        f.write('Enicar — Weekly Data Review\n')
        f.write('Generated: ' + datetime.now().strftime('%a %d %b %Y') + '\n')
        f.write('=' * 50 + '\n\n')
        f.write(run('check_party_aliases.py') + '\n')
        f.write(run('check_batch_typos.py') + '\n')
        f.write(run('check_date_sanity.py') + '\n')
        f.write(run('check_rm_pipeline.py') + '\n')
    print('✓ Report built.')

    # Email it (always-team mode in the cloud).
    env = dict(os.environ, DASHBOARD_ROOT=HERE, ALWAYS_SEND_TEAM='1')
    r = subprocess.run([sys.executable, os.path.join(HERE, 'send_weekly_report.py')],
                       env=env, capture_output=True, text=True)
    print(r.stdout.strip() or r.stderr.strip())
    return 0


if __name__ == '__main__':
    sys.exit(main())
