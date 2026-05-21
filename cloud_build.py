#!/usr/bin/env python3
"""
cloud_build.py  —  runs on GitHub Actions (24/7 cloud automation)
─────────────────────────────────────────────────────────────────
1. Authenticates to Google Drive with a SERVICE ACCOUNT
   (JSON provided via the GOOGLE_SERVICE_ACCOUNT secret).
2. Downloads the production .xlsx from Drive.
3. Runs generate_enicar_html.py to build the dashboard HTML.
4. Copies the result to index.html (the page the live URL serves).

The workflow (.github/workflows/update-dashboard.yml) then commits
index.html + Enicar_Dashboard.html if anything changed.

If the GOOGLE_SERVICE_ACCOUNT secret is missing, this exits cleanly
(code 0) so the workflow does NOT fail / send error emails — handy
before the one-time setup is finished.
"""

import os, sys, json, io, shutil, subprocess

DRIVE_FILE_ID = os.environ.get('DRIVE_FILE_ID', '').strip()
SA_JSON       = os.environ.get('GOOGLE_SERVICE_ACCOUNT', '').strip()
HERE          = os.path.dirname(os.path.abspath(__file__))
XLSX_OUT      = os.path.join(HERE, 'Enicar_Dashboard_Template.xlsx')


def main():
    if not SA_JSON:
        print('ℹ️  GOOGLE_SERVICE_ACCOUNT secret not set yet — skipping cloud build.')
        print('    (Add the secret in GitHub → Settings → Secrets to activate.)')
        return 0
    if not DRIVE_FILE_ID:
        print('❌  DRIVE_FILE_ID env var not set in the workflow.')
        return 1

    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload

    # ── Auth with service account ──────────────────────────────
    info  = json.loads(SA_JSON)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    svc = build('drive', 'v3', credentials=creds)

    # ── Download the .xlsx (raw file, not a native Sheet) ──────
    request = svc.files().get_media(fileId=DRIVE_FILE_ID)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    with open(XLSX_OUT, 'wb') as f:
        f.write(fh.getvalue())
    print(f'✓ Downloaded Drive file → Enicar_Dashboard_Template.xlsx ({len(fh.getvalue())//1024} KB)')

    # ── Generate the dashboard HTML ────────────────────────────
    env = dict(os.environ, DASHBOARD_ROOT=HERE)
    result = subprocess.run(
        [sys.executable, os.path.join(HERE, 'generate_enicar_html.py')],
        env=env, capture_output=True, text=True
    )
    print(result.stdout)
    if result.returncode != 0:
        print(result.stderr)
        return result.returncode

    # ── Copy generated HTML → index.html (served at the live URL) ─
    generated = os.path.join(HERE, 'Enicar_Dashboard.html')
    if not os.path.exists(generated):
        print('❌  Enicar_Dashboard.html was not generated.')
        return 1
    shutil.copyfile(generated, os.path.join(HERE, 'index.html'))
    print('✓ index.html updated from generated dashboard.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
