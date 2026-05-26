#!/usr/bin/env python3
"""
cloud_build.py  —  runs on GitHub Actions (24/7 cloud automation)
─────────────────────────────────────────────────────────────────
1. Authenticates to Google Drive using the SAME OAuth login that
   works on the Mac (token.json), supplied via the GOOGLE_OAUTH_TOKEN
   secret. Refreshes the access token headlessly — no browser, no
   service-account key needed (the org blocks SA keys).
2. Downloads the production .xlsx from Drive.
3. Runs generate_enicar_html.py to build the dashboard HTML.
4. Copies the result to index.html (the page the live URL serves).

The workflow (.github/workflows/update-dashboard.yml) then commits
index.html + Enicar_Dashboard.html if anything changed.

If GOOGLE_OAUTH_TOKEN is missing, this exits cleanly (code 0) so the
workflow does NOT fail / send error emails before setup is finished.
"""

import os, sys, json, io, shutil, subprocess

DRIVE_FILE_ID = os.environ.get('DRIVE_FILE_ID', '').strip()
OAUTH_JSON    = os.environ.get('GOOGLE_OAUTH_TOKEN', '').strip()
HERE          = os.path.dirname(os.path.abspath(__file__))
XLSX_OUT      = os.path.join(HERE, 'Enicar_Dashboard_Template.xlsx')


def main():
    if not OAUTH_JSON:
        print('ℹ️  GOOGLE_OAUTH_TOKEN secret not set yet — skipping cloud build.')
        print('    (Add the secret in GitHub → Settings → Secrets to activate.)')
        return 0
    if not DRIVE_FILE_ID:
        print('❌  DRIVE_FILE_ID env var not set in the workflow.')
        return 1

    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload

    # ── Auth with the existing OAuth login (refresh headlessly) ─
    info  = json.loads(OAUTH_JSON)
    creds = Credentials.from_authorized_user_info(
        info, scopes=['https://www.googleapis.com/auth/drive']
    )
    creds.refresh(Request())          # uses refresh_token — no browser
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

    # ── Product-name mismatch check (emails the team if needed) ────
    # Runs every build, but emails are deduped by its own state file
    # so the team only gets a first alert + 3-day reminders. Won't fail
    # the build if email creds are missing or SMTP errors out.
    try:
        r = subprocess.run(
            [sys.executable, os.path.join(HERE, 'check_product_mismatches.py')],
            env=env, capture_output=True, text=True, timeout=60
        )
        if r.stdout: print(r.stdout.strip())
        if r.returncode != 0 and r.stderr:
            print('mismatch checker stderr:', r.stderr.strip())
    except Exception as e:
        print(f'mismatch checker skipped: {e}')

    return 0


if __name__ == '__main__':
    sys.exit(main())
