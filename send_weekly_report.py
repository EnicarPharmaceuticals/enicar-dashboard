#!/usr/bin/env python3
"""
send_weekly_report.py  —  emails the weekly data review
─────────────────────────────────────────────────────────────────
Sends a friendly, plain-English summary of the weekly data check.

Credentials:
  • GMAIL_SENDER / GMAIL_APP_PASSWORD env vars (used in the cloud), or
  • falls back to MY_EMAIL / APP_PASSWORD in email_config.py (local).

Recipients:
  • If ALWAYS_SEND_TEAM=1 (the cloud) → always the whole team.
  • Otherwise → first run is a REVIEW copy to the owner, then the team.

Only emails when there is actually something to fix.
"""

import os, sys, smtplib, ssl
from email.message import EmailMessage
from datetime import datetime

HERE   = os.path.dirname(os.path.abspath(__file__))
ROOT   = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
REPORT = os.path.join(ROOT, 'Weekly_Data_Review.txt')
FLAG   = os.path.join(HERE, '.weekly_email_sent_once')

# Credentials — env first (cloud), then local email_config.py
SENDER = os.environ.get('GMAIL_SENDER', '').strip()
APP_PW = os.environ.get('GMAIL_APP_PASSWORD', '').strip()
if not (SENDER and APP_PW):
    try:
        sys.path.insert(0, HERE)
        import email_config as cfg
        SENDER = SENDER or cfg.MY_EMAIL
        APP_PW = APP_PW or cfg.APP_PASSWORD
    except Exception:
        pass

TEAM        = ['packing@enicarpharma.com', 'exports@enicarpharma.com', 'swaralisave@enicarpharma.com']
REVIEW_ONLY = ['swaralisave@enicarpharma.com']

CLEAN_CUSTOMERS = 'All customer names are recognised'
CLEAN_BATCHES   = 'every batch number lines up'


def has_findings(text):
    return not (CLEAN_CUSTOMERS in text and CLEAN_BATCHES in text)


def build_body(report_text, first_time):
    stamp = datetime.now().strftime('%d %B %Y')
    if first_time:
        opening = ("Hi,\n\nThis is the very first weekly data check — sent to you only so you "
                   "can have a look first. From next week it will go straight to the team "
                   f"({', '.join(TEAM)}).\n\n")
    else:
        opening = "Hi team,\n\n"

    intro = (
        "Here is this week's quick check on our production logs (Filling, Packing and "
        "Dispatch).\n\n"
        "When the same batch number or customer is written two slightly different ways, "
        "the dashboard can't link filling → packing → dispatch correctly. A small fix in "
        "the Google Sheet keeps the numbers accurate.\n\n"
        "How to read it:\n"
        "  • \"change X to Y\"  →  please update X to Y in the sheet (Y is the correct one).\n"
        "  • \"please double-check\"  →  our best guess; kindly confirm which spelling is right.\n\n"
        "──────────────────────────────────────────────────\n\n"
    )
    closing = ("\n──────────────────────────────────────────────────\n\n"
               "Thanks for keeping the data clean!\n\n"
               "— Enicar Dashboard (automatic weekly check)")
    return opening + intro + report_text.strip() + closing


def main():
    if not (SENDER and APP_PW):
        print('No email credentials available — skipping send.'); return 0
    if not os.path.exists(REPORT):
        print('No report file to send.'); return 0

    report_text = open(REPORT, encoding='utf-8').read()
    if not has_findings(report_text):
        print('Nothing to fix this week — no email sent.'); return 0

    always_team = os.environ.get('ALWAYS_SEND_TEAM') == '1'
    first_time  = (not always_team) and (not os.path.exists(FLAG))
    recipients  = REVIEW_ONLY if first_time else TEAM

    stamp = datetime.now().strftime('%d %b %Y')
    msg = EmailMessage()
    msg['From'] = SENDER
    msg['To']   = ', '.join(recipients)
    msg['Subject'] = (f'[Please review] Enicar weekly data check — {stamp}' if first_time
                      else f'Enicar weekly data check — a few corrections needed ({stamp})')
    msg.set_content(build_body(report_text, first_time))
    msg.add_attachment(report_text.encode('utf-8'), maintype='text', subtype='plain',
                       filename='Weekly_Data_Review.txt')

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ctx) as s:
        s.login(SENDER, APP_PW)
        s.send_message(msg)

    if first_time:
        open(FLAG, 'w').write(datetime.now().isoformat())
        print(f'✓ Review copy sent to {recipients[0]} (team auto-send starts next run).')
    else:
        print(f'✓ Sent to: {", ".join(recipients)}')
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        print(f'✗ Email send failed: {e}')
        sys.exit(1)
