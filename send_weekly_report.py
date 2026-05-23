#!/usr/bin/env python3
"""
send_weekly_report.py  —  emails the weekly data review
─────────────────────────────────────────────────────────────────
• First ever run  → sends a REVIEW copy to the owner only.
• Every run after → auto-sends to the whole team.
• Sends ONLY when the report actually contains something to fix.

Uses the Gmail App Password already in email_config.py (SMTP) — no
extra Google permissions needed. Run by run_party_review.sh weekly.
"""

import os, sys, smtplib, ssl
from email.message import EmailMessage
from datetime import datetime

HERE   = os.path.dirname(os.path.abspath(__file__))
ROOT   = os.path.join(HERE, '..')
REPORT = os.path.join(ROOT, 'Weekly_Data_Review.txt')
FLAG   = os.path.join(HERE, '.weekly_email_sent_once')

sys.path.insert(0, HERE)
import email_config as cfg

SENDER     = cfg.MY_EMAIL                 # authenticated Gmail account
APP_PW     = cfg.APP_PASSWORD
TEAM       = ['packing@enicarpharma.com', 'exports@enicarpharma.com', 'swaralisave@enicarpharma.com']
REVIEW_ONLY = ['swaralisave@enicarpharma.com']


def has_findings(text):
    return ('NOT yet in the alias list' in text) or ('appear to be mistyped' in text)


def main():
    if not os.path.exists(REPORT):
        print('No report file to send.'); return 0
    body = open(REPORT, encoding='utf-8').read()

    if not has_findings(body):
        print('Nothing to fix this week — no email sent.'); return 0

    first_time = not os.path.exists(FLAG)
    recipients = REVIEW_ONLY if first_time else TEAM

    stamp = datetime.now().strftime('%d %b %Y')
    msg = EmailMessage()
    msg['From'] = SENDER
    msg['To']   = ', '.join(recipients)
    if first_time:
        msg['Subject'] = f'[REVIEW] Enicar Weekly Data Review — {stamp}'
        intro = ('This is the FIRST weekly data-review email, sent to you only for review.\n'
                 'From next week it will auto-send to: ' + ', '.join(TEAM) + '\n'
                 '(If anything looks wrong, tell the dashboard admin before next week.)\n\n')
    else:
        msg['Subject'] = f'Enicar Weekly Data Review — {stamp}'
        intro = ('Weekly data review below. Please correct the flagged batch numbers / '
                 'customer names in the Google Sheet.\n\n')
    msg.set_content(intro + body)

    # attach the .txt as well
    msg.add_attachment(body.encode('utf-8'), maintype='text', subtype='plain',
                       filename='Weekly_Data_Review.txt')

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ctx) as s:
        s.login(SENDER, APP_PW)
        s.send_message(msg)

    if first_time:
        open(FLAG, 'w').write(datetime.now().isoformat())
        print(f'✓ REVIEW copy sent to {recipients[0]} (auto-send to team starts next run).')
    else:
        print(f'✓ Sent to: {", ".join(recipients)}')
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        print(f'✗ Email send failed: {e}')
        sys.exit(1)
