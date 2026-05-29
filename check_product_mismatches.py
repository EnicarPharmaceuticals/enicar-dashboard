#!/usr/bin/env python3
"""
check_product_mismatches.py
─────────────────────────────────────────────────────────────────
Catches batches where the SAME batch number was logged with
DIFFERENT product spellings across Filling / Packing / Dispatch.
Same batch number = same product, so one spelling is wrong.

Behaviour:
  • On first detection → email the team (humanized list of fixes).
  • Re-check on every run (cloud runs this every 15 min).
  • If an issue is still present 3+ days after the last email
    about it → send a reminder.
  • When the issue disappears from the sheet (fixed) → drop it
    silently from the tracking state.

State is persisted in product_mismatch_state.json (committed to
the repo by the workflow's existing commit step, since we git-add
the file from here).

Reads creds from env (cloud) or email_config.py (local).
Reuses GMAIL_SENDER / GMAIL_APP_PASSWORD secrets — no new ones.
"""

import os, sys, re, json, smtplib, ssl, subprocess
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from collections import defaultdict

import pandas as pd

HERE  = os.path.dirname(os.path.abspath(__file__))
ROOT  = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
XLSX  = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')
STATE = os.path.join(HERE, 'product_mismatch_state.json')

REMIND_AFTER_DAYS = 3   # initial email, then every 3 days while uncorrected

TEAM = ['packing@enicarpharma.com',
        'exports@enicarpharma.com',
        'swaralisave@enicarpharma.com']

# ─── creds ──────────────────────────────────────────────────────
SENDER = os.environ.get('GMAIL_SENDER', '').strip()
APP_PW = os.environ.get('GMAIL_APP_PASSWORD', '').strip()
if not (SENDER and APP_PW):
    try:
        sys.path.insert(0, HERE)
        import email_config as cfg
        SENDER = SENDER or getattr(cfg, 'MY_EMAIL', '')
        APP_PW = APP_PW or getattr(cfg, 'APP_PASSWORD', '')
    except Exception:
        pass


# ─── mismatch detection ─────────────────────────────────────────
def _bkey(b):  return re.sub(r'\s+', '', str(b)).upper()
def _pkey(s):  return ' '.join(str(s).strip().lower().split())
def _root(s):  return re.sub(r'[^a-z0-9]', '', _pkey(s))[:6]    # for typo-vs-different-product heuristic
def _canon(s): return re.sub(r'[^a-z0-9]', '', _pkey(s))        # strict: ignore ALL whitespace / punctuation / case


def find_mismatches():
    """Return list of issue dicts and a stable fingerprint for state."""
    sources = [
        ('Filling',  '➕ Filling Log',  'B:J', 6, 2),
        ('Packing',  '➕ Packing Log',  'B:N', 5, 2),
        ('Dispatch', '➕ Dispatch Log', 'B:I', 5, 1),
    ]
    by_batch = defaultdict(lambda: defaultdict(lambda: {'rep': None, 'logs': set(), 'rows': 0}))
    raw_batch = {}
    for log_name, sheet, usecols, bidx, pidx in sources:
        try:
            df = pd.read_excel(XLSX, sheet_name=sheet, header=3, usecols=usecols)
        except Exception:
            continue
        for _, r in df.iterrows():
            b = r.iloc[bidx]; p = r.iloc[pidx]
            if pd.isna(b) or pd.isna(p) or not str(p).strip():
                continue
            k = _bkey(b)
            raw_batch.setdefault(k, str(b).strip())
            raw_prod = str(p).strip()
            # Group by _canon so spellings that differ ONLY by whitespace, case
            # or punctuation collapse into one — those are silent normalizations,
            # not worth emailing the team about.
            e = by_batch[k][_canon(raw_prod)]
            if e['rep'] is None: e['rep'] = raw_prod
            e['logs'].add(log_name); e['rows'] += 1

    issues = []
    for k, spellings in by_batch.items():
        if len(spellings) < 2:
            continue
        ranked = sorted(spellings.values(),
                        key=lambda e: (len(e['logs']), e['rows']), reverse=True)
        correct = ranked[0]
        for wrong in ranked[1:]:
            # Likely-different-products vs likely-typo — both worth flagging,
            # but we label the seriously different ones distinctly.
            severity = 'typo' if _root(correct['rep']) == _root(wrong['rep']) else 'different_product'
            fp = f"{k}|{_pkey(correct['rep'])}|{_pkey(wrong['rep'])}"
            issues.append({
                'fingerprint':  fp,
                'batch':        raw_batch[k],
                'correct':      correct['rep'],
                'correct_logs': sorted(correct['logs']),
                'wrong':        wrong['rep'],
                'wrong_logs':   sorted(wrong['logs']),
                'severity':     severity,
            })
    issues.sort(key=lambda m: (m['severity'] != 'different_product', m['batch']))
    return issues


# ─── state ──────────────────────────────────────────────────────
def load_state():
    if not os.path.exists(STATE): return {}
    try:    return json.load(open(STATE))
    except Exception: return {}

def save_state(state):
    json.dump(state, open(STATE, 'w'), indent=2, sort_keys=True)
    # Stage the file so the workflow's existing commit step picks it up.
    # Local runs (no git) are fine — error is swallowed.
    try:
        subprocess.run(['git', 'add', STATE], cwd=HERE,
                       capture_output=True, check=False, timeout=10)
    except Exception:
        pass


# ─── email body ─────────────────────────────────────────────────
def build_body(to_send, is_reminder):
    stamp = datetime.now().strftime('%d %B %Y')
    typos    = [i for i in to_send if i['severity'] == 'typo']
    seriouse = [i for i in to_send if i['severity'] == 'different_product']

    lead = (
        "Hi team,\n\n"
        "Quick fix request — we've spotted batches where the product name is "
        "written differently in different logs. Since the batch number is the same, "
        "it's really the same product, so one of the spellings needs correcting in "
        "the Google Sheet.\n\n"
    ) if not is_reminder else (
        "Hi team,\n\n"
        "Friendly reminder — the items below were flagged a few days ago and are "
        "still uncorrected in the spreadsheet. Please update them when you get a moment.\n\n"
    )

    body = lead + "──────────────────────────────────────────────────\n\n"

    if typos:
        body += "SPELLING FIXES — please update in the Google Sheet:\n\n"
        for i in typos:
            body += (
                f"  • Batch {i['batch']}:\n"
                f"      In the {', '.join(i['wrong_logs'])} log → "
                f'change "{i["wrong"]}" to "{i["correct"]}"\n'
                f"      (already entered as \"{i['correct']}\" in {', '.join(i['correct_logs'])})\n\n"
            )

    if seriouse:
        body += "──────────────────────────────────────────────────\n\n"
        body += "PLEASE DOUBLE-CHECK — same batch number used for what look like\n"
        body += "different products. Likely one of the rows has the wrong batch number:\n\n"
        for i in seriouse:
            body += (
                f"  • Batch {i['batch']}:\n"
                f'      {", ".join(i["correct_logs"])} log says "{i["correct"]}"\n'
                f'      {", ".join(i["wrong_logs"])} log says "{i["wrong"]}"\n'
                f"      → kindly check which is right and correct the wrong row.\n\n"
            )

    body += (
        "──────────────────────────────────────────────────\n\n"
        "We re-check the sheet automatically — once you fix these in the spreadsheet "
        f"they'll drop off this list. If anything is still pending, a reminder will "
        f"come in {REMIND_AFTER_DAYS} days.\n\n"
        "Thanks for keeping the data clean!\n\n"
        "— Enicar Dashboard (automatic check)"
    )
    return body


# ─── main ───────────────────────────────────────────────────────
def main():
    if not os.path.exists(XLSX):
        return 0   # nothing to do
    if not (SENDER and APP_PW):
        print('check_product_mismatches: no email creds — skipping.')
        return 0

    issues = find_mismatches()
    current_fps = {i['fingerprint']: i for i in issues}
    state = load_state()
    today = date.today()
    today_iso = today.isoformat()

    to_send = []      # issues to email this run
    is_reminder = False

    for fp, issue in current_fps.items():
        last = state.get(fp)
        if not last:
            # First time we're seeing this issue → email today.
            to_send.append(issue)
            state[fp] = today_iso
        else:
            try: last_d = date.fromisoformat(last)
            except Exception: last_d = today
            if (today - last_d) >= timedelta(days=REMIND_AFTER_DAYS):
                to_send.append(issue)
                state[fp] = today_iso
                is_reminder = True

    # Drop fingerprints that are no longer in current data (corrected).
    for fp in list(state):
        if fp not in current_fps:
            del state[fp]

    if to_send:
        stamp = datetime.now().strftime('%d %b %Y')
        msg = EmailMessage()
        msg['From'] = SENDER
        msg['To']   = ', '.join(TEAM)
        prefix = '[Reminder] ' if is_reminder else ''
        plural = 'fix' if len(to_send) == 1 else 'fixes'
        msg['Subject'] = f'{prefix}Enicar — {len(to_send)} product-name {plural} needed in spreadsheet ({stamp})'
        msg.set_content(build_body(to_send, is_reminder))
        try:
            ctx = ssl.create_default_context()
            with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ctx) as s:
                s.login(SENDER, APP_PW)
                s.send_message(msg)
            print(f"✓ Mismatch email sent ({len(to_send)} item{'s' if len(to_send)!=1 else ''}, "
                  f"{'reminder' if is_reminder else 'initial'}) → {', '.join(TEAM)}")
        except Exception as e:
            print(f'✗ Mismatch email send failed: {e}')
            return 1
    else:
        if not issues:
            print('check_product_mismatches: no mismatches — nothing to email.')
        else:
            print(f'check_product_mismatches: {len(issues)} issue(s) tracked, '
                  f'none due for re-email today.')

    save_state(state)
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        print(f'✗ check_product_mismatches: {e}')
        sys.exit(1)
