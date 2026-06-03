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
    """Cross-verify product AND party names across all 4 logs by batch number.

    Rules:
      * Trivial whitespace / case / punctuation differences are ignored (_canon).
      * If a batch is present in RM Dispensing, the RM spelling is the source
        of truth — all other logs must match it.
      * Otherwise, fall back to the 2-of-3-logs rule (most logs wins).
    """
    # (log_name, sheet, usecols, batch_idx, product_idx, party_idx) — indices
    # are 0-based within the usecols slice.
    # RM tab is now 24 columns (RM team uses extra fields). With usecols='B:I'
    # we get: [0]SR.NO  [1]CUSTOMER  [2]BMR DATE  [3]BMR TIME
    #         [4]NAME OF PRODUCT  [5]BATCH NUMBER  [6]BATCH SIZE  [7]UOM
    sources = [
        ('RM Dispensing','➕ RM Dispensing Log','B:I',5,4,1),
        ('Filling',      '➕ Filling Log',      'B:J',6,2,7),
        ('Packing',      '➕ Packing Log',      'B:N',5,2,10),
        ('Dispatch',     '➕ Dispatch Log',     'B:I',5,1,6),
    ]
    # batches[batch_key]['product'|'party'][canon] -> {rep, logs, rows}
    batches = defaultdict(lambda: {
        'product': defaultdict(lambda: {'rep': None, 'logs': set(), 'rows': 0}),
        'party':   defaultdict(lambda: {'rep': None, 'logs': set(), 'rows': 0}),
        'raw':     None,
    })
    for log_name, sheet, usecols, bidx, pidx, partyidx in sources:
        try:
            df = pd.read_excel(XLSX, sheet_name=sheet, header=3, usecols=usecols)
        except Exception:
            continue
        for _, r in df.iterrows():
            b = r.iloc[bidx]
            if pd.isna(b) or not str(b).strip(): continue
            k = _bkey(b)
            if batches[k]['raw'] is None: batches[k]['raw'] = str(b).strip()
            for field, idx in [('product', pidx), ('party', partyidx)]:
                v = r.iloc[idx]
                if pd.isna(v) or not str(v).strip(): continue
                val = str(v).strip()
                e = batches[k][field][_canon(val)]
                if e['rep'] is None: e['rep'] = val
                e['logs'].add(log_name); e['rows'] += 1

    issues = []
    for k, info in batches.items():
        for field in ('product', 'party'):
            spellings = info[field]
            if len(spellings) < 2: continue

            # RM wins if it has any spelling for this batch+field.
            rm_entry = next((s for s in spellings.values() if 'RM Dispensing' in s['logs']), None)
            if rm_entry is not None:
                correct = rm_entry
                wrong_list = [s for s in spellings.values() if s is not correct]
                from_rm = True
            else:
                ranked = sorted(spellings.values(),
                                key=lambda e: (len(e['logs']), e['rows']), reverse=True)
                correct = ranked[0]
                wrong_list = ranked[1:]
                from_rm = False

            for wrong in wrong_list:
                severity = 'typo' if _root(correct['rep']) == _root(wrong['rep']) else 'different_value'
                fp = f"{k}|{field}|{_canon(correct['rep'])}|{_canon(wrong['rep'])}"
                issues.append({
                    'fingerprint':  fp,
                    'batch':        info['raw'],
                    'field':        field,
                    'correct':      correct['rep'],
                    'correct_logs': sorted(correct['logs']),
                    'wrong':        wrong['rep'],
                    'wrong_logs':   sorted(wrong['logs']),
                    'severity':     severity,
                    'from_rm':      from_rm,
                })
    # Sort: serious issues first, then by batch, then field
    issues.sort(key=lambda m: (m['severity'] != 'different_value', m['batch'], m['field']))
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
def _why(i):
    """One-liner explaining why the 'correct' is correct."""
    if i.get('from_rm'):
        return f"RM Dispensing has \"{i['correct']}\" for this batch"
    return f"already entered as \"{i['correct']}\" in {', '.join(i['correct_logs'])}"

def build_body(to_send, is_reminder):
    stamp = datetime.now().strftime('%d %B %Y')

    # Split into product fixes vs party fixes vs serious (different-product) cases
    prod_fixes  = [i for i in to_send if i['field']=='product' and i['severity']=='typo']
    party_fixes = [i for i in to_send if i['field']=='party'   and i['severity']=='typo']
    serious     = [i for i in to_send if i['severity']=='different_value']

    lead = (
        "Hi team,\n\n"
        "Quick fix request — we've spotted batches where the product or customer "
        "name is written differently in different logs. Since the batch number is "
        "the same, the entries should match. RM Dispensing is the source of truth, "
        "so please update Filling / Packing / Dispatch to match the RM spelling.\n\n"
    ) if not is_reminder else (
        "Hi team,\n\n"
        "Friendly reminder — the items below were flagged a few days ago and are "
        "still uncorrected in the spreadsheet. Please update them when you can.\n\n"
    )

    body = lead + "──────────────────────────────────────────────────\n\n"

    if prod_fixes:
        body += f"PRODUCT-NAME FIXES  ({len(prod_fixes)}):\n\n"
        for i in prod_fixes:
            body += (
                f"  • Batch {i['batch']}:\n"
                f"      In the {', '.join(i['wrong_logs'])} log → "
                f'change "{i["wrong"]}" to "{i["correct"]}"\n'
                f"      ({_why(i)})\n\n"
            )

    if party_fixes:
        body += "──────────────────────────────────────────────────\n\n"
        body += f"CUSTOMER-NAME FIXES  ({len(party_fixes)}):\n\n"
        for i in party_fixes:
            body += (
                f"  • Batch {i['batch']}:\n"
                f"      In the {', '.join(i['wrong_logs'])} log → "
                f'change "{i["wrong"]}" to "{i["correct"]}"\n'
                f"      ({_why(i)})\n\n"
            )

    if serious:
        body += "──────────────────────────────────────────────────\n\n"
        body += f"PLEASE DOUBLE-CHECK  ({len(serious)}) — same batch number used for what look like\n"
        body += "different products/customers. Likely one of the rows has the wrong batch number:\n\n"
        for i in serious:
            label = 'product' if i['field']=='product' else 'customer'
            body += (
                f"  • Batch {i['batch']}  ({label}):\n"
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
        msg['Subject'] = f'{prefix}Enicar — {len(to_send)} spreadsheet {plural} (product/customer names) ({stamp})'
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
