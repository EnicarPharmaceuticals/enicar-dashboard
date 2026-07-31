#!/usr/bin/env python3
"""
check_plan_changes.py — watch the monthly production plan, email store on change
────────────────────────────────────────────────────────────────────────────────
The Plant Head edits the plan in the sheet's "AUG PLAN" tab (any tab whose name
contains PLAN). Every cloud build (15 min) this script fingerprints the plan
and compares it with the snapshot stored under the "_plan_snapshot" key of
product_mismatch_state.json (that file is already committed by the workflow,
which is how the snapshot survives between runs).

On a difference it emails the CHANGES (added / removed / quantity-changed
items) to the store team, so RM dispensing always knows what to prepare —
this notification flow was explicitly requested by the Director (31 Jul 2026).

Rules:
  • First run (no snapshot yet) → save silently, never email.
  • PLAN_EMAILS=0 in the environment disables sending (report-only).
  • Credentials: GMAIL_SENDER + GMAIL_APP_PASSWORD env vars (cloud), falling
    back to email_config.py (local Mac).
Recipients: store@enicarpharma.com
"""

import os, sys, json, re
from datetime import datetime

HERE  = os.path.dirname(os.path.abspath(__file__))
ROOT  = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
XLSX  = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')
PLAN_JSON = os.path.join(HERE, 'plan_aug_2026.json')
STATE = os.path.join(HERE, 'product_mismatch_state.json')
SNAP_KEY = '_plan_snapshot'
TO = ['store@enicarpharma.com']


def load_plan_rows():
    """Return (rows, source): rows = list of dicts with product/party/units/pack/priority."""
    try:
        import pandas as pd
        tab = next((s for s in pd.ExcelFile(XLSX).sheet_names
                    if 'PLAN' in s.upper() and 'DISPENS' not in s.upper()), None)
        if tab:
            df = pd.read_excel(XLSX, sheet_name=tab, header=0)
            df.columns = [' '.join(str(c).split()).upper() for c in df.columns]
            def col(*names):
                return next((c for c in df.columns if any(n in c for n in names)), None)
            cp, cparty = col('PRODUCT'), col('PARTY', 'CUSTOMER')
            cq, cpack, cprio = col('QTY', 'PLANNED'), col('PACK'), col('PRIORITY')
            cstat, cdate = col('RM STATUS', 'STATUS'), col('RM DATE', 'DISPENSE DATE')
            cbatch = col('BATCH')
            def _num(v):
                x = pd.to_numeric(v, errors='coerce')
                return 0.0 if pd.isna(x) else float(x)
            def _txt(v):
                # NaN is truthy in Python — never use `v or ''` on sheet cells
                return '' if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v).strip()
            rows = []
            for _, r in df.iterrows():
                prod = r.get(cp)
                if pd.isna(prod) or not str(prod).strip():
                    continue
                pr = _num(r.get(cprio))
                rows.append({'product': str(prod).strip(),
                             'party': _txt(r.get(cparty)),
                             'units': _num(r.get(cq)),
                             'pack': _txt(r.get(cpack)),
                             'priority': str(int(pr)) if pr else '',
                             'status': _txt(r.get(cstat)) if cstat else '',
                             'date': _txt(r.get(cdate)) if cdate else '',
                             'batch': _txt(r.get(cbatch)) if cbatch else ''})
            if rows:
                return rows, f'sheet tab "{tab}"'
    except Exception as e:
        print(f'  (plan tab unreadable: {e})')
    try:
        pj = json.load(open(PLAN_JSON))
        rows = [{'product': i['product'], 'party': i['party'],
                 'units': float(i.get('planned_units') or 0),
                 'pack': str(i.get('pack') or ''), 'priority': str(i.get('priority') or ''),
                 'status': '', 'date': '', 'batch': ''}
                for i in pj.get('items', [])]
        return rows, 'bundled plan JSON'
    except Exception:
        return [], 'no plan'


def key_of(r):
    """Identity of a plan line: product + party + pack (case/punct-insensitive)."""
    c = lambda s: re.sub(r'[^a-z0-9]', '', str(s).lower())
    return f"{c(r['product'])}|{c(r['party'])}|{c(r['pack'])}"


def fingerprint(rows):
    """key → summary used both for change detection and for the diff email."""
    fp = {}
    for r in rows:
        k = key_of(r)
        e = fp.setdefault(k, {'product': r['product'], 'party': r['party'],
                              'pack': r['pack'], 'units': 0.0, 'priority': r['priority'],
                              'status': r.get('status', ''), 'date': r.get('date', ''),
                              'batch': r.get('batch', '')})
        e['units'] += r['units']
        for fld in ('status', 'date', 'batch'):
            if r.get(fld) and not e.get(fld):
                e[fld] = r[fld]
    return fp


def main():
    send_enabled = os.environ.get('PLAN_EMAILS', '1').strip().lower() not in ('0', 'false', 'no')
    rows, source = load_plan_rows()
    print(f'── Plan-change check · {len(rows)} plan lines from {source} ──')
    if not rows:
        return 0

    try:
        state = json.load(open(STATE))
    except Exception:
        state = {}
    new_fp = fingerprint(rows)
    old = state.get(SNAP_KEY)

    if old is None or '--init' in sys.argv:
        state[SNAP_KEY] = {'fp': new_fp, 'ts': datetime.now().isoformat(timespec='seconds'),
                           'source': source}
        json.dump(state, open(STATE, 'w'), indent=2, sort_keys=True)
        print('  Snapshot initialised silently (no email on first run).')
        return 0

    old_fp = old.get('fp', {})
    added   = [new_fp[k] for k in new_fp if k not in old_fp]
    removed = [old_fp[k] for k in old_fp if k not in new_fp]
    def _differs(a, b):
        if abs(b['units'] - a['units']) > 0.5:
            return True
        for fld in ('priority', 'status', 'date', 'batch'):
            if str(b.get(fld, '')) != str(a.get(fld, '')):
                return True
        return False
    changed = [(old_fp[k], new_fp[k]) for k in new_fp if k in old_fp and _differs(old_fp[k], new_fp[k])]

    if not (added or removed or changed):
        print('  No plan changes.')
        return 0

    n_changes = len(added) + len(removed) + len(changed)
    print(f'  PLAN CHANGED: +{len(added)} added, -{len(removed)} removed, ~{len(changed)} modified')

    lines = ['Store team,', '',
             f'The production plan was updated ({source}). Changes below — please adjust '
             f'RM dispensing preparation accordingly. Items marked "Dispense next" are needed first.', '']
    if added:
        lines.append(f'ADDED TO PLAN ({len(added)}):')
        for e in added:
            lines.append(f"   + {e['product']}  |  {e['party']}  |  {e['units']:,.0f} x {e['pack']}"
                         + (f"  |  Priority {e['priority']}" if e['priority'] else ''))
        lines.append('')
    if removed:
        lines.append(f'REMOVED FROM PLAN ({len(removed)}):')
        for e in removed:
            lines.append(f"   - {e['product']}  |  {e['party']}  |  {e['units']:,.0f} x {e['pack']}")
        lines.append('')
    if changed:
        lines.append(f'CHANGED ({len(changed)}):')
        for o, nw in changed:
            bits = []
            if abs(nw['units'] - o['units']) > 0.5:
                bits.append(f"qty {o['units']:,.0f} -> {nw['units']:,.0f}")
            if str(nw['priority']) != str(o['priority']):
                bits.append(f"priority {o['priority'] or '-'} -> {nw['priority'] or '-'}")
            if str(nw.get('status', '')) != str(o.get('status', '')):
                bits.append(f"STATUS {o.get('status') or '-'} -> {nw.get('status') or '-'}")
            if str(nw.get('date', '')) != str(o.get('date', '')):
                bits.append(f"date {o.get('date') or '-'} -> {nw.get('date') or '-'}")
            if str(nw.get('batch', '')) != str(o.get('batch', '')):
                bits.append(f"batch {o.get('batch') or '-'} -> {nw.get('batch') or '-'}")
            lines.append(f"   ~ {nw['product']}  |  {nw['party']}  |  " + ', '.join(bits))
        lines.append('')
    lines += ['This is an automatic notification from the production dashboard '
              '(sent whenever the Plant Head updates the plan).', '']
    body = '\n'.join(lines)

    if not send_enabled:
        print('  PLAN_EMAILS disabled — change detected but NOT emailed:\n')
        print(body)
    else:
        sender = os.environ.get('GMAIL_SENDER', '').strip()
        pw     = os.environ.get('GMAIL_APP_PASSWORD', '').strip()
        if not (sender and pw):
            try:
                sys.path.insert(0, HERE)
                import email_config as cfg
                sender, pw = cfg.MY_EMAIL, cfg.APP_PASSWORD
            except Exception:
                sender = pw = ''
        if not (sender and pw):
            print('  ✗ No email credentials — change detected but not emailed.')
        else:
            import smtplib
            from email.message import EmailMessage
            msg = EmailMessage()
            msg['From'] = sender
            msg['To'] = ', '.join(TO)
            msg['Subject'] = f'Production plan updated — {n_changes} change(s), please review RM dispensing'
            msg.set_content(body)
            try:
                s = smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30)
                s.login(sender, pw)
                s.send_message(msg)
                s.quit()
                print(f'  ✉ Change notification sent to {", ".join(TO)}')
            except Exception as e:
                print(f'  ✗ Email send failed: {e}')
                return 0   # snapshot still updated below; next change emails again

    state[SNAP_KEY] = {'fp': new_fp, 'ts': datetime.now().isoformat(timespec='seconds'),
                       'source': source}
    json.dump(state, open(STATE, 'w'), indent=2, sort_keys=True)
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        print(f'  (plan-change check unavailable: {e})')
        sys.exit(0)
