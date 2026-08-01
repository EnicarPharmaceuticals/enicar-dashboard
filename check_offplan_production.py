#!/usr/bin/env python3
"""
check_offplan_production.py — alert when production happens OUTSIDE the plan
─────────────────────────────────────────────────────────────────────────────
Requested by management (1 Aug 2026):
  1. IMMEDIATE alert — the moment a product that is NOT on the monthly plan is
     taken into production (RM dispensed or filling started in the plan month),
     email right away.
  2. DIGEST — on the 15th and on the last day of the month, a formatted
     summary of everything manufactured off-plan so far (sent even when the
     answer is "nothing — plan fully followed").

Recipients: nimishpatil@enicarpharma.com, swaralisave@enicarpharma.com
Runs every cloud build (15 min). State lives under "_offplan_alerted" /
"_offplan_digest_sent" in product_mismatch_state.json (committed by the
workflow, so it survives between runs).

Rules mirror the dashboard's off-plan logic exactly:
  • July carry-over is never off-plan (dispensed OR first-filled before the
    plan month).
  • A batch is ON plan if its product fuzzy-matches a plan line, or the store
    wrote its batch number into the plan tab's BATCH NO. column.
  • Opening-stock baseline batches are ignored.
OFFPLAN_EMAILS=0 disables sending (report-only). --test prints, never sends.
"""

import os, sys, json, re, calendar
from datetime import date, datetime

HERE  = os.path.dirname(os.path.abspath(__file__))
ROOT  = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
XLSX  = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')
PLAN_JSON = os.path.join(HERE, 'plan_aug_2026.json')
BASELINE  = os.path.join(HERE, 'batch_baseline.json')
STATE = os.path.join(HERE, 'product_mismatch_state.json')
ALERT_KEY  = '_offplan_alerted'
DIGEST_KEY = '_offplan_digest_sent'
TO = ['nimishpatil@enicarpharma.com', 'swaralisave@enicarpharma.com']

PLAN_MONTH_START = date(2026, 8, 1)
PLAN_LABEL = 'AUG 2026'


def bk(b):  return re.sub(r'\s+', '', str(b)).upper()
def pc(s):  return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def load_plan_products_and_batches():
    """Returns (set of canon product names on the plan, set of batch keys the
    store wrote into the plan tab)."""
    prods, batches = set(), set()
    import pandas as pd
    try:
        tab = next((s for s in pd.ExcelFile(XLSX).sheet_names
                    if 'PLAN' in s.upper() and 'DISPENS' not in s.upper()), None)
        if tab:
            df = pd.read_excel(XLSX, sheet_name=tab, header=0)
            df.columns = [' '.join(str(c).split()).upper() for c in df.columns]
            cp = next((c for c in df.columns if 'PRODUCT' in c), None)
            cb = next((c for c in df.columns if 'BATCH' in c), None)
            for _, r in df.iterrows():
                p = r.get(cp)
                if pd.notna(p) and str(p).strip():
                    prods.add(pc(p))
                if cb is not None:
                    b = r.get(cb)
                    if pd.notna(b) and str(b).strip():
                        for x in re.split(r'[,/;]+', str(b)):
                            if x.strip():
                                batches.add(bk(x))
            if prods:
                return prods, batches
    except Exception as e:
        print(f'  (plan tab unreadable: {e})')
    try:
        pj = json.load(open(PLAN_JSON))
        prods = {pc(i['product']) for i in pj.get('items', [])}
    except Exception:
        pass
    return prods, batches


def on_plan(product, key, plan_prods, plan_batches):
    if key in plan_batches:
        return True
    c = pc(product)
    if len(c) < 4:
        return False
    return any(len(pp) >= 4 and (c in pp or pp in c) for pp in plan_prods)


def find_offplan():
    """Batches taken into production this month whose product is not planned."""
    import pandas as pd
    plan_prods, plan_batches = load_plan_products_and_batches()
    if not plan_prods:
        print('  No plan found — skipping.')
        return None
    try:
        baseline = {bk(b) for b in json.load(open(BASELINE)).get('batches', [])}
    except Exception:
        baseline = set()

    def load(sheet):
        df = pd.read_excel(XLSX, sheet_name=sheet, header=3)
        df.columns = [' '.join(str(c).split()) for c in df.columns]
        return df

    fill = load('➕ Filling Log')
    rm = load('➕ RM Dispensing Log')

    # first filling date + qty per batch
    fmap = {}
    for _, r in fill.iterrows():
        b = r.get('Batch No.')
        if pd.isna(b) or not str(b).strip():
            continue
        d = pd.to_datetime(r.get('Date'), format='mixed', dayfirst=True, errors='coerce')
        k = bk(b)
        e = fmap.setdefault(k, {'first': None, 'qty': 0.0, 'product': str(r.get('Product Name') or '').strip()})
        e['qty'] += float(pd.to_numeric(r.get('Qty Filled (Units)'), errors='coerce') or 0)
        if pd.notna(d) and (e['first'] is None or d.date() < e['first']):
            e['first'] = d.date()

    out = {}
    # RM dispensed in the plan month
    for _, r in rm.iterrows():
        b = r.get('BATCH NUMBER')
        if pd.isna(b) or str(b).strip() in ('', '-'):
            continue
        k = bk(b)
        if k in baseline or 'TLB' in k:
            continue
        d = pd.to_datetime(r.get('DISPENSING DATE'), errors='coerce')
        if pd.isna(d) or d.date() < PLAN_MONTH_START:
            continue                                    # July plan work
        f = fmap.get(k)
        if f and f['first'] and f['first'] < PLAN_MONTH_START:
            continue                                    # filling began in July
        prod = str(r.get('NAME OF THE PRODUCT') or '').strip()
        if on_plan(prod, k, plan_prods, plan_batches):
            continue
        out[k] = {'batch': str(b).strip(), 'product': prod,
                  'company': str(r.get('CUSTOMER') or '').strip(),
                  'stage': 'RM dispensed', 'date': d.date().isoformat(),
                  'qty': float(pd.to_numeric(r.get('BATCH SIZE'), errors='coerce') or 0)}
    # filled in the plan month without matching plan (covers missing-RM cases)
    for k, f in fmap.items():
        if k in out or k in baseline or not f['first'] or f['first'] < PLAN_MONTH_START:
            continue
        if on_plan(f['product'], k, plan_prods, plan_batches):
            continue
        out[k] = {'batch': k, 'product': f['product'], 'company': '',
                  'stage': 'Filling started', 'date': f['first'].isoformat(),
                  'qty': f['qty']}
    return out


def html_table(rows):
    tr = ''.join(
        f'<tr><td style="padding:6px 10px;border:1px solid #B0BEC5;font-weight:600">{r["batch"]}</td>'
        f'<td style="padding:6px 10px;border:1px solid #B0BEC5">{r["product"] or "—"}</td>'
        f'<td style="padding:6px 10px;border:1px solid #B0BEC5">{r["company"] or "—"}</td>'
        f'<td style="padding:6px 10px;border:1px solid #B0BEC5">{r["stage"]}</td>'
        f'<td style="padding:6px 10px;border:1px solid #B0BEC5">{r["date"]}</td>'
        f'<td style="padding:6px 10px;border:1px solid #B0BEC5;text-align:right">{r["qty"]:,.0f}</td></tr>'
        for r in rows)
    return (f'<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px">'
            f'<thead><tr style="background:#004D40;color:#fff">'
            f'<th style="padding:7px 10px;border:1px solid #004D40">BATCH</th>'
            f'<th style="padding:7px 10px;border:1px solid #004D40">PRODUCT</th>'
            f'<th style="padding:7px 10px;border:1px solid #004D40">COMPANY (RM)</th>'
            f'<th style="padding:7px 10px;border:1px solid #004D40">STAGE</th>'
            f'<th style="padding:7px 10px;border:1px solid #004D40">DATE</th>'
            f'<th style="padding:7px 10px;border:1px solid #004D40">QTY</th>'
            f'</tr></thead><tbody>{tr}</tbody></table>')


def send(subject, text, html, test=False):
    if test or os.environ.get('OFFPLAN_EMAILS', '1').strip().lower() in ('0', 'false', 'no'):
        print(f'  [not sent] {subject}\n{text}\n')
        return True
    sender = os.environ.get('GMAIL_SENDER', '').strip()
    pw = os.environ.get('GMAIL_APP_PASSWORD', '').strip()
    if not (sender and pw):
        try:
            sys.path.insert(0, HERE)
            import email_config as cfg
            sender, pw = cfg.MY_EMAIL, cfg.APP_PASSWORD
        except Exception:
            print('  ✗ no email credentials'); return False
    import smtplib
    from email.message import EmailMessage
    msg = EmailMessage()
    msg['From'] = sender; msg['To'] = ', '.join(TO); msg['Subject'] = subject
    msg.set_content(text)
    msg.add_alternative(f'<html><body style="font-family:Arial,sans-serif">{html}</body></html>',
                        subtype='html')
    try:
        s = smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30)
        s.login(sender, pw); s.send_message(msg); s.quit()
        print(f'  ✉ sent: {subject}')
        return True
    except Exception as e:
        print(f'  ✗ send failed: {e}'); return False


def main():
    test = '--test' in sys.argv
    off = find_offplan()
    if off is None:
        return 0
    print(f'── Off-plan production check · {len(off)} off-plan batch(es) this month ──')
    try:
        state = json.load(open(STATE))
    except Exception:
        state = {}
    alerted = state.get(ALERT_KEY, {})
    today = date.today()

    # 1. IMMEDIATE alerts for newly-seen off-plan batches
    new = {k: v for k, v in off.items() if k not in alerted}
    if new:
        rows = sorted(new.values(), key=lambda r: r['date'])
        text = ('Sir/Madam,\n\nThe following product(s) have been taken into production '
                f'but are NOT on the {PLAN_LABEL} production plan:\n\n'
                + '\n'.join(f"  • {r['batch']} — {r['product']} ({r['company'] or 'company n/a'}) — "
                            f"{r['stage']} on {r['date']}, qty {r['qty']:,.0f}" for r in rows)
                + '\n\nPlease review whether this is approved additional production or a '
                  'plan/naming gap.\n\n— Enicar Dashboard (automatic)')
        html = (f'<p>Sir/Madam,</p><p>The following product(s) have been taken into production '
                f'but are <strong>NOT on the {PLAN_LABEL} production plan</strong>:</p>'
                + html_table(rows)
                + '<p>Please review whether this is approved additional production or a plan/naming gap.</p>'
                  '<p style="color:#90A4AE">— Enicar Dashboard (automatic notification)</p>')
        if send(f'⚠ Off-plan production started — {len(new)} product(s) not on the {PLAN_LABEL} plan',
                text, html, test):
            for k, v in new.items():
                alerted[k] = {'date': today.isoformat(), **v}
            state[ALERT_KEY] = alerted
    else:
        print('  No new off-plan production.')

    # 2. DIGEST on the 15th and the last day of the month
    last_day = calendar.monthrange(today.year, today.month)[1]
    if today.day in (15, last_day) or '--force-digest' in sys.argv:
        if state.get(DIGEST_KEY) != today.isoformat() or '--force-digest' in sys.argv:
            allrows = sorted({**{k: v for k, v in alerted.items() if not k.startswith('_')},
                              **off}.values(), key=lambda r: r.get('date', ''))
            allrows = [r for r in allrows if 'batch' in r]
            period = 'mid-month' if today.day == 15 else 'end-of-month'
            if allrows:
                text = (f'Sir/Madam,\n\n{period.title()} summary — products manufactured OUTSIDE '
                        f'the {PLAN_LABEL} production plan ({len(allrows)} batch(es)):\n\n'
                        + '\n'.join(f"  • {r['batch']} — {r['product']} ({r.get('company') or 'n/a'}) — "
                                    f"{r.get('stage','')} {r.get('date','')}, qty {r.get('qty',0):,.0f}"
                                    for r in allrows)
                        + '\n\n— Enicar Dashboard (automatic)')
                html = (f'<p>Sir/Madam,</p><p><strong>{period.title()} summary</strong> — products '
                        f'manufactured <strong>outside the {PLAN_LABEL} production plan</strong> '
                        f'({len(allrows)} batches):</p>' + html_table(allrows)
                        + '<p style="color:#90A4AE">— Enicar Dashboard (automatic report)</p>')
            else:
                text = (f'Sir/Madam,\n\n{period.title()} summary: NO off-plan production — every '
                        f'batch manufactured so far this month is on the {PLAN_LABEL} plan. ✅\n\n'
                        '— Enicar Dashboard (automatic)')
                html = (f'<p>Sir/Madam,</p><p><strong>{period.title()} summary:</strong> '
                        f'<span style="color:#1B5E20;font-weight:700">No off-plan production</span> — '
                        f'every batch manufactured so far this month is on the {PLAN_LABEL} plan. ✅</p>'
                        '<p style="color:#90A4AE">— Enicar Dashboard (automatic report)</p>')
            if send(f'{PLAN_LABEL} plan compliance — {period} off-plan production report',
                    text, html, test):
                state[DIGEST_KEY] = today.isoformat()

    json.dump(state, open(STATE, 'w'), indent=2, sort_keys=True)
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        print(f'  (off-plan check unavailable: {e})')
        sys.exit(0)
