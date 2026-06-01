#!/usr/bin/env python3
"""
check_store_mismatches.py
─────────────────────────────────────────────────────────────────
Catches discrepancies in the RM Dispensing Log that the STORE team
needs to fix. Looks at three things, all keyed by batch number:

  1) BATCH NUMBER  — RM batch is one tiny edit away from a Filling
                     batch (treat Filling as accurate, RM is wrong).
  2) PRODUCT NAME  — for a given batch, RM's product name disagrees
                     with the majority of Filling/Packing/Dispatch.
  3) CUSTOMER NAME — same idea, for the customer column.

Trivial whitespace/case/punctuation differences are silently ignored.

Behaviour:
  • First time an issue is detected → email store@enicarpharma.com.
  • Re-check every 15 min (runs after each cloud_build).
  • If an issue is still present 2+ days later → reminder email.
  • When the issue disappears from the sheet → silently dropped.

Reads creds from env (cloud) or email_config.py (local).
"""

import os, sys, re, json, smtplib, ssl, subprocess
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from collections import defaultdict, Counter
import pandas as pd

HERE  = os.path.dirname(os.path.abspath(__file__))
ROOT  = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
XLSX  = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')
STATE = os.path.join(HERE, 'store_mismatch_state.json')

REMIND_AFTER_DAYS = 3
TEAM = ['store@enicarpharma.com']

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


# ─── normalisation ──────────────────────────────────────────────
def bk(b):    return re.sub(r'\s+', '', str(b)).upper()
def fkey(b):  return re.sub(r'[^A-Z0-9]', '', str(b).upper())     # strip ALL punctuation
def pkey(s):  return ' '.join(str(s).strip().lower().split())
def canon(s): return re.sub(r'[^a-z0-9]', '', pkey(s))            # for comparing names
def root(s):  return canon(s)[:6]

def is_one_indel(a, b):
    if abs(len(a)-len(b)) != 1: return False
    s, l = (a, b) if len(a) < len(b) else (b, a)
    i = j = 0; skipped = False
    while i < len(s) and j < len(l):
        if s[i] == l[j]: i += 1; j += 1
        elif skipped:    return False
        else:            skipped = True; j += 1
    return True

def one_sub(a, b):
    return len(a) == len(b) and sum(1 for x, y in zip(a, b) if x != y) == 1


# ─── detection ──────────────────────────────────────────────────
def find_store_issues():
    """Returns a list of issue dicts the store team should fix in the RM sheet."""
    sources = [
        ('Filling',      '➕ Filling Log',       'B:J', 6, 2, 7),
        ('Packing',      '➕ Packing Log',       'B:N', 5, 2, 10),
        ('Dispatch',     '➕ Dispatch Log',      'B:I', 5, 1, 6),
    ]
    # batch (bk) → for each downstream log: {raw_batch, products{canon→raw}, parties{canon→raw}}
    downstream = defaultdict(lambda: {'raw': None, 'products': Counter(), 'parties': Counter(),
                                       'product_raw': {}, 'party_raw': {}, 'logs': set()})
    downstream_fkeys = {}   # fkey → list of raw batch numbers (for fuzzy lookup)
    for log, sheet, uc, bi, pi, pri in sources:
        try:
            df = pd.read_excel(XLSX, sheet_name=sheet, header=3, usecols=uc)
        except Exception:
            continue
        for _, r in df.iterrows():
            b = r.iloc[bi]
            if pd.isna(b) or not str(b).strip(): continue
            rb = str(b).strip(); k = bk(b)
            downstream[k]['raw'] = rb
            downstream[k]['logs'].add(log)
            downstream_fkeys.setdefault(fkey(b), set()).add(rb)
            p = r.iloc[pi]
            if not pd.isna(p) and str(p).strip():
                rp = str(p).strip()
                downstream[k]['products'][canon(rp)] += 1
                downstream[k]['product_raw'].setdefault(canon(rp), rp)
            pa = r.iloc[pri]
            if not pd.isna(pa) and str(pa).strip():
                rpa = str(pa).strip()
                downstream[k]['parties'][canon(rpa)] += 1
                downstream[k]['party_raw'].setdefault(canon(rpa), rpa)

    # Read RM
    try:
        rm = pd.read_excel(XLSX, sheet_name='➕ RM Dispensing Log', header=3)
    except Exception:
        return []

    issues = []
    for idx, r in rm.iterrows():
        b = r.get('BATCH NUMBER')
        if pd.isna(b) or not str(b).strip(): continue
        rm_raw = str(b).strip(); k = bk(b); fk = fkey(b)
        cust = '' if pd.isna(r.get('CUSTOMER')) else str(r.get('CUSTOMER')).strip()
        prod = '' if pd.isna(r.get('NAME OF THE PRODUCT')) else str(r.get('NAME OF THE PRODUCT')).strip()

        # ── 1. BATCH NUMBER: RM doesn't match Filling exactly, but is 1 edit away
        if k not in downstream:
            # Look for fuzzy match in downstream
            target = None; kind = None
            # exact-alphanumeric (just formatting differs)
            if fk in downstream_fkeys:
                for cand in downstream_fkeys[fk]:
                    if bk(cand) != k:
                        target = cand; kind = 'formatting (spaces/dashes)'; break
            # 1 character diff
            if target is None:
                for f_inner, raws in downstream_fkeys.items():
                    if f_inner == fk: continue
                    if is_one_indel(fk, f_inner) or one_sub(fk, f_inner):
                        # Need to also confirm product/customer roughly match the downstream's product
                        cand_raw = next(iter(raws))
                        cand_info = downstream.get(bk(cand_raw))
                        if cand_info and prod:
                            cand_prods = [downstream[bk(cand_raw)]['product_raw'].get(c, '') for c in downstream[bk(cand_raw)]['products']]
                            if any(root(cp) == root(prod) for cp in cand_prods):
                                target = cand_raw
                                kind = 'one-char typo + same product'
                                break
            if target:
                fp = f"BATCH|{fk}|{bk(target)}"
                issues.append({
                    'fingerprint': fp,
                    'kind': 'batch',
                    'rm_value': rm_raw,
                    'correct': target,
                    'detail': kind,
                    'product': prod, 'customer': cust,
                    'rm_row': idx + 5,   # 1-indexed sheet row
                })
                # We don't ALSO flag product/customer for this row — fixing the batch is the first step
                continue

        # ── 2. PRODUCT NAME and CUSTOMER NAME (RM vs downstream majority)
        if k in downstream:
            for field, rm_val, downstream_counts, raw_lookup, label in [
                ('product', prod, downstream[k]['products'], downstream[k]['product_raw'], 'product'),
                ('party',   cust, downstream[k]['parties'],  downstream[k]['party_raw'],   'customer'),
            ]:
                if not rm_val or not downstream_counts: continue
                rm_canon = canon(rm_val)
                # Find downstream majority spelling
                top_canon, top_n = downstream_counts.most_common(1)[0]
                if rm_canon == top_canon: continue
                # Only flag when ≥ 2 downstream logs agree, OR all downstream uses the same spelling (any count)
                num_logs_agreeing = sum(1 for c, n in downstream_counts.items() if c == top_canon and n > 0)
                if top_n < 1: continue
                # confidence: high if multiple logs agree, medium if just one
                conf = 'high' if top_n >= 2 else 'medium'
                top_raw = raw_lookup.get(top_canon, top_canon)
                fp = f"{field.upper()}|{k}|{rm_canon}|{top_canon}"
                issues.append({
                    'fingerprint': fp,
                    'kind': field,
                    'rm_value': rm_val,
                    'correct': top_raw,
                    'detail': f"downstream uses this ({top_n} row{'s' if top_n != 1 else ''})",
                    'product': prod, 'customer': cust,
                    'rm_row': idx + 5,
                    'batch': rm_raw,
                    'confidence': conf,
                })

    # De-dup by fingerprint (in case the same issue arises across multiple RM rows for same batch)
    seen = set(); deduped = []
    for it in issues:
        if it['fingerprint'] in seen: continue
        seen.add(it['fingerprint']); deduped.append(it)
    # Sort: batch fixes first, then product, then customer; within each, by batch
    order = {'batch': 0, 'product': 1, 'party': 2}
    deduped.sort(key=lambda i: (order.get(i['kind'], 9), i.get('rm_value','')))
    return deduped


# ─── state ──────────────────────────────────────────────────────
def load_state():
    if not os.path.exists(STATE): return {}
    try:    return json.load(open(STATE))
    except: return {}
def save_state(state):
    json.dump(state, open(STATE, 'w'), indent=2, sort_keys=True)
    try:
        subprocess.run(['git', 'add', STATE], cwd=HERE, capture_output=True, check=False, timeout=10)
    except Exception:
        pass


# ─── email body ─────────────────────────────────────────────────
def build_body(to_send, is_reminder):
    batch_fixes = [i for i in to_send if i['kind'] == 'batch']
    prod_fixes  = [i for i in to_send if i['kind'] == 'product']
    party_fixes = [i for i in to_send if i['kind'] == 'party']

    lead = (
        "Hi Store team,\n\n"
        "A few entries in the RM Dispensing Log don't match the production logs "
        "(Filling / Packing / Dispatch). Could you please review the items below "
        "and make sure the RM-side entry is correct? If it is, no action needed on "
        "your end — the production team has been told separately to update their "
        "side. If the RM side has the wrong value, please fix it in the sheet.\n\n"
    ) if not is_reminder else (
        "Hi Store team,\n\n"
        "Friendly reminder — the items below were flagged a couple of days ago and "
        "are still showing as mismatches. Please take a quick look when you can.\n\n"
    )
    body = lead + "──────────────────────────────────────────────────\n\n"

    if batch_fixes:
        body += f"BATCH NUMBER — please correct in the RM sheet  ({len(batch_fixes)}):\n\n"
        body += "  These are clear formatting / typo differences where the production\n"
        body += "  log's batch number is the canonical one. Please update RM to match:\n\n"
        for i in batch_fixes:
            body += (
                f"  • Row {i['rm_row']} — {i.get('product','')} for {i.get('customer','')}:\n"
                f"      Change batch \"{i['rm_value']}\" → \"{i['correct']}\"\n"
                f"      ({i['detail']})\n\n"
            )

    if prod_fixes:
        body += "──────────────────────────────────────────────────\n\n"
        body += f"PRODUCT NAME — please verify  ({len(prod_fixes)}):\n\n"
        body += "  For each item below: is the RM-side spelling correct? If yes, the\n"
        body += "  production team will fix theirs. If no, please update RM in the sheet.\n\n"
        for i in prod_fixes:
            body += (
                f"  • Batch {i.get('batch','')} (row {i['rm_row']}):\n"
                f"      RM has        → \"{i['rm_value']}\"\n"
                f"      Other logs have → \"{i['correct']}\"  ({i['detail']})\n\n"
            )

    if party_fixes:
        body += "──────────────────────────────────────────────────\n\n"
        body += f"CUSTOMER NAME — please verify  ({len(party_fixes)}):\n\n"
        body += "  Same idea — is the RM-side customer correct, or should it be updated?\n\n"
        for i in party_fixes:
            body += (
                f"  • Batch {i.get('batch','')} (row {i['rm_row']}):\n"
                f"      RM has        → \"{i['rm_value']}\"\n"
                f"      Other logs have → \"{i['correct']}\"  ({i['detail']})\n\n"
            )

    body += (
        "──────────────────────────────────────────────────\n\n"
        "We re-check the RM sheet automatically — once a row is corrected (or stops "
        "appearing because production matched yours) it silently drops off this list. "
        f"If anything is still pending after {REMIND_AFTER_DAYS} days, you'll get a "
        f"reminder.\n\n"
        "Thanks for keeping the data clean!\n\n"
        "— Enicar Dashboard (automatic store-side check)"
    )
    return body


# ─── main ───────────────────────────────────────────────────────
def main():
    if not os.path.exists(XLSX): return 0
    if not (SENDER and APP_PW):
        print('check_store_mismatches: no email creds — skipping.'); return 0

    issues = find_store_issues()
    current_fps = {i['fingerprint']: i for i in issues}
    state = load_state()
    today = date.today(); today_iso = today.isoformat()

    to_send = []; is_reminder = False
    for fp, it in current_fps.items():
        last = state.get(fp)
        if not last:
            to_send.append(it); state[fp] = today_iso
        else:
            try: last_d = date.fromisoformat(last)
            except: last_d = today
            if (today - last_d) >= timedelta(days=REMIND_AFTER_DAYS):
                to_send.append(it); state[fp] = today_iso; is_reminder = True

    # Drop fingerprints no longer present (fixed)
    for fp in list(state):
        if fp not in current_fps:
            del state[fp]

    if to_send:
        stamp = datetime.now().strftime('%d %b %Y')
        prefix = '[Reminder] ' if is_reminder else ''
        plural = 'fix' if len(to_send) == 1 else 'fixes'
        msg = EmailMessage()
        msg['From'] = SENDER
        msg['To']   = ', '.join(TEAM)
        msg['Subject'] = f'{prefix}Enicar Store — {len(to_send)} RM-sheet {plural} needed ({stamp})'
        msg.set_content(build_body(to_send, is_reminder))
        try:
            ctx = ssl.create_default_context()
            with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ctx) as s:
                s.login(SENDER, APP_PW)
                s.send_message(msg)
            print(f"✓ Store-mismatch email sent ({len(to_send)} item{'s' if len(to_send)!=1 else ''}, "
                  f"{'reminder' if is_reminder else 'initial'}) → {', '.join(TEAM)}")
        except Exception as e:
            print(f'✗ Store-mismatch email send failed: {e}')
            return 1
    else:
        if not issues:
            print('check_store_mismatches: RM sheet is clean — no email.')
        else:
            print(f'check_store_mismatches: {len(issues)} issue(s) tracked, none due for re-email today.')

    save_state(state)
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        print(f'✗ check_store_mismatches: {e}')
        sys.exit(1)
