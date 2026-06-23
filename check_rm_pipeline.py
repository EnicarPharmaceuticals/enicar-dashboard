#!/usr/bin/env python3
"""
check_rm_pipeline.py  —  Weekly RM pipeline-status summary
─────────────────────────────────────────────────────────────────
Reads the RM Dispensing Log + Filling/Packing/Dispatch logs and
produces a humanized end-to-end pipeline status report for the team:

    ✅ End-to-end          — batch went RM → Fill → Pack → Disp
    🟡 Packed, awaiting    — RM → Fill → Pack, not yet dispatched
    🔵 Filled, awaiting    — RM → Fill, not yet packed
    ⚪ Done before tracking — already shipped before tracking started
                              (matched via the opening-stock baseline)
    ⚠ Pending manufacture  — RM dispensed, no production record yet

Filters applied (matching the user's review rules):
  * DISPENSING DATE not blank (only actually-dispensed rows)
  * PLAN = REGULAR             (skips TRIAL / RETURN / ADDITIONAL)
  * BATCH NUMBER does not contain 'TLB' (skips R&D trial batches)

Output is plain text on stdout — cloud_weekly_review.py appends it
to Weekly_Data_Review.txt which gets emailed.
"""

import os, sys, re, json
from datetime import date
import pandas as pd

HERE     = os.path.dirname(os.path.abspath(__file__))
ROOT     = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
XLSX     = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')
BASELINE = os.path.join(HERE, 'batch_baseline.json')


def bk(b): return re.sub(r'\s+', '', str(b)).upper()


def main():
    print('── Production-pipeline status (RM → Fill → Pack → Disp) ──\n')
    if not os.path.exists(XLSX):
        print('  Could not read the data this week.\n'); return 0

    # Filtered RM set
    try:
        rm = pd.read_excel(XLSX, sheet_name='➕ RM Dispensing Log', header=3)
        # Normalize column headers — collapse newlines / multi-spaces so names like
        # 'DISPENSING \nDATE' match 'DISPENSING DATE'. Survives sheet-formatting tweaks.
        rm.columns = [' '.join(str(c).split()) for c in rm.columns]
    except Exception:
        print('  No RM Dispensing Log tab yet — pipeline view unavailable.\n'); return 0
    if 'PLAN' not in rm.columns or 'DISPENSING DATE' not in rm.columns:
        print('  RM tab is present but the expected columns are missing — skipping this week.\n'); return 0
    rm = rm[rm['DISPENSING DATE'].notna() &
            (rm['PLAN'].astype(str).str.strip().str.upper() == 'REGULAR') &
            ~rm['BATCH NUMBER'].astype(str).str.upper().str.contains('TLB', na=False)].copy()
    rm['DISPENSING DATE'] = pd.to_datetime(rm['DISPENSING DATE'], errors='coerce')

    # Downstream log batch keys
    def keys(sheet, uc, bi):
        try:
            df = pd.read_excel(XLSX, sheet_name=sheet, header=3, usecols=uc)
            return set(bk(x) for x in df.iloc[:, bi].dropna() if str(x).strip())
        except Exception:
            return set()
    fill_keys = keys('➕ Filling Log',  'B:J', 6)
    pack_keys = keys('➕ Packing Log',  'B:N', 5)
    disp_keys = keys('➕ Dispatch Log', 'B:I', 5)
    try:
        baseline = set(json.load(open(BASELINE)).get('batches', []))
    except Exception:
        baseline = set()

    # Bucket each RM row
    buckets = {'end':0, 'pack':0, 'fill':0, 'legacy':0, 'pending':0}
    pending_rows = []
    for _, r in rm.iterrows():
        b = r.get('BATCH NUMBER')
        if pd.isna(b) or not str(b).strip():
            continue
        k = bk(b)
        if   k in disp_keys: buckets['end'] += 1
        elif k in pack_keys: buckets['pack'] += 1
        elif k in fill_keys: buckets['fill'] += 1
        elif k in baseline:  buckets['legacy'] += 1
        else:
            buckets['pending'] += 1
            d = r['DISPENSING DATE']
            pending_rows.append({
                'batch':    str(b).strip(),
                'customer': str(r.get('CUSTOMER', '')).strip(),
                'product':  str(r.get('NAME OF THE PRODUCT', '')).strip(),
                'date':     d.date() if not pd.isna(d) else None,
                'size':     float(pd.to_numeric(r.get('BATCH SIZE'), errors='coerce') or 0),
            })

    total = sum(buckets.values())
    if total == 0:
        print('  No qualifying RM rows this week — nothing to summarise.\n'); return 0

    def pct(n): return f'{(100*n/total):4.1f}%' if total else '0%'

    print(f'  Out of {total} batches dispensed (REGULAR plan), here is where each one stands today:\n')
    print(f'    ✅  End-to-end (RM → Fill → Pack → Dispatched) ............... {buckets["end"]:>4}   ({pct(buckets["end"])})')
    print(f'    🟡  Packed, awaiting dispatch ................................ {buckets["pack"]:>4}   ({pct(buckets["pack"])})')
    print(f'    🔵  Filled, awaiting packing ................................. {buckets["fill"]:>4}   ({pct(buckets["fill"])})')
    print(f'    ⚪  Done before tracking started (legacy baseline) ........... {buckets["legacy"]:>4}   ({pct(buckets["legacy"])})')
    print(f'    ⚠  Pending manufacture (dispensed, no production yet) ....... {buckets["pending"]:>4}   ({pct(buckets["pending"])})')
    print()

    # Show the pending list inline (oldest first), capped at 30 lines
    if pending_rows:
        pending_rows.sort(key=lambda r: (r['date'] or date.min, r['customer'].lower(), r['batch']))
        print('  The pending-manufacture batches (oldest first):\n')
        print(f'    {"Dispensed":12}{"Customer":30}{"Product":30}{"Batch":16}{"Size":>8}')
        print('    ' + '-' * 96)
        for r in pending_rows[:30]:
            d = str(r['date']) if r['date'] else '—'
            print(f'    {d:12}{r["customer"][:29]:30}{r["product"][:29]:30}{r["batch"][:15]:16}{r["size"]:>8.0f}')
        if len(pending_rows) > 30:
            print(f'    … and {len(pending_rows)-30} more (open the dashboard search to look up any specific batch)')
        print()

    print('  Good news rule of thumb — when "Pending" stays close to "what was dispensed this week",')
    print('  production is keeping pace with dispensing. If "Pending" grows week-over-week, batches')
    print('  are piling up faster than they\'re being made.\n')
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        # Never fail the cloud workflow because of this check.
        print(f'  (pipeline status unavailable: {e})\n')
        sys.exit(0)
