#!/usr/bin/env python3
"""
check_batch_typos.py  —  weekly batch-number typo review
─────────────────────────────────────────────────────────────────
Batch number is the key that links Filling → Packing → Dispatch.
Small typing slips break that link. This flags HIGH-CONFIDENCE typos
only, so legitimate sequential batches are NOT falsely flagged:

  1. Dropped / extra character   e.g.  6141C8501  vs  6141C85001
  2. Letter typo, digits identical e.g. SC-05376  vs  SL-05376
  3. Case mismatch                e.g. El-2417   vs  EL-2417

Pure single-digit changes (e.g. 6138 vs 6139) are deliberately NOT
flagged — those are almost always different real batches.

Run:  python3 check_batch_typos.py
"""

import os, sys, re
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, '..')
XLSX = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')


def load():
    specs = [('Filling', '➕ Filling Log', 'B:J', {'batch': 6, 'product': 2, 'party': 7}),
             ('Packing', '➕ Packing Log', 'B:N', {'batch': 5, 'product': 2, 'party': 11}),
             ('Dispatch', '➕ Dispatch Log', 'B:I', {'batch': 5, 'product': 1, 'party': 6})]
    rows = []
    for log, sheet, usecols, idx in specs:
        try:
            df = pd.read_excel(XLSX, sheet_name=sheet, header=3, usecols=usecols)
        except Exception:
            continue
        for _, r in df.iterrows():
            b = r.iloc[idx['batch']]
            if pd.isna(b):
                continue
            rows.append({
                'log': log,
                'batch': str(b).strip(),
                'product': str(r.iloc[idx['product']]).strip() if not pd.isna(r.iloc[idx['product']]) else '',
                'party': str(r.iloc[idx['party']]).strip() if not pd.isna(r.iloc[idx['party']]) else '',
            })
    return rows


def norm(b):
    return re.sub(r'\s+', '', str(b)).upper()


def digits(b):
    return re.sub(r'[^0-9]', '', str(b))


def is_one_indel(a, b):
    """True if a and b differ by exactly one inserted/deleted character."""
    if abs(len(a) - len(b)) != 1:
        return False
    short, long = (a, b) if len(a) < len(b) else (b, a)
    i = j = 0
    skipped = False
    while i < len(short) and j < len(long):
        if short[i] == long[j]:
            i += 1; j += 1
        elif skipped:
            return False
        else:
            skipped = True; j += 1
    return True


def classify(a, b):
    """Return a typo reason if (a,b) look like the same batch mistyped, else None."""
    na, nb = norm(a), norm(b)
    if na == nb:
        return 'case/spacing only' if a != b else None
    if is_one_indel(na, nb):
        return 'dropped/extra character'
    if len(na) == len(nb) and digits(na) == digits(nb) and digits(na):
        diffs = sum(1 for x, y in zip(na, nb) if x != y)
        if diffs == 1:
            return 'letter typo (digits identical)'
    return None


def main():
    rows = load()
    # unique batch strings per log
    by_log = {}
    meta = {}
    for r in rows:
        by_log.setdefault(r['log'], set()).add(r['batch'])
        meta.setdefault(r['batch'], {'product': set(), 'party': set(), 'logs': set()})
        meta[r['batch']]['product'].add(r['product'])
        meta[r['batch']]['party'].add(r['party'])
        meta[r['batch']]['logs'].add(r['log'])

    all_batches = sorted(meta)
    seen = set()
    findings = []
    for i, a in enumerate(all_batches):
        for b in all_batches[i+1:]:
            reason = classify(a, b)
            if not reason:
                continue
            # only interesting if they sit in DIFFERENT logs (a broken link)
            if meta[a]['logs'] == meta[b]['logs'] and len(meta[a]['logs']) == 1:
                # same single log — still worth flagging as a duplicate typo
                pass
            key = tuple(sorted((a, b)))
            if key in seen:
                continue
            seen.add(key)
            findings.append((a, b, reason))

    print('── Batch-number typo review ──')
    if not findings:
        print('✓ No likely batch-number typos found. Filling/Packing/Dispatch link cleanly.')
        return 0

    print(f'\n{len(findings)} possible batch typo(s) — review and fix at the source:\n')
    for a, b, reason in findings:
        la = '/'.join(sorted(meta[a]['logs']))
        lb = '/'.join(sorted(meta[b]['logs']))
        prod = ' / '.join(sorted({p for p in (meta[a]['product'] | meta[b]['product']) if p}))[:60]
        print(f'  {a!r} [{la}]  ⟷  {b!r} [{lb}]')
        print(f'     reason: {reason}   product: {prod}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
