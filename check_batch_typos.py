#!/usr/bin/env python3
"""
check_batch_typos.py  —  weekly batch-number discrepancy review
─────────────────────────────────────────────────────────────────
Batch number links Filling → Packing → Dispatch. When the same batch
is mistyped in one log, the link breaks. This finds those and decides
which spelling is CORRECT using this rule (in priority order):
  1. The spelling that appears in MORE logs wins (e.g. Filling+Packing
     beats only Dispatch).
  2. Then the one used in more rows.
  3. Then the one whose prefix / length is more common across batches.

Only high-confidence typo shapes are flagged, so legitimate sequential
batches (6138… vs 6139…) are never flagged.

Run:  python3 check_batch_typos.py
"""

import os, sys, re
from collections import Counter
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
XLSX = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')
LOG_ORDER = ['Filling', 'Packing', 'Dispatch']


def load():
    specs = [('Filling', '➕ Filling Log', 'B:J', {'batch': 6, 'product': 2}),
             ('Packing', '➕ Packing Log', 'B:N', {'batch': 5, 'product': 2}),
             ('Dispatch', '➕ Dispatch Log', 'B:I', {'batch': 5, 'product': 1})]
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
            prod = r.iloc[idx['product']]
            rows.append({'log': log, 'batch': str(b).strip(),
                         'product': '' if pd.isna(prod) else str(prod).strip()})
    return rows


def norm(b):    return re.sub(r'\s+', '', str(b)).upper()
def digits(b):  return re.sub(r'[^0-9]', '', str(b))
def prefix(b):  m = re.match(r'^([A-Z]+)', norm(b)); return m.group(1) if m else ''


def is_one_indel(a, b):
    if abs(len(a) - len(b)) != 1:
        return False
    short, long = (a, b) if len(a) < len(b) else (b, a)
    i = j = 0; skipped = False
    while i < len(short) and j < len(long):
        if short[i] == long[j]:
            i += 1; j += 1
        elif skipped:
            return False
        else:
            skipped = True; j += 1
    return True


def classify(a, b):
    na, nb = norm(a), norm(b)
    if na == nb:
        return 'a small case/spacing difference' if a != b else None
    if is_one_indel(na, nb):
        return 'a missing or extra character'
    if len(na) == len(nb) and digits(na) == digits(nb) and digits(na):
        if sum(1 for x, y in zip(na, nb) if x != y) == 1:
            return 'one wrong letter (the numbers match)'
    return None


def nice_logs(logs):
    names = sorted(logs, key=LOG_ORDER.index)
    if len(names) == 1:
        return names[0]
    return ' and '.join([', '.join(names[:-1]), names[-1]]) if len(names) > 2 else ' and '.join(names)


def main():
    rows = load()
    print('── Batch number check ──\n')
    if not rows:
        print('Could not read the data this week.'); return 0

    meta = {}
    prod_count = {}
    for r in rows:
        m = meta.setdefault(r['batch'], {'logs': set(), 'rows': 0})
        m['logs'].add(r['log']); m['rows'] += 1
        if r['product']:
            prod_count.setdefault(r['batch'], Counter())[r['product']] += 1

    batches = sorted(meta)
    pref_freq = Counter(prefix(b) for b in batches if prefix(b))
    len_freq  = Counter(len(norm(b)) for b in batches)

    parent = {b: b for b in batches}
    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]; x = parent[x]
        return x
    for i, a in enumerate(batches):
        for b in batches[i+1:]:
            if classify(a, b):
                parent[find(a)] = find(b)

    clusters = {}
    for b in batches:
        clusters.setdefault(find(b), []).append(b)
    clusters = [c for c in clusters.values() if len(c) > 1]

    if not clusters:
        print('Good news — every batch number lines up across the logs. Nothing to fix. ✅')
        return 0

    def rank_key(b):
        return (len(meta[b]['logs']), meta[b]['rows'],
                pref_freq.get(prefix(b), 0), len_freq.get(len(norm(b)), 0))

    def product_of(cluster):
        c = Counter()
        for b in cluster:
            c += prod_count.get(b, Counter())
        return c.most_common(1)[0][0] if c else 'product unknown'

    confident, uncertain = [], []
    for c in clusters:
        ranked = sorted(c, key=rank_key, reverse=True)
        correct, wrong = ranked[0], ranked[1]
        if len(meta[correct]['logs']) > len(meta[wrong]['logs']):
            confident.append((c, correct, ranked[1:]))
        else:
            uncertain.append((c, correct, ranked[1:]))

    print('A few batch numbers look like small typing slips — the same batch was')
    print('written differently in different logs. Please fix the wrong one in the sheet.\n')

    if confident:
        print('CONFIDENT — the correct number is the one used in 2 logs:\n')
        for c, correct, wrongs in confident:
            prod = product_of(c)
            for w in wrongs:
                print(f'  • {prod}: change "{w}" to "{correct}"')
                print(f'       (because "{correct}" appears in {nice_logs(meta[correct]["logs"])}, '
                      f'but "{w}" only in {nice_logs(meta[w]["logs"])} — {classify(correct, w)})')
            print()

    if uncertain:
        print('PLEASE DOUBLE-CHECK — each spelling appears only once, so this is our best guess:\n')
        for c, correct, wrongs in uncertain:
            prod = product_of(c)
            for w in wrongs:
                print(f'  • {prod}: likely change "{w}" to "{correct}"  ({classify(correct, w)})')
            print()

    return 0


if __name__ == '__main__':
    sys.exit(main())
