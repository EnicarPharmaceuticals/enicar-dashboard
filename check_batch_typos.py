#!/usr/bin/env python3
"""
check_batch_typos.py  —  weekly batch-number discrepancy review
─────────────────────────────────────────────────────────────────
Batch number links Filling → Packing → Dispatch. When the same batch
is mistyped in one log, the link breaks. This finds those and decides
which spelling is CORRECT using this rule (in priority order):

  1. The spelling that appears in MORE logs wins
     (e.g. in Filling + Packing  beats  only Dispatch).
  2. Then: the one used in more rows.
  3. Then: the one whose prefix / length is more common across all
     batches (handles the case where each appears once).

Only HIGH-CONFIDENCE typo shapes are considered, so legitimate
sequential batches (6138… vs 6139…) are never flagged:
  • dropped / extra character     6141C85001 vs 6141C8501
  • letter typo, digits identical SC-05376  vs SL-05376
  • case / spacing only           El-2417    vs EL-2417

Run:  python3 check_batch_typos.py
"""

import os, sys, re
from collections import Counter
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, '..')
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
        return 'case/spacing only' if a != b else None
    if is_one_indel(na, nb):
        return 'dropped/extra character'
    if len(na) == len(nb) and digits(na) == digits(nb) and digits(na):
        if sum(1 for x, y in zip(na, nb) if x != y) == 1:
            return 'letter typo (digits identical)'
    return None


def main():
    rows = load()
    if not rows:
        print('── Batch-number discrepancy review ──')
        print('No data could be read.'); return 0

    # Per-batch metadata
    meta = {}
    for r in rows:
        m = meta.setdefault(r['batch'], {'logs': set(), 'rows': 0, 'product': set()})
        m['logs'].add(r['log']); m['rows'] += 1
        if r['product']:
            m['product'].add(r['product'])

    batches = sorted(meta)
    # global frequencies for tie-breaking
    pref_freq = Counter(prefix(b) for b in batches if prefix(b))
    len_freq  = Counter(len(norm(b)) for b in batches)

    # union-find to cluster typo-variants
    parent = {b: b for b in batches}
    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]; x = parent[x]
        return x
    def union(a, b):
        parent[find(a)] = find(b)

    for i, a in enumerate(batches):
        for b in batches[i+1:]:
            if classify(a, b):
                union(a, b)

    clusters = {}
    for b in batches:
        clusters.setdefault(find(b), []).append(b)
    clusters = [c for c in clusters.values() if len(c) > 1]

    def rank_key(b):
        # higher = more likely the CORRECT spelling
        return (len(meta[b]['logs']), meta[b]['rows'],
                pref_freq.get(prefix(b), 0), len_freq.get(len(norm(b)), 0))

    print('── Batch-number discrepancy review ──')
    if not clusters:
        print('✓ No likely batch-number typos found. Logs link cleanly.')
        return 0

    # sort clusters: confident ones (decided by log-count) first
    def cluster_confident(c):
        s = sorted(c, key=rank_key, reverse=True)
        return len(meta[s[0]]['logs']) > len(meta[s[1]]['logs'])

    clusters.sort(key=lambda c: (not cluster_confident(c)))

    print(f'\n{len(clusters)} batch number(s) appear to be mistyped. For each, the')
    print('CORRECT spelling (used in the most logs) is shown first, then the')
    print('WRONG one(s) to fix in the sheet:\n')

    for n, c in enumerate(clusters, 1):
        ranked = sorted(c, key=rank_key, reverse=True)
        correct = ranked[0]
        wrongs = ranked[1:]
        prod = ' / '.join(sorted({p for b in c for p in meta[b]['product']}))[:70]
        c_logs = '+'.join(sorted(meta[correct]['logs'], key=LOG_ORDER.index))
        confident = len(meta[correct]['logs']) > max(len(meta[w]['logs']) for w in wrongs)
        tag = '' if confident else '   ⚠ each appears about equally — please confirm'
        print(f'{n}. Product: {prod}')
        print(f'   ✅ CORRECT : {correct!r:16} (in {c_logs})')
        for w in wrongs:
            w_logs = '+'.join(sorted(meta[w]['logs'], key=LOG_ORDER.index))
            print(f'   ❌ FIX     : {w!r:16} (in {w_logs})  — {classify(correct, w)}')
        if tag:
            print(tag)
        print()
    return 0


if __name__ == '__main__':
    sys.exit(main())
