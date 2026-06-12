#!/usr/bin/env python3
"""
check_batch_typos.py  —  weekly batch-number discrepancy review
─────────────────────────────────────────────────────────────────
Batch number links RM → Filling → Packing → Dispatch. When the same
batch is mistyped in one log, the link breaks. This finds those and
decides which spelling is CORRECT using this rule (in priority order):
  1. A spelling that appears in RM Dispensing wins (RM is filled in
     first, straight from the BMR — source of truth).
  2. Then the spelling that appears in MORE logs.
  3. Then the one used in more rows.
  4. Then the one whose prefix / length is more common across batches.

Only high-confidence typo shapes are flagged, so legitimate sequential
batches (6138… vs 6139…) are never flagged.

A second section catches ORPHAN batches: a batch that appears in only
one production log, has no RM record, and is one small slip away from a
same-product batch that does exist elsewhere — e.g. Packing "RE-023230"
when Filling/RM have "RE-021330" for the same product.

Run:  python3 check_batch_typos.py
"""

import os, sys, re
from collections import Counter
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
XLSX = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')
LOG_ORDER = ['RM Dispensing', 'Filling', 'Packing', 'Dispatch']


def load():
    specs = [('RM Dispensing', '➕ RM Dispensing Log', 'B:I', {'batch': 5, 'product': 4}),
             ('Filling', '➕ Filling Log', 'B:J', {'batch': 6, 'product': 2}),
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
def pcanon(p):  return re.sub(r'[^a-z0-9]', '', str(p).lower())


def products_compatible(a, b):
    """Same product allowing for naming-style differences:
    'Redsun jelly ( Chocolate)' vs 'REDSUN JELLY (Chocolate Flavour)'."""
    ca, cb = pcanon(a), pcanon(b)
    if not ca or not cb or min(len(ca), len(cb)) < 6:
        return False
    return ca.startswith(cb) or cb.startswith(ca)


def lev(a, b):
    """Plain edit distance for short batch strings."""
    if a == b:
        return 0
    if len(a) < len(b):
        a, b = b, a
    prev = list(range(len(b) + 1))
    for i, ch in enumerate(a, 1):
        cur = [i]
        for j, ch2 in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j-1] + 1, prev[j-1] + (ch != ch2)))
        prev = cur
    return prev[-1]


def near_slip(a, b):
    """True when a→b looks like a typing slip: 1 edit, an adjacent
    transposition, an inserted/dropped run of ≤2 chars, or 2 substitutions
    side-by-side (e.g. RE-023230 vs RE-021330). Same-prefix required so
    e.g. EL-#### never matches RE-####."""
    na, nb = norm(a), norm(b)
    if na == nb or prefix(na) != prefix(nb):
        return False
    d = lev(na, nb)
    if d == 1:
        return True
    if d == 2:
        if abs(len(na) - len(nb)) == 2:        # 2-char insert/drop (RE-1329 vs RE-011329)
            return True
        if len(na) == len(nb):
            diff = [i for i in range(len(na)) if na[i] != nb[i]]
            if len(diff) == 2 and diff[1] - diff[0] == 1:
                return True                     # adjacent pair (transposition or 2-slip)
    return False


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

    def rank_key(b):
        return ('RM Dispensing' in meta[b]['logs'],
                len(meta[b]['logs']), meta[b]['rows'],
                pref_freq.get(prefix(b), 0), len_freq.get(len(norm(b)), 0))

    # ── Orphan batches: in ONE production log, no RM record, one slip away
    #    from a same-product batch that exists elsewhere ──
    clustered = {b for c in clusters for b in c}
    try:
        import json
        baseline = {norm(x) for x in json.load(
            open(os.path.join(HERE, 'batch_baseline.json'))).get('batches', [])}
    except Exception:
        baseline = set()

    def main_product(b):
        c = prod_count.get(b, Counter())
        return c.most_common(1)[0][0] if c else ''

    orphans = []
    for x in batches:
        mx = meta[x]
        logs_x = mx['logs'] - {'RM Dispensing'}
        if ('RM Dispensing' in mx['logs'] or len(logs_x) != 1
                or x in clustered or norm(x) in baseline):
            continue
        px = main_product(x)
        cands = []
        for y in batches:
            if y == x or y in clustered:
                continue
            my = meta[y]
            stronger = ('RM Dispensing' in my['logs']
                        or len(my['logs']) > len(mx['logs']))
            if stronger and near_slip(x, y) and products_compatible(px, main_product(y)):
                cands.append(y)
        if cands:
            # When several candidates are equally plausible (same edit distance,
            # same product), prefer one that exists in RM Dispensing — RM is the
            # source of truth, so an RM batch is much more likely to be the real
            # one the production team mis-typed. Without this, two siblings like
            # RE-881271 (RM ✓) and RE-891272 (RM ✓) both match orphan RE-881272,
            # and the checker would pick at random; we want the RM one.
            def orphan_key(y):
                return ('RM Dispensing' in meta[y]['logs'], rank_key(y))
            best = max(cands, key=orphan_key)
            # Surface other equally-plausible siblings so a human can compare.
            alt = [y for y in cands if y != best
                   and ('RM Dispensing' in meta[y]['logs'])
                   and lev(norm(x), norm(y)) == lev(norm(x), norm(best))]
            orphans.append((x, best, alt))

    if not clusters and not orphans:
        print('Good news — every batch number lines up across the logs. Nothing to fix. ✅')
        return 0

    def product_of(cluster):
        c = Counter()
        for b in cluster:
            c += prod_count.get(b, Counter())
        return c.most_common(1)[0][0] if c else 'product unknown'

    confident, uncertain = [], []
    for c in clusters:
        ranked = sorted(c, key=rank_key, reverse=True)
        correct, wrong = ranked[0], ranked[1]
        rm_wins = ('RM Dispensing' in meta[correct]['logs']
                   and 'RM Dispensing' not in meta[wrong]['logs'])
        if rm_wins or len(meta[correct]['logs']) > len(meta[wrong]['logs']):
            confident.append((c, correct, ranked[1:]))
        else:
            uncertain.append((c, correct, ranked[1:]))

    if clusters:
        print('A few batch numbers look like small typing slips — the same batch was')
        print('written differently in different logs. Please fix the wrong one in the sheet.\n')

    if confident:
        print('CONFIDENT — the correct number is the one in RM Dispensing / the majority of logs:\n')
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

    if orphans:
        print('ORPHAN BATCHES — these appear in only ONE log, have no RM Dispensing record,')
        print('and are one small slip away from a same-product batch that does exist.')
        print('Most likely the batch number was mistyped — please verify against the BMR:\n')
        for x, best, alt in orphans:
            log_x = nice_logs(meta[x]['logs'] - {'RM Dispensing'})
            print(f'  • {main_product(x)}: "{x}" (only in {log_x}) — should this be '
                  f'"{best}" (seen in {nice_logs(meta[best]["logs"])})?')
            if alt:
                others = ' or '.join(f'"{a}"' for a in alt)
                print(f'      (note: {others} also matches — please check the BMR to be sure)')
        print()

    return 0


if __name__ == '__main__':
    sys.exit(main())
