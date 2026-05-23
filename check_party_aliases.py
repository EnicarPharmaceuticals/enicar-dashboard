#!/usr/bin/env python3
"""
check_party_aliases.py  —  weekly customer-alias review helper
─────────────────────────────────────────────────────────────────
Scans the latest data for customer/party names that are NOT yet in the
alias list in generate_enicar_html.py, and flags likely duplicates of
names that ARE already known (via fuzzy similarity). Prints a short
report so the alias list can be updated.

Run:  python3 check_party_aliases.py
"""

import os, sys, re
from difflib import SequenceMatcher

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, '..')
sys.path.insert(0, HERE)

import pandas as pd
import importlib.util

# Pull the alias maps straight from the generator (single source of truth) by
# parsing just the _PARTY_GROUPS / _pkey definitions — no data read needed.
gen_src = open(os.path.join(HERE, 'generate_enicar_html.py')).read()
_start = gen_src.index('_PARTY_GROUPS = {')
_end   = gen_src.index('def normalise_party')
ns = {'re': re}
exec(gen_src[_start:_end], ns)
LOOKUP = ns['_PARTY_LOOKUP']
CANONS = list(ns['_PARTY_GROUPS'].keys())
pkey   = ns['_pkey']

XLSX = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')


def all_party_names():
    names = []
    for sheet, usecols, idx in [('➕ Dispatch Log', 'B:I', 6),
                                ('➕ Packing Log', 'B:N', 11),
                                ('➕ Filling Log', 'B:J', 7)]:
        try:
            df = pd.read_excel(XLSX, sheet_name=sheet, header=3, usecols=usecols)
            names += df.iloc[:, idx].dropna().astype(str).str.strip().tolist()
        except Exception:
            pass
    return [n for n in names if n and n.lower() != 'nan']


def main():
    names = all_party_names()
    unknown = {}
    for n in names:
        if pkey(n) not in LOOKUP:
            unknown[n] = unknown.get(n, 0) + 1

    print('── Customer / party alias review ──')
    if not unknown:
        print('✓ Every customer name in the data is already in the alias list. Nothing to do.')
        return 0

    print(f'\n{len(unknown)} name(s) NOT yet in the alias list:\n')
    for name, cnt in sorted(unknown.items(), key=lambda x: -x[1]):
        # suggest the closest known canonical
        best, score = None, 0.0
        for c in CANONS:
            s = SequenceMatcher(None, pkey(name), pkey(c)).ratio()
            if s > score:
                best, score = c, s
        hint = f'   ↳ looks like "{best}"?' if score >= 0.6 else '   ↳ (looks like a NEW customer)'
        print(f'  {cnt:3}×  {name!r}')
        print(hint)
    print('\nTo fix: open generate_enicar_html.py → _PARTY_GROUPS, add each')
    print('typo under the right customer, or add a new entry for new customers.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
