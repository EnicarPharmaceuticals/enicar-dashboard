#!/usr/bin/env python3
"""
generate_enicar_html.py
Reads Enicar_Dashboard_Template.xlsx → generates Enicar_Dashboard.html

Usage:
  python3 generate_enicar_html.py              # uses current month
  python3 generate_enicar_html.py 2026 5       # specify year and month

Double-click Refresh_Dashboard.command to run from Finder.
"""

import os, sys, calendar, base64, warnings, re
from datetime import date, datetime

warnings.filterwarnings('ignore', category=UserWarning, module='pandas')

import pandas as pd

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION — edit these if needed
# ══════════════════════════════════════════════════════════════════════════════
HERE         = os.path.dirname(os.path.abspath(__file__))
# ROOT holds the .xlsx input and the .html output. Defaults to the parent
# folder (Enicar Report/) on the Mac, but can be overridden via the
# DASHBOARD_ROOT env var when running in the cloud (GitHub Actions).
ROOT         = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
TEMPLATE     = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')
OUTPUT       = os.path.join(ROOT, 'Enicar_Dashboard.html')
BSR_OPENING  = 0      # ← Set your opening BSR stock balance here (units)

today = datetime.today()

def _latest_month_with_data():
    """Look at the Filling/Packing/Dispatch logs and return (year, month) of
    the most recent row. This avoids showing an empty current month right after
    a month rollover (e.g. on 1 June when only May has data so far)."""
    try:
        import pandas as _pd
        latest = None
        for sheet, uc in [('➕ Filling Log','B:J'),
                          ('➕ Packing Log','B:N'),
                          ('➕ Dispatch Log','B:I')]:
            df = _pd.read_excel(TEMPLATE, sheet_name=sheet, header=3, usecols=uc)
            d = _pd.to_datetime(df.iloc[:,0], format='mixed', dayfirst=True, errors='coerce').dropna()
            if len(d):
                m = d.max()
                if latest is None or m > latest: latest = m
        return (latest.year, latest.month) if latest is not None else (today.year, today.month)
    except Exception:
        return (today.year, today.month)

if len(sys.argv) > 2:
    YEAR  = int(sys.argv[1])
    MONTH = int(sys.argv[2])
else:
    YEAR, MONTH = _latest_month_with_data()

# Lines and parties are read dynamically from your Excel — no hardcoding needed

C_PRI = '#004D40'; C_SEC = '#00695C'; C_AMB = '#BF360C'
C_ORG = '#E65100'; C_GRN = '#1B5E20'; C_LBG = '#E0F2F1'

# ══════════════════════════════════════════════════════════════════════════════
# READ EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def read_log(sheet, usecols, col_names, numeric_cols):
    try:
        df = pd.read_excel(TEMPLATE, sheet_name=sheet, header=3, usecols=usecols)
        df.columns = col_names
        df = df.dropna(subset=['Date'])
        df['Date'] = pd.to_datetime(df['Date'], format='mixed', dayfirst=True, errors='coerce').dt.date
        df = df.dropna(subset=['Date'])
        for c in numeric_cols:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        return df
    except Exception as e:
        print(f'  Warning: could not read "{sheet}" — {e}')
        return pd.DataFrame(columns=col_names)

# Read filling by HEADER NAME instead of position — so the dashboard doesn't
# break when the team adds or removes a column. (Same lookup-by-name fix we
# applied to Packing after the Naked column drama.)
_fill_raw = pd.read_excel(TEMPLATE, sheet_name='➕ Filling Log', header=3)

def _fcol(*candidates):
    norm = {re.sub(r'\s+',' ', str(c)).strip().lower(): c for c in _fill_raw.columns}
    for cand in candidates:
        key = re.sub(r'\s+',' ', cand).strip().lower()
        if key in norm:
            return _fill_raw[norm[key]]
    return pd.Series([pd.NA]*len(_fill_raw))

fill_df = pd.DataFrame({
    'Date':        _fcol('Date'),
    'Line':        _fcol('Packing Line','Line'),
    'Product':     _fcol('Product Name','Product'),
    'PackSize':    _fcol('Pack Size'),
    'ProductType': _fcol('Product Type'),
    'Qty':         _fcol('Qty Filled (Units)','Qty Filled','Qty'),
    'Batch':       _fcol('Batch No.','Batch'),
    'Party':       _fcol('Party Name','Customer','Party'),
    'Remarks':     _fcol('Remarks'),
})
fill_df = fill_df.dropna(subset=['Date'])
fill_df['Date'] = pd.to_datetime(fill_df['Date'], format='mixed', dayfirst=True, errors='coerce').dt.date
fill_df = fill_df.dropna(subset=['Date'])
fill_df['Qty'] = pd.to_numeric(fill_df['Qty'], errors='coerce').fillna(0)
# When a column is missing from the sheet (e.g. Product Name was deleted),
# the lookup returns pd.NA — convert to None so it renders as blank, not "<NA>".
for c in ['Product','PackSize','ProductType','Batch','Party','Line','Remarks']:
    fill_df[c] = fill_df[c].where(fill_df[c].notna(), None)

def parse_packed_total(val):
    """Handles plain numbers AND text formulas like '38 x 1600= 60800' or '527x24=12648'."""
    if pd.isna(val):
        return 0
    try:
        return float(val)
    except Exception:
        s = str(val).strip()
        if '=' in s:
            after_eq = s.split('=')[-1].strip().replace(',', '').replace(' ', '')
            try:
                return float(after_eq)
            except Exception:
                pass
        return 0

# Read packing. We look up columns by HEADER NAME instead of position so the
# dashboard doesn't break when the team adds/removes a column (e.g. when the
# "Naked" packing-quantity column was removed in June 2026).
_pack_raw = pd.read_excel(TEMPLATE, sheet_name='➕ Packing Log', header=3)

def _pcol(*candidates):
    """Find a column in pack_raw whose header matches any candidate (case-insensitive,
    whitespace-collapsed). Returns the matching pandas Series, or a Series of NaN."""
    norm = {re.sub(r'\s+',' ', str(c)).strip().lower(): c for c in _pack_raw.columns}
    for cand in candidates:
        key = re.sub(r'\s+',' ', cand).strip().lower()
        if key in norm:
            return _pack_raw[norm[key]]
    return pd.Series([pd.NA]*len(_pack_raw))

pack_df = pd.DataFrame({
    'Date':         _pcol('Date'),
    'Line':         _pcol('PACKING LINE','Packing Line','Line'),
    'Product':      _pcol('Product Name','Product'),
    'PackSize':     _pcol('Pack Size'),
    'ProdType':     _pcol('Product Type'),
    'Batch':        _pcol('Batch No.','Batch'),
    'AutoCarton':   _pcol('Carton (Auto Cartinator)','Carton(Auto Cartinator)','Auto Cartinator','AutoCarton'),
    'ManualCarton': _pcol('Carton (Manual)','Carton(Manual)','Manual Carton','ManualCarton'),
    'Sleeve':       _pcol('Sleeve'),
    'Naked':        _pcol('Naked'),   # absent in current sheet — column of NaNs is fine
    'Party':        _pcol('Party Name','Customer','Party'),
    'Remarks':      _pcol('Remarks / Urgency','Remarks/Urgency','Remarks'),
})
pack_df = pack_df.dropna(subset=['Date'])
pack_df['Date'] = pd.to_datetime(pack_df['Date'], format='mixed', dayfirst=True, errors='coerce').dt.date
pack_df = pack_df.dropna(subset=['Date'])
# Each sub-column may be a plain number OR a formula string like "38 x 1600= 60800";
# parse_packed_total handles both. Total packed = sum of all packing-qty columns.
for c in ['AutoCarton','ManualCarton','Sleeve','Naked']:
    pack_df[c] = pack_df[c].apply(parse_packed_total)
pack_df['TotalPacked'] = pack_df[['AutoCarton','ManualCarton','Sleeve','Naked']].sum(axis=1)

# Canonicalise line names so typo variants collapse onto one clean label.
#   "line no1", "Line No.4", "LINE NO 03", "lineno5" → "Line No 1/4/3/5"
#   sachet / stick pack / tube / ointment / external variants → canonical
_LINE_NUM_RE = re.compile(r'^line\s*no\.?\s*0*(\d+)$', re.IGNORECASE)
_LINE_SPECIAL = {
    'flat sachet': 'Flat Sachet', 'flat sachets': 'Flat Sachet',
    'stick pack sachet': 'Stick Pack Sachet', 'stick pack': 'Stick Pack Sachet',
    'stick-pack': 'Stick Pack Sachet', 'stickpack': 'Stick Pack Sachet',
    'stick pack line': 'Stick Pack Sachet',
    'sachet': 'Sachet', 'sachets': 'Sachet', 'pouch': 'Sachet', 'sachet line': 'Sachet',
    'ointment': 'Ointment', 'ointments': 'Ointment', 'ointment line': 'Ointment',
    'tube': 'Ointment', 'tubes': 'Ointment',
    'external': 'External', 'external line': 'External',
}
def normalise_line(s):
    if pd.isna(s): return s
    raw = ' '.join(str(s).strip().split())   # collapse double spaces
    m = _LINE_NUM_RE.match(raw)
    if m:
        return f'Line No {int(m.group(1))}'
    return _LINE_SPECIAL.get(raw.lower(), raw.title())

# ──────────────────────────────────────────────────────────────────────────────
# CUSTOMER / PARTY ALIASES  ←──  EDIT THIS LIST WEEKLY
# ------------------------------------------------------------------------------
# Each customer has ONE canonical name (the dict key) and a list of every
# spelling/typo that should map to it. To merge a newly-spotted duplicate,
# just add the misspelling to the right list — or add a new "Canonical": [...]
# entry for a brand-new customer.  Matching ignores case, extra spaces and dots.
# ──────────────────────────────────────────────────────────────────────────────
_PARTY_GROUPS = {
    'Procter & Gamble':              ['Procter & Gamble', 'P&G'],
    'Ronak Exim pvt Ltd':            ['Ronak Exim', 'Ronak exim ltd', 'Ronak Exim Ltd',
                                      'Ronak Exim P vt Ltd', 'Ronak Exim pvt Ltd',
                                      'Ronak Exim pvt Ltd.', 'Ranak Exim pvt Ltd.', 'RONAK'],
    'Galaxy Pharma':                 ['Galaxy', 'Galaxy pharma', 'Galxy pharma',
                                      'Galexy pharma', 'GALEXY'],
    'Macleods':                      ['Macleods', 'Macleoads', 'MACLEODES'],
    'Sapphire Lifescience Pvt. Ltd.': [
                                      'Sapphire Lifescience Pvt Ltd',
                                      'Sapphire Lifescience Pvt. Ltd',
                                      'Sapphire Lifescience Pvt. Ltd.',
                                      'Sapphire Life Science L TD',
                                      'Sapphire Life Science Ltd',
                                      'Sapphire Life Sciences Ltd',
                                      'Sapphire Lifesciences  pvt. Ltd.',
                                      'Sapphire lifesciences pvt ltd',
                                      'Sapphire life science Ltd',
                                      'Saphaire lifescience ltd',
                                      'Sapphire lifescianes p ltd',
                                      'SAPPHIRE'],
    'Lesanto Laboratories':          ['Lesanto', 'Lesanto Laboratories',
                                      'Lesanto labroatories', 'LESANTO'],
    'Group Pharma':                  ['GROUP', 'Group Pharma', 'Group Pharmaceutical',
                                      'Group Pharmaceutical Pharma', 'Group Pharmaceutical pharma'],
    'IPC Healthcare':                ['IPC Healthcare', 'IPC Healthcare Pvt Ltd',
                                      'Ipc Healthcare pvt ltd', 'ipc Healthcare', 'IPC'],
    'Parnax Lab':                    ['PARNAX', 'Parnax Lab Ltd', 'Parnax Lab Ltd Ltd',
                                      'Parnax Lab Ltd.', 'Parnex lab', 'PARNEX'],
    'Pharmatec Pvt. Ltd.':           ['Pharmatec', 'Pharmatec Pvt Ltd',
                                      'Pharmatec Pvt. Ltd', 'Pharmatech', 'PHARMATEC'],
    'Pharmatrust Ltd':               ['Pharma trust', 'Pharmatrust ltd', 'Pharmatrust limited',
                                      'Pharmatrust Limited', 'Pharmatrust Limited ltd',
                                      'PHARMATRUST'],
    'Socomed Pharma':                ['Socomed', 'Socomed pharma', 'Socomed pharma Pvt Ltd.',
                                      'SOCOMED'],
    'Bliss GVS':                     ['BLISS', 'Bliss GVS', 'Bliss GVS pharma Ltd'],
    'Careth Corporation':            ['Careth Corporation', 'Careth corporation'],
    'Shalina Laboratories Pvt Ltd':  ['Salina', 'Shalina', 'Shalina Laboratories Pvt Ltd.',
                                      'Shalina  Laboratories Pvt Ltd.', 'SHALINA'],
    'UC Rebok Investment Ltd':       ['UC Rebok Investment', 'UC-rebok investment ltd',
                                      'UC-Rebok investment Ltd.', 'UC-Rebok investment ltd',
                                      'UC-REBOK'],
    'Blue Map Pharmachem':           ['Blue Map Pharmachem', 'BLUE MAP', 'Blue Map'],
    'Kanvid Pharmaceutical':         ['Kanvid Pharmaceutical'],
    'Unique Pharma':                 ['Unique Pharma', 'Unique pharma', 'UNIQUE'],
    'Workcell Solution':             ['Workcell Solution', 'WORKCELL SOLUTION'],
    'Alvita Pharma Pvt Ltd':         ['Alvita pharma p ltd', 'Alvita Pharma',
                                      'Alvita Pharma Pvt Ltd.', 'Alvita Pharma Pvt.Ltd',
                                      'ALVITA'],
    'Indoco Remedies':               ['Indoco Remedies', 'INDOCO', 'Indoco Remedies Ltd.'],
    'Kamal':                         ['KAMAL'],
    'Enicar Pharmaceutical Pvt Ltd': ['Enicar Pharmaceutical Pvt Ltd',
                                      'Enicar Pharmaceuticals pvt ltd', 'ENICAR'],
    'London United Exports Ltd':     ['London United', 'London United Exports Ltd', 'LONDON', 'LONEDON'],
    'Yellow & Orange Pharmacy':      ['YELLOW ORANGE', 'Yellow & Orange Pharmacy',
                                      'Yellow & Orange Pharmacy.'],
    'Tushu Pharma':                  ['TUSHU', 'Tushu Pharma'],
    'Unichem Ghana':                 ['UNICHEM', 'Unichem Ghana'],
    'Nelpa Lifescience':             ['NELPA', 'Nelpa Lifescience'],
    'Corporsano Mediaid':            ['Corporsano', 'Corporsano Mediaid',
                                      'Corporsano Media id', 'CORPORSANO',
                                      'Corposano', 'CORPOSANO'],
    'Aura Pharmaceuticals Pvt Ltd':  ['Aura', 'Aura Pharma', 'Aura Pharmaceuticals',
                                      'Aura pharmaceuticals Pvt Ltd',
                                      'Aura pharmaceticals Pvt Ltd', 'AURA'],
    'Cariesco Exports':              ['Cariesco', 'Cariesco Exports', 'CARIESCO'],
    'Cedar Point Chemist Ltd':       ['Cedar Point', 'Cedar point chemist limited',
                                      'Cedar Point Chemist Ltd', 'Cedar Point Chemist Limited',
                                      'CEDAR'],
}

def _pkey(s):
    """Loose key for matching: lowercase, punctuation→space, single spaces."""
    return ' '.join(re.sub(r'[.\-,/]', ' ', str(s).lower()).split())

_PARTY_LOOKUP = {}
for _canon, _variants in _PARTY_GROUPS.items():
    _PARTY_LOOKUP[_pkey(_canon)] = _canon
    for _v in _variants:
        _PARTY_LOOKUP[_pkey(_v)] = _canon

def normalise_party(s):
    if pd.isna(s): return s
    raw = ' '.join(str(s).strip().split())
    return _PARTY_LOOKUP.get(_pkey(raw), raw)

fill_df['Line'] = fill_df['Line'].apply(normalise_line)
pack_df['Line'] = pack_df['Line'].apply(normalise_line)

disp_df = read_log('➕ Dispatch Log', 'B:I',
    ['Date','Product','PackSize','ProductType','Qty','Batch','Party','Remarks'],
    ['Qty'])

staff_df = read_log('➕ Staff Log', 'B:F',
    ['Date','Total','Female','Male','Remarks'],
    ['Total','Female','Male'])

# Merge duplicate customer spellings via the alias list above.
for _df in (fill_df, pack_df, disp_df):
    if 'Party' in _df.columns:
        _df['Party'] = _df['Party'].apply(normalise_party)

# ══════════════════════════════════════════════════════════════════════════════
# PERIOD SETUP
# ══════════════════════════════════════════════════════════════════════════════
# Period spans the latest TWO months whenever both have data (so we don't lose
# the previous month right after a month rollover). When only one month exists,
# we just show that one. PERIOD_MONTHS is a sorted list of (year, month) tuples.
def _months_with_data():
    months = set()
    for sheet, uc in [('➕ Filling Log','B:J'),
                      ('➕ Packing Log','B:N'),
                      ('➕ Dispatch Log','B:I')]:
        try:
            _df = pd.read_excel(TEMPLATE, sheet_name=sheet, header=3, usecols=uc)
            _d = pd.to_datetime(_df.iloc[:,0], format='mixed', dayfirst=True, errors='coerce').dropna()
            # Guard against typos like '30.050203' that parse to 1970 — only
            # accept dates from 2024 onwards.
            for ts in _d:
                if ts.year >= 2024:
                    months.add((ts.year, ts.month))
        except Exception:
            pass
    return sorted(months)

if len(sys.argv) > 2:
    PERIOD_MONTHS = [(YEAR, MONTH)]            # explicit user override
else:
    mm = _months_with_data()
    PERIOD_MONTHS = mm[-2:] if len(mm) >= 2 else (mm or [(today.year, today.month)])

# month_start / month_end now span the FULL period (could be 1 or 2 months).
_y0, _m0 = PERIOD_MONTHS[0]
_y1, _m1 = PERIOD_MONTHS[-1]
month_start = date(_y0, _m0, 1)
month_end   = date(_y1, _m1, calendar.monthrange(_y1, _m1)[1])

# Previous comparison window = same length immediately before month_start.
from datetime import timedelta
_window_days = (month_end - month_start).days + 1
prev_end   = month_start - timedelta(days=1)
prev_start = prev_end - timedelta(days=_window_days - 1)

if len(PERIOD_MONTHS) == 1:
    PERIOD = month_start.strftime('%B %Y').upper()
else:
    PERIOD = (month_start.strftime('%b %Y').upper() + ' – '
              + date(_y1, _m1, 1).strftime('%b %Y').upper())
PREV = prev_start.strftime('%b %Y')

# Keep YEAR/MONTH pointing at the latest month for downstream code that still
# references them (e.g. JS year/month labels).
YEAR, MONTH = _y1, _m1

def filt(df, s, e): return df[(df['Date'] >= s) & (df['Date'] <= e)]
def cur(df):        return filt(df, month_start, month_end)
def prv(df):        return filt(df, prev_start,  prev_end)

# ══════════════════════════════════════════════════════════════════════════════
# KPI CALCULATIONS
# ══════════════════════════════════════════════════════════════════════════════

# — Filling
f_cur   = cur(fill_df)['Qty'].sum()
f_prv   = prv(fill_df)['Qty'].sum()
f_rec   = len(cur(fill_df))
f_avg   = f_cur / f_rec if f_rec else 0
f_mom   = (f_cur - f_prv) / f_prv if f_prv else 0
f_lines = cur(fill_df)['Line'].nunique()

# Derive lines dynamically from actual data (both current and previous month)
LINES = sorted(set(fill_df['Line'].dropna().unique()) | set(pack_df['Line'].dropna().unique()))
# Filter out blank/NaN line names
LINES = [l for l in LINES if str(l).strip() not in ('', 'nan')]

fill_by_line = {
    ln: (cur(fill_df[fill_df['Line']==ln])['Qty'].sum(),
         prv(fill_df[fill_df['Line']==ln])['Qty'].sum())
    for ln in LINES
}

# — Packing
p_cur   = cur(pack_df)['TotalPacked'].sum()
p_prv   = prv(pack_df)['TotalPacked'].sum()
p_auto  = cur(pack_df)['AutoCarton'].sum()
p_man   = cur(pack_df)['ManualCarton'].sum()
p_mom   = (p_cur - p_prv) / p_prv if p_prv else 0
p_ratio = p_cur / f_cur if f_cur else 0

pack_by_line = {
    ln: (cur(pack_df[pack_df['Line']==ln])['TotalPacked'].sum(),
         prv(pack_df[pack_df['Line']==ln])['TotalPacked'].sum())
    for ln in LINES
}

# — Dispatch & BSR
d_cur        = cur(disp_df)['Qty'].sum()
d_prv        = prv(disp_df)['Qty'].sum()
d_mom        = (d_cur - d_prv) / d_prv if d_prv else 0
d_all        = disp_df['Qty'].sum()
f_all        = fill_df['Qty'].sum()
bsr_stock    = BSR_OPENING + f_all - d_all
bsr_pending  = f_cur - d_cur
d_fill_ratio = d_cur / f_cur if f_cur else 0

# — Staff
sc = cur(staff_df)
s_fem  = sc['Female'].mean() if len(sc) else 0
s_male = sc['Male'].mean()   if len(sc) else 0

# — Party-wise (dynamic from actual dispatch data)
party_cur = cur(disp_df).groupby('Party')['Qty'].sum().sort_values(ascending=False)
party_prv = prv(disp_df).groupby('Party')['Qty'].sum()
# All parties that appear in current or previous month
PARTIES = sorted(set(party_cur.index) | set(party_prv.index))
PARTIES = [p for p in PARTIES if str(p).strip() not in ('', 'nan')]

# — Product type breakdown (Filling, Packing, Dispatch)
PRODUCT_TYPES = ['Bottle', 'Flat Sachet', 'Stick Pack Sachet', 'Ointment', 'External']

# Normalise product type into the five canonical categories above.
#   • Flat Sachet        ← sachet / flat sachet / pouch variants
#   • Stick Pack Sachet  ← stick pack variants
#   • Ointment           ← ointment / tube variants
_FLAT_VARIANTS     = {'sachet', 'sachets', 'flat sachet', 'flat sachets',
                      'pouch', 'pouch/sachet', 'pouch/sachets'}
_STICK_VARIANTS    = {'stick pack', 'stick-pack', 'stickpack',
                      'stick pack sachet', 'stick-pack sachet', 'stickpack sachet'}
_OINTMENT_VARIANTS = {'ointment', 'ointments', 'tube', 'tubes'}
def normalise_pt(v):
    if pd.isna(v): return v
    s = str(v).strip()
    l = s.lower()
    if l in _STICK_VARIANTS:    return 'Stick Pack Sachet'
    if l in _FLAT_VARIANTS:     return 'Flat Sachet'
    if l in _OINTMENT_VARIANTS: return 'Ointment'
    return s.title()   # "bottle" → "Bottle", "external" → "External"

fill_df['ProductType'] = fill_df['ProductType'].apply(normalise_pt)
pack_df['ProdType']    = pack_df['ProdType'].apply(normalise_pt)
disp_df['ProductType'] = disp_df['ProductType'].apply(normalise_pt)

fill_by_type = cur(fill_df).groupby('ProductType')['Qty'].sum()
pack_by_type = cur(pack_df).groupby('ProdType')['TotalPacked'].sum()
disp_by_type = cur(disp_df).groupby('ProductType')['Qty'].sum()

# ══════════════════════════════════════════════════════════════════════════════
# BATCH JOURNEY  (Filling → Packing → Dispatch traceability, lifetime)
# ──────────────────────────────────────────────────────────────────────────────
# Batch number is the thread linking the three logs. For every batch we total
# how much was filled, packed and dispatched, and assign a clear status.
# A frozen baseline (batch_baseline.json) holds the "opening stock" batches that
# were made BEFORE filling-log tracking began — those are expected to have no
# filling record and are never flagged as a problem.
# Designed to extend: when Production Plan + Raw Material Dispatch arrive, they
# become two more stages keyed on the same batch number.
# ══════════════════════════════════════════════════════════════════════════════
import json, re as _re_bj

def _bkey(b):
    """Batch key — uppercase, no spaces (matches the typo checker)."""
    return _re_bj.sub(r'\s+', '', str(b)).upper()

# Load the frozen opening-stock baseline (created once, never auto-rewritten).
_BASELINE_PATH = os.path.join(HERE, 'batch_baseline.json')
try:
    OPENING_STOCK = set(json.load(open(_BASELINE_PATH)).get('batches', []))
except Exception:
    OPENING_STOCK = set()
    print('  Note: batch_baseline.json not found — opening stock not classified.')

def _batch_journey():
    """Return list of per-batch dicts with filled/packed/dispatched + status."""
    j = {}
    def add(df, qty_col, role, ptype_col):
        for _, r in df.iterrows():
            b = r.get('Batch')
            if pd.isna(b) or not str(b).strip():
                continue
            k = _bkey(b)
            e = j.setdefault(k, {'batch': str(b).strip(), 'product': None, 'ptype': None,
                                 'party': None, 'packsize': None,
                                 'filled': 0.0, 'packed': 0.0, 'dispatched': 0.0,
                                 'last': None, 'last_disp': None,
                                 'auto_cleared': False, 'leftover_cleared': 0.0})
            e[role] += float(r.get(qty_col) or 0)
            if e['product'] is None and not pd.isna(r.get('Product')):
                e['product'] = str(r.get('Product')).strip()
            if e['ptype'] is None and not pd.isna(r.get(ptype_col)):
                e['ptype'] = str(r.get(ptype_col)).strip()
            if e['party'] is None and 'Party' in r.index and not pd.isna(r.get('Party')):
                e['party'] = str(r.get('Party')).strip()
            if e['packsize'] is None and 'PackSize' in r.index and not pd.isna(r.get('PackSize')):
                e['packsize'] = str(r.get('PackSize')).strip()
            d = r.get('Date')
            if d is not None and (e['last'] is None or d > e['last']):
                e['last'] = d
            # Track most recent DISPATCH date separately — used for auto-clearing
            # tiny leftover stock 60+ days after the last dispatch.
            if role == 'dispatched' and d is not None and (e['last_disp'] is None or d > e['last_disp']):
                e['last_disp'] = d
    # Packing first so packed items take their product type / party from the packing log.
    add(pack_df, 'TotalPacked','packed',     'ProdType')
    add(fill_df, 'Qty',        'filled',     'ProductType')
    add(disp_df, 'Qty',        'dispatched', 'ProductType')

    # Auto-clear rule: small leftover (≤500 units) sitting >60 days after the
    # last dispatch is treated as effectively gone (samples, shrinkage, returns,
    # etc.) so the BSR stock total and Batch Journey don't carry stale remnants
    # forever. Threshold tuned to cover the typical 100–300 bottle leftover.
    AUTO_CLEAR_UNITS = 500
    AUTO_CLEAR_DAYS  = 60
    today_d = date.today()
    for e in j.values():
        leftover = e['filled'] - e['dispatched']
        if (e['last_disp'] is not None and 0 < leftover <= AUTO_CLEAR_UNITS
                and (today_d - e['last_disp']).days >= AUTO_CLEAR_DAYS):
            e['auto_cleared']     = True
            e['leftover_cleared'] = leftover

    for k, e in j.items():
        f, p, d = e['filled'], e['packed'], e['dispatched']
        in_fill = f > 0
        if k in OPENING_STOCK:                   # frozen pre-tracking stock — never flag
            e['status'], e['rank'] = ('Opening stock (pre-system)', 4)
        elif not in_fill:
            e['status'], e['rank'] = ('⚠ Dispatched/packed but never filled', 0)
        elif d > f * 1.02 and d - f > 1:         # shipped clearly more than made
            e['status'], e['rank'] = ('⚠ Dispatched more than filled', 0)
        elif e['auto_cleared']:
            e['status'], e['rank'] = (
                f"✅ Complete (small {int(e['leftover_cleared']):,}-unit leftover auto-cleared after 60d)", 3)
        elif d > 0:
            e['status'], e['rank'] = ('Filled → Packed → Dispatched', 3)
        elif p > 0:
            e['status'], e['rank'] = ('Filled & packed (in stock)', 2)
        else:
            e['status'], e['rank'] = ('Filled only', 2)
    # Problems first (rank 0), then most recent activity within each group
    return sorted(j.values(),
                  key=lambda e: (e['rank'], -(e['last'].toordinal() if e['last'] else 0)))

BATCH_JOURNEY = _batch_journey()

# Subtract auto-cleared small leftovers from the BSR stock total so old
# remnants stop inflating "in stock" forever. Logged on the dashboard side
# only (the underlying logs are not modified).
AUTO_CLEARED_TOTAL = sum(e.get('leftover_cleared', 0) for e in BATCH_JOURNEY)
AUTO_CLEARED_COUNT = sum(1 for e in BATCH_JOURNEY if e.get('auto_cleared'))
if AUTO_CLEARED_TOTAL:
    bsr_stock = bsr_stock - AUTO_CLEARED_TOTAL
    print(f"   Auto-cleared {AUTO_CLEARED_COUNT} batch(es) "
          f"({int(AUTO_CLEARED_TOTAL):,} units of small leftover stock) from BSR.")
_bj_problems = sum(1 for e in BATCH_JOURNEY if e['rank'] == 0)
_bj_flowing  = sum(1 for e in BATCH_JOURNEY if e['status'].startswith('Filled → Packed'))
_bj_stock    = sum(1 for e in BATCH_JOURNEY if e['rank'] == 2)
_bj_opening  = sum(1 for e in BATCH_JOURNEY if e['rank'] == 4)

# ── Monthly summary (RM → Fill → Pack → Disp) ─────────────────────────────────
# Tracking-start cutoff: dispenses before this date were from the pre-system
# era and are excluded so they don't pollute the monthly numbers.
TRACKING_START = date(2026, 5, 11)

def _monthly_summary():
    try:
        rm_all = pd.read_excel(TEMPLATE, sheet_name='➕ RM Dispensing Log', header=3)
    except Exception:
        return []
    if 'PLAN' not in rm_all.columns or 'DISPENSING DATE' not in rm_all.columns:
        return []
    rm = rm_all[rm_all['DISPENSING DATE'].notna()
                & (rm_all['PLAN'].astype(str).str.strip().str.upper() == 'REGULAR')
                & ~rm_all['BATCH NUMBER'].astype(str).str.upper().str.contains('TLB', na=False)].copy()
    rm['_d'] = pd.to_datetime(rm['DISPENSING DATE'], errors='coerce')
    # Tracking-start cutoff (May 11) — drop earlier dispenses
    rm = rm[rm['_d'].dt.date >= TRACKING_START]
    rm['_m'] = rm['_d'].dt.to_period('M').astype(str)
    rm['_k'] = rm['BATCH NUMBER'].astype(str).apply(_bkey)
    # batch_key → its RM dispense month (so we can ask "this June filling came from which dispense month?")
    rm_month_by_key = dict(zip(rm['_k'], rm['_m']))

    fkey_all = set(fill_df['Batch'].dropna().apply(lambda x: _bkey(x)))
    pkey_all = set(pack_df['Batch'].dropna().apply(lambda x: _bkey(x)))
    dkey_all = set(disp_df['Batch'].dropna().apply(lambda x: _bkey(x)))

    fill_with_m = fill_df.copy()
    fill_with_m['_m'] = pd.to_datetime(fill_with_m['Date'], errors='coerce').apply(
        lambda d: d.strftime('%Y-%m') if pd.notna(d) and hasattr(d, 'strftime') else None)
    pack_with_m = pack_df.copy()
    pack_with_m['_m'] = pd.to_datetime(pack_with_m['Date'], errors='coerce').apply(
        lambda d: d.strftime('%Y-%m') if pd.notna(d) and hasattr(d, 'strftime') else None)
    disp_with_m = disp_df.copy()
    disp_with_m['_m'] = pd.to_datetime(disp_with_m['Date'], errors='coerce').apply(
        lambda d: d.strftime('%Y-%m') if pd.notna(d) and hasattr(d, 'strftime') else None)

    months = sorted(set(rm['_m'].dropna()) | set(fill_with_m['_m'].dropna())
                    | set(pack_with_m['_m'].dropna()) | set(disp_with_m['_m'].dropna()))
    months = [m for m in months if m >= '2026-05']

    def _split_by_rm_month(df, qty_col, this_month):
        """Split a production month's rows by which RM-dispensing month each batch came from."""
        from_this   = {'b': set(), 'u': 0.0}
        from_prev   = {'b': set(), 'u': 0.0}
        from_other  = {'b': set(), 'u': 0.0}    # earlier than previous, or unknown RM
        for _, r in df.iterrows():
            b = r.get('Batch')
            if pd.isna(b) or not str(b).strip(): continue
            k = _bkey(b)
            q = float(pd.to_numeric(r.get(qty_col), errors='coerce') or 0)
            rm_m = rm_month_by_key.get(k)
            if rm_m == this_month:
                from_this['b'].add(k); from_this['u'] += q
            elif rm_m is not None and rm_m < this_month:
                from_prev['b'].add(k); from_prev['u'] += q
            else:
                from_other['b'].add(k); from_other['u'] += q
        return {
            'this':  {'b': len(from_this['b']),  'u': from_this['u']},
            'prev':  {'b': len(from_prev['b']),  'u': from_prev['u']},
            'other': {'b': len(from_other['b']), 'u': from_other['u']},
        }

    out = []
    for m in months:
        rmk = set(rm[rm['_m']==m]['_k'].dropna())
        f_m = fill_with_m[fill_with_m['_m']==m]
        p_m = pack_with_m[pack_with_m['_m']==m]
        d_m = disp_with_m[disp_with_m['_m']==m]

        # Product type breakdown (overall for this month)
        pt_table = {}
        for pt in PRODUCT_TYPES:
            pt_table[pt] = {
                'f': float(pd.to_numeric(f_m[f_m['ProductType']==pt]['Qty'], errors='coerce').sum()),
                'p': float(p_m[p_m['ProdType']==pt]['TotalPacked'].sum()),
                'd': float(pd.to_numeric(d_m[d_m['ProductType']==pt]['Qty'], errors='coerce').sum()),
            }

        # Per customer × product type (party names already alias-normalised on load)
        cust_data = {}
        for df_x, qty_col, ptype_col, stage in [
            (f_m, 'Qty',         'ProductType', 'f'),
            (p_m, 'TotalPacked', 'ProdType',    'p'),
            (d_m, 'Qty',         'ProductType', 'd'),
        ]:
            for _, r in df_x.iterrows():
                cust = r.get('Party')
                if pd.isna(cust) or not str(cust).strip(): continue
                cust = str(cust).strip()
                pt = r.get(ptype_col)
                if pd.isna(pt) or pt not in PRODUCT_TYPES: continue
                q = float(pd.to_numeric(r.get(qty_col), errors='coerce') or 0)
                cust_data.setdefault(cust, {p:{'f':0.0,'p':0.0,'d':0.0} for p in PRODUCT_TYPES})
                cust_data[cust][pt][stage] += q
        # Sort customers by total volume desc, keep only non-empty
        def _cust_total(c):
            return sum(cust_data[c][pt]['f']+cust_data[c][pt]['p']+cust_data[c][pt]['d'] for pt in PRODUCT_TYPES)
        customers = sorted([c for c in cust_data if _cust_total(c) > 0], key=_cust_total, reverse=True)

        out.append({
            'm':         m,
            'rm_b':      len(rmk),
            'f_b':       len(set(f_m['Batch'].dropna().apply(_bkey))),
            'p_b':       len(set(p_m['Batch'].dropna().apply(_bkey))),
            'd_b':       len(set(d_m['Batch'].dropna().apply(_bkey))),
            'f_u':       float(pd.to_numeric(f_m.get('Qty', pd.Series()), errors='coerce').sum()),
            'p_u':       float(p_m.get('TotalPacked', pd.Series()).sum()),
            'd_u':       float(pd.to_numeric(d_m.get('Qty', pd.Series()), errors='coerce').sum()),
            'bf_filled': sum(1 for k in rmk if k in fkey_all),
            'bf_packed': sum(1 for k in rmk if k in pkey_all),
            'bf_disp':   sum(1 for k in rmk if k in dkey_all),
            'bf_pending':sum(1 for k in rmk if k not in fkey_all),
            'fill_split': _split_by_rm_month(f_m, 'Qty',         m),
            'pack_split': _split_by_rm_month(p_m, 'TotalPacked', m),
            'disp_split': _split_by_rm_month(d_m, 'Qty',         m),
            'pt_table':   pt_table,
            'customers':  customers,
            'cust_data':  cust_data,
        })
    return out

MONTHLY_SUMMARY = _monthly_summary()

# ══════════════════════════════════════════════════════════════════════════════
# BUILD JSON DATA FOR JS DAY FILTER
# ══════════════════════════════════════════════════════════════════════════════

def safe(v):
    if v is None: return None
    try:
        import math
        if isinstance(v, float) and math.isnan(v): return None
    except: pass
    if hasattr(v, 'strftime'): return v.strftime('%Y-%m-%d')
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip()
    return None if s in ('', 'nan', 'None') else s

fill_rows = [{'date':safe(r['Date']),'line':safe(r['Line']),'product':safe(r['Product']),
              'packSize':safe(r['PackSize']),'productType':safe(r['ProductType']),
              'qty':safe(r['Qty']),'batch':safe(r['Batch']),'party':safe(r['Party'])}
             for _,r in cur(fill_df).iterrows()]

pack_rows = [{'date':safe(r['Date']),'line':safe(r['Line']),'product':safe(r['Product']),
              'packSize':safe(r['PackSize']),'productType':safe(r['ProdType']),
              'batch':safe(r['Batch']),'totalPacked':safe(r['TotalPacked']),'party':safe(r['Party'])}
             for _,r in cur(pack_df).iterrows()]

disp_rows = [{'date':safe(r['Date']),'product':safe(r['Product']),'packSize':safe(r['PackSize']),
              'productType':safe(r['ProductType']),'qty':safe(r['Qty']),
              'batch':safe(r['Batch']),'party':safe(r['Party'])}
             for _,r in cur(disp_df).iterrows()]

staff_rows = [{'date':safe(r['Date']),'female':safe(r['Female']),'male':safe(r['Male'])}
              for _,r in cur(staff_df).iterrows()]

batch_rows = [
    {'batch':e['batch'], 'product':e['product'], 'ptype':e['ptype'],
     'party':e.get('party'), 'packSize':e.get('packsize'),
     'filled':float(e['filled']), 'packed':float(e['packed']), 'dispatched':float(e['dispatched']),
     'status':e['status']}
    for e in BATCH_JOURNEY
]

DATA_JSON = json.dumps({
    'fill': fill_rows, 'pack': pack_rows, 'disp': disp_rows, 'staff': staff_rows,
    'lines': LINES, 'productTypes': PRODUCT_TYPES,
    'bsrOpening': BSR_OPENING, 'fillAll': float(f_all), 'dispAll': float(d_all),
    'batches': batch_rows
})


# ══════════════════════════════════════════════════════════════════════════════
# HTML HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def n(v):    return f'{int(v):,}'
def pct(v):  return f'{v*100:.1f}%'
def pctd(v): return f'{v*100:.2f}%'
def mom_badge(v):
    cls = 'badge-up' if v > 0 else ('badge-dn' if v < 0 else 'badge-fl')
    sym = '▲' if v > 0 else ('▼' if v < 0 else '—')
    return f'<span class="{cls}">{sym} {abs(v)*100:.1f}%</span>'

def tile(label, value, sub, color=C_AMB):
    return f'''<div class="tile">
      <div class="tlabel">{label}</div>
      <div class="tvalue" style="color:{color}">{value}</div>
      <div class="tsub">{sub}</div>
    </div>'''

def sec(title, bg=C_PRI):
    return f'<div class="sec-hdr" style="background:{bg}">{title}</div>'

def product_type_rows():
    rows = ''
    for i, pt in enumerate(PRODUCT_TYPES):
        fv = fill_by_type.get(pt, 0)
        pv = pack_by_type.get(pt, 0)
        dv = disp_by_type.get(pt, 0)
        # Always show all product types, even if 0
        bg = '#F1F8F6' if i % 2 == 0 else '#FFFFFF'
        rows += (
            f'<tr style="background:{bg}">'
            f'<td class="td-name">{pt}</td>'
            f'<td class="td-num" style="color:{C_SEC};font-weight:600">{n(fv)}</td>'
            f'<td class="td-num" style="color:{C_AMB};font-weight:600">{n(pv)}</td>'
            f'<td class="td-num" style="color:{C_ORG};font-weight:600">{n(dv)}</td>'
            f'</tr>'
        )
    return rows

# Packed and sitting in BSR stock — i.e. packed but NOT yet dispatched.
# Primary sort: party name (so all of one customer's stock is grouped together);
# then product type, then product, then batch.
def _ptype_rank(pt):
    return PRODUCT_TYPES.index(pt) if pt in PRODUCT_TYPES else len(PRODUCT_TYPES)
IN_STOCK = sorted(
    [e for e in BATCH_JOURNEY if e['packed'] > 0 and e['dispatched'] == 0],
    key=lambda e: ((e['party'] or 'zzz').lower(),
                   _ptype_rank(e['ptype']),
                   (e['product'] or '').lower(),
                   e['batch'])
)
IN_STOCK_UNITS = sum(e['packed'] for e in IN_STOCK)

def batch_journey_rows():
    rows = ''
    for i, e in enumerate(IN_STOCK):
        bg = '#F1F8F6' if i % 2 == 0 else '#FFFFFF'
        rows += (
            f'<tr style="background:{bg}">'
            f'<td class="td-name" style="font-weight:600">{e["party"] or "—"}</td>'
            f'<td class="td-name">{e["product"] or "—"}</td>'
            f'<td class="td-name" style="color:#607D8B">{e["ptype"] or "—"}</td>'
            f'<td class="td-name">{e["batch"]}</td>'
            f'<td class="td-num" style="color:{C_AMB};font-weight:700">{n(e["packed"])}</td>'
            f'</tr>'
        )
    return rows or '<tr><td colspan="5" style="text-align:center;color:#90A4AE;padding:12px">Nothing packed and waiting — all packed stock has been dispatched.</td></tr>'

def party_table_rows():
    rows = ''
    for i, p in enumerate(PARTIES):
        cv = party_cur.get(p, 0)
        pv = party_prv.get(p, 0)
        delta = cv - pv
        delta_s = ('+' if delta >= 0 else '') + n(delta)
        share = pct(cv / d_cur) if d_cur else '0.0%'
        bg = '#FFF8F1' if i % 2 == 0 else '#FFFFFF'
        trend_color = '#1B5E20' if cv >= pv else C_AMB
        trend = '▲' if cv > pv else ('▼' if cv < pv else '—')
        rows += (
            f'<tr style="background:{bg}">'
            f'<td class="td-name">{p}</td>'
            f'<td class="td-num" style="color:{C_ORG};font-weight:700">{n(cv)}</td>'
            f'<td class="td-num" style="color:#607D8B">{delta_s}</td>'
            f'<td class="td-num">{share}</td>'
            f'<td class="td-num" style="color:{trend_color};font-size:16px">{trend}</td>'
            f'</tr>'
        )
    return rows

def line_table_rows(data_dict, total_cur, total_prv):
    rows = ''
    for i, ln in enumerate(LINES):
        cur_v, prv_v = data_dict[ln]
        pct_v = cur_v / total_cur if total_cur else 0
        delta = cur_v - prv_v
        trend = '▲' if cur_v > prv_v else ('▼' if cur_v < prv_v else '—')
        tc = C_GRN if cur_v >= prv_v else C_AMB
        bg = '#F1F8F6' if i % 2 == 0 else '#FFFFFF'
        rows += f'''<tr style="background:{bg}">
          <td class="td-name">{ln}</td>
          <td class="td-num" style="color:{C_AMB};font-weight:700">{n(cur_v)}</td>
          <td class="td-num" style="color:#607D8B">{('+' if delta>=0 else '')}{n(delta)}</td>
          <td class="td-num">{pct(pct_v)}</td>
          <td class="td-num" style="color:{tc};font-size:16px">{trend}</td>
        </tr>'''
    return rows

# ══════════════════════════════════════════════════════════════════════════════
# ASSEMBLE HTML
# ══════════════════════════════════════════════════════════════════════════════
def monthly_summary_html():
    """Render the per-month pipeline glance — bottles funnel + batch funnel."""
    if not MONTHLY_SUMMARY:
        return ''
    blocks = ''
    for s in MONTHLY_SUMMARY:
        # Pretty month label
        try:
            yr, mo = s['m'].split('-')
            from datetime import date as _date
            label = _date(int(yr), int(mo), 1).strftime('%B %Y').upper()
        except Exception:
            label = s['m']
        f_u, p_u, d_u = s['f_u'], s['p_u'], s['d_u']
        pf = (p_u/f_u*100) if f_u else 0
        dp = (d_u/p_u*100) if p_u else 0
        delta = f_u - d_u
        # bar widths (relative to filled = 100%)
        wf = 100 if f_u else 0
        wp = (p_u/f_u*100) if f_u else 0
        wd = (d_u/f_u*100) if f_u else 0
        # batch funnel widths (relative to RM dispensed = 100%)
        rm_b = max(s['rm_b'], 1)
        bf_w_fill = s['bf_filled'] / rm_b * 100
        bf_w_pack = s['bf_packed'] / rm_b * 100
        bf_w_disp = s['bf_disp']   / rm_b * 100
        bf_w_pend = s['bf_pending']/ rm_b * 100

        blocks += f'''
        <div id="monthly-block-{s['m']}" class="monthly-block" style="background:#F1F8F6;border-radius:8px;padding:14px 16px;margin-bottom:14px">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px">
            <div style="font-weight:700;color:{C_PRI};font-size:16px">{label}</div>
            <div style="color:#607D8B;font-size:12px">In-stock delta: <strong style="color:{C_SEC}">{n(delta) if delta>=0 else '-'+n(-delta)} units</strong></div>
          </div>
          <!-- Bottle funnel (units) -->
          <table style="width:100%;font-size:13px;margin-bottom:8px">
            <tr><td style="width:80px;color:#546E7A">Filled</td>
                <td><div style="display:inline-block;background:{C_SEC};color:#fff;padding:3px 8px;border-radius:4px;width:{wf:.0f}%;box-sizing:border-box;min-width:80px">{n(f_u)} units &nbsp;|&nbsp; {s['f_b']} batches</div></td></tr>
            <tr><td style="color:#546E7A">Packed</td>
                <td><div style="display:inline-block;background:{C_AMB};color:#fff;padding:3px 8px;border-radius:4px;width:{wp:.0f}%;box-sizing:border-box;min-width:80px">{n(p_u)} units &nbsp;|&nbsp; {s['p_b']} batches &nbsp;<span style="opacity:.85">({pf:.1f}% of Filled)</span></div></td></tr>
            <tr><td style="color:#546E7A">Dispatched</td>
                <td><div style="display:inline-block;background:{C_ORG};color:#fff;padding:3px 8px;border-radius:4px;width:{wd:.0f}%;box-sizing:border-box;min-width:80px">{n(d_u)} units &nbsp;|&nbsp; {s['d_b']} batches &nbsp;<span style="opacity:.85">({dp:.1f}% of Packed)</span></div></td></tr>
          </table>
          <!-- Batch funnel for this month's RM dispenses -->
          <div style="border-top:1px solid #B0BEC5;padding-top:8px;margin-top:8px;font-size:12px;color:#546E7A">
            <span style="font-weight:600;color:{C_PRI}">RM dispensed in this month — where they are now (by batch number):</span>
            <table style="width:100%;font-size:12px;margin-top:6px">
              <tr><td style="width:160px">RM dispensed</td>
                  <td><div style="display:inline-block;background:#90A4AE;color:#fff;padding:2px 6px;border-radius:3px;min-width:50px">{s['rm_b']} batches</div></td></tr>
              <tr><td>↳ Reached Filling</td>
                  <td><div style="display:inline-block;background:{C_SEC};color:#fff;padding:2px 6px;border-radius:3px;width:{bf_w_fill:.0f}%;min-width:50px;box-sizing:border-box">{s['bf_filled']} ({bf_w_fill:.0f}%)</div></td></tr>
              <tr><td>↳ Reached Packing</td>
                  <td><div style="display:inline-block;background:{C_AMB};color:#fff;padding:2px 6px;border-radius:3px;width:{bf_w_pack:.0f}%;min-width:50px;box-sizing:border-box">{s['bf_packed']} ({bf_w_pack:.0f}%)</div></td></tr>
              <tr><td>↳ Reached Dispatch</td>
                  <td><div style="display:inline-block;background:{C_ORG};color:#fff;padding:2px 6px;border-radius:3px;width:{bf_w_disp:.0f}%;min-width:50px;box-sizing:border-box">{s['bf_disp']} ({bf_w_disp:.0f}%)</div></td></tr>
              <tr><td>⚠ Still pending (no production yet)</td>
                  <td><div style="display:inline-block;background:#B71C1C;color:#fff;padding:2px 6px;border-radius:3px;width:{bf_w_pend:.0f}%;min-width:50px;box-sizing:border-box">{s['bf_pending']} ({bf_w_pend:.0f}%)</div></td></tr>
            </table>
          </div>

          <!-- Cross-month: this month's production split by which RM-dispense month it came from -->
          <div style="border-top:1px solid #B0BEC5;padding-top:8px;margin-top:10px;font-size:12px;color:#546E7A">
            <span style="font-weight:600;color:{C_PRI}">Where {label.title()}'s production came from (RM dispense source):</span>
            <table style="width:100%;font-size:12px;margin-top:6px">
              <tr style="color:#90A4AE;font-size:11px"><td style="width:140px">&nbsp;</td><td>From THIS month's dispenses</td><td>From earlier-month dispenses</td><td>Other / unmatched</td></tr>
              <tr><td>Filled in {label.title()}</td>
                  <td><span style="color:{C_SEC};font-weight:600">{n(s['fill_split']['this']['u'])} u</span> <span style="color:#90A4AE">({s['fill_split']['this']['b']}b)</span></td>
                  <td><span style="color:{C_SEC};font-weight:600">{n(s['fill_split']['prev']['u'])} u</span> <span style="color:#90A4AE">({s['fill_split']['prev']['b']}b)</span></td>
                  <td><span style="color:#90A4AE">{n(s['fill_split']['other']['u'])} u ({s['fill_split']['other']['b']}b)</span></td></tr>
              <tr><td>Packed in {label.title()}</td>
                  <td><span style="color:{C_AMB};font-weight:600">{n(s['pack_split']['this']['u'])} u</span> <span style="color:#90A4AE">({s['pack_split']['this']['b']}b)</span></td>
                  <td><span style="color:{C_AMB};font-weight:600">{n(s['pack_split']['prev']['u'])} u</span> <span style="color:#90A4AE">({s['pack_split']['prev']['b']}b)</span></td>
                  <td><span style="color:#90A4AE">{n(s['pack_split']['other']['u'])} u ({s['pack_split']['other']['b']}b)</span></td></tr>
              <tr><td>Dispatched in {label.title()}</td>
                  <td><span style="color:{C_ORG};font-weight:600">{n(s['disp_split']['this']['u'])} u</span> <span style="color:#90A4AE">({s['disp_split']['this']['b']}b)</span></td>
                  <td><span style="color:{C_ORG};font-weight:600">{n(s['disp_split']['prev']['u'])} u</span> <span style="color:#90A4AE">({s['disp_split']['prev']['b']}b)</span></td>
                  <td><span style="color:#90A4AE">{n(s['disp_split']['other']['u'])} u ({s['disp_split']['other']['b']}b)</span></td></tr>
            </table>
            <div style="margin-top:4px;color:#90A4AE;font-size:11px">"Other / unmatched" = batches whose RM dispense isn't in the tracking window (pre-11-May or missing).</div>
          </div>
          {_pt_and_customer_html(s)}
        </div>
        '''
    return blocks

def _pt_and_customer_html(s):
    """Renders the Product Type breakdown table + per-customer breakdown for one month."""
    pt = s['pt_table']
    tot_f = sum(pt[p]['f'] for p in PRODUCT_TYPES)
    tot_p = sum(pt[p]['p'] for p in PRODUCT_TYPES)
    tot_d = sum(pt[p]['d'] for p in PRODUCT_TYPES)
    rows = ''
    for i, p in enumerate(PRODUCT_TYPES):
        bg = '#F1F8F6' if i % 2 == 0 else '#FFFFFF'
        rows += (f'<tr style="background:{bg}"><td class="td-name">{p}</td>'
                 f'<td class="td-num" style="color:{C_SEC};font-weight:600">{n(pt[p]["f"])}</td>'
                 f'<td class="td-num" style="color:{C_AMB};font-weight:600">{n(pt[p]["p"])}</td>'
                 f'<td class="td-num" style="color:{C_ORG};font-weight:600">{n(pt[p]["d"])}</td></tr>')

    pt_html = f'''
    <div style="border-top:1px solid #B0BEC5;padding-top:10px;margin-top:14px">
      <div style="font-weight:700;color:{C_PRI};font-size:13px;margin-bottom:6px">Product type breakdown — this month</div>
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead><tr style="background:{C_PRI};color:#fff">
          <th style="padding:6px 8px;text-align:left">PRODUCT TYPE</th>
          <th style="padding:6px 8px;text-align:right">UNITS FILLED</th>
          <th style="padding:6px 8px;text-align:right">UNITS PACKED</th>
          <th style="padding:6px 8px;text-align:right">UNITS DISPATCHED</th>
        </tr></thead>
        <tbody>{rows}</tbody>
        <tfoot><tr style="background:{C_LBG};font-weight:700">
          <td style="padding:6px 8px">TOTAL</td>
          <td class="td-num" style="color:{C_SEC}">{n(tot_f)}</td>
          <td class="td-num" style="color:{C_AMB}">{n(tot_p)}</td>
          <td class="td-num" style="color:{C_ORG}">{n(tot_d)}</td>
        </tr></tfoot>
      </table>
    </div>
    '''

    # Per-customer
    cust_blocks = ''
    for cust in s['customers']:
        cd = s['cust_data'][cust]
        # Build only non-empty rows
        nonzero_rows = []
        ctot_f = ctot_p = ctot_d = 0.0
        for p in PRODUCT_TYPES:
            f, pa, d = cd[p]['f'], cd[p]['p'], cd[p]['d']
            if f + pa + d == 0:
                continue
            nonzero_rows.append((p, f, pa, d))
            ctot_f += f; ctot_p += pa; ctot_d += d
        if not nonzero_rows:
            continue
        rows_html = ''
        for i, (p, f, pa, d) in enumerate(nonzero_rows):
            bg = '#FAFAFA' if i % 2 == 0 else '#FFFFFF'
            rows_html += (f'<tr style="background:{bg}"><td style="padding:3px 8px;color:#37474F">{p}</td>'
                          f'<td class="td-num" style="color:{C_SEC};padding:3px 8px">{n(f)}</td>'
                          f'<td class="td-num" style="color:{C_AMB};padding:3px 8px">{n(pa)}</td>'
                          f'<td class="td-num" style="color:{C_ORG};padding:3px 8px">{n(d)}</td></tr>')
        cust_blocks += f'''
        <div style="margin-top:10px;border:1px solid #ECEFF1;border-radius:6px;overflow:hidden">
          <div style="background:#ECEFF1;padding:6px 10px;font-weight:700;color:{C_PRI};font-size:13px">
            {cust} <span style="color:#90A4AE;font-weight:400;font-size:11px">— total F:{n(ctot_f)}  P:{n(ctot_p)}  D:{n(ctot_d)}</span>
          </div>
          <table style="width:100%;border-collapse:collapse;font-size:12px">
            <thead><tr style="background:#fff;color:#607D8B">
              <th style="padding:4px 8px;text-align:left">Product Type</th>
              <th style="padding:4px 8px;text-align:right">FILLED</th>
              <th style="padding:4px 8px;text-align:right">PACKED</th>
              <th style="padding:4px 8px;text-align:right">DISPATCHED</th>
            </tr></thead>
            <tbody>{rows_html}</tbody>
          </table>
        </div>
        '''

    cust_html = f'''
    <div style="margin-top:16px">
      <div style="font-weight:700;color:{C_PRI};font-size:13px;margin-bottom:4px">By customer × product type — this month</div>
      <div style="font-size:11px;color:#90A4AE;margin-bottom:8px">Sorted by total volume. Empty product types are hidden per customer.</div>
      {cust_blocks if cust_blocks else '<div style="color:#90A4AE;padding:8px">No customer activity this month.</div>'}
    </div>
    '''
    return pt_html + cust_html

generated_at = datetime.now().strftime('%d %b %Y, %I:%M %p')

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="900">
<title>Enicar Dashboard — {PERIOD}</title>
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:Arial,sans-serif; background:#ECEFF1; color:#263238; }}

  /* ── Header ── */
  .header {{ background:{C_PRI}; color:#fff; padding:18px 28px 12px; }}
  .header h1 {{ font-size:32px; font-weight:900; letter-spacing:3px; line-height:1; }}
  .header-sub {{ font-size:13px; color:#B2DFDB; margin-top:4px; }}
  .period-bar {{ background:{C_SEC}; color:#fff; text-align:center; padding:8px;
                 font-size:13px; font-weight:700; letter-spacing:1px; }}

  /* ── Layout ── */
  .container {{ max-width:1280px; margin:0 auto; padding:16px; }}
  .card {{ background:#fff; border-radius:10px; box-shadow:0 2px 8px rgba(0,0,0,0.09);
           margin-bottom:16px; overflow:hidden; }}

  /* ── Section header ── */
  .sec-hdr {{ color:#fff; font-size:13px; font-weight:700; padding:9px 16px;
              letter-spacing:1px; }}

  /* ── Tiles ── */
  .tile-row {{ display:flex; gap:10px; padding:14px 14px 10px; flex-wrap:wrap; }}
  .tile {{ flex:1; min-width:150px; background:#fff; border:1px solid #E0F2F1;
           border-radius:8px; padding:12px 14px; text-align:center; }}
  .tlabel {{ font-size:9px; font-weight:700; color:{C_SEC}; text-transform:uppercase;
             letter-spacing:0.8px; margin-bottom:6px; }}
  .tvalue {{ font-size:26px; font-weight:700; line-height:1.1; }}
  .tsub {{ font-size:9px; color:#90A4AE; margin-top:5px; font-style:italic; }}

  /* ── Tables ── */
  .tbl-wrap {{ padding:0 14px 14px; }}
  table {{ width:100%; border-collapse:collapse; font-size:12px; }}
  .th-row th {{ background:{C_SEC}; color:#fff; padding:8px 10px; font-size:10px;
                font-weight:700; letter-spacing:0.5px; }}
  .th-row th:first-child {{ text-align:left; padding-left:14px; }}
  tr td {{ padding:7px 10px; vertical-align:middle; }}
  .td-name {{ text-align:left; padding-left:14px; font-weight:500; color:#37474F; }}
  .td-num {{ text-align:center; }}
  .tot-row td {{ background:{C_SEC}!important; color:#fff; font-weight:700;
                 font-size:13px; padding:8px 10px; }}
  .tot-row .td-name {{ padding-left:14px; }}

  /* ── Badges ── */
  .badge-up {{ background:#E8F5E9; color:{C_GRN}; border-radius:4px;
               padding:2px 7px; font-size:11px; font-weight:700; }}
  .badge-dn {{ background:#FBE9E7; color:{C_AMB}; border-radius:4px;
               padding:2px 7px; font-size:11px; font-weight:700; }}
  .badge-fl {{ background:#ECEFF1; color:#78909C; border-radius:4px;
               padding:2px 7px; font-size:11px; font-weight:700; }}

  /* ── Charts ── */
  .chart-row {{ display:flex; gap:16px; padding:14px; }}
  .chart-box {{ flex:1; text-align:center; }}
  .chart-box img {{ width:100%; border-radius:6px; }}

  /* ── Footer ── */
  .footer {{ text-align:center; color:#90A4AE; font-size:10px; padding:20px; }}
  .bsr-note {{ background:#FFF3E0; border-left:4px solid {C_ORG};
               padding:8px 14px; font-size:11px; color:{C_ORG}; margin:4px 14px 10px; }}

  /* ── Date Filter Bar ── */
  .filter-bar {{ background:#fff; border-bottom:2px solid {C_LBG}; padding:10px 28px;
                 display:flex; align-items:center; gap:14px; flex-wrap:wrap; }}
  .filter-bar label {{ font-size:12px; font-weight:700; color:{C_SEC}; letter-spacing:0.5px; }}
  .filter-bar select {{ border:1.5px solid {C_SEC}; border-radius:6px; padding:6px 12px;
                        font-size:13px; color:#263238; background:#F9FAFB; cursor:pointer;
                        outline:none; }}
  .filter-bar select:focus {{ border-color:{C_PRI}; }}
  .filter-tag {{ background:{C_LBG}; color:{C_PRI}; border-radius:4px; padding:4px 10px;
                 font-size:11px; font-weight:700; letter-spacing:0.5px; }}
</style>
</head>
<body>

<!-- HEADER -->
<div class="header">
  <h1>ENICAR</h1>
  <div class="header-sub">PRODUCTION DASHBOARD &nbsp;|&nbsp; Generated {generated_at}</div>
</div>
<div class="period-bar">PRODUCTION &nbsp; DASHBOARD &nbsp;&nbsp;|&nbsp;&nbsp; {PERIOD}</div>

<div class="filter-bar">
  <label>📅 VIEW BY DATE:</label>
  <select id="date-filter" onchange="applyFilter()">
    <option value="all">All Month (MTD) — {PERIOD}</option>
  </select>
  <span class="filter-tag" id="filter-tag">MONTHLY TOTAL</span>
</div>

<div class="container">

<!-- ════════════════════════════════════════════════════════════
     SECTION 00 — MONTHLY SUMMARY (RM → Fill → Pack → Disp)
     Hidden by default — shown only when the date-filter is set to
     a "Monthly view: YYYY-MM" option. JS toggles individual blocks.
════════════════════════════════════════════════════════════ -->
<div class="card" id="monthly-summary-card" style="display:none">
  {sec('  ━━&nbsp;&nbsp;MONTHLY &nbsp; SUMMARY &nbsp; (RM &nbsp;→&nbsp; Fill &nbsp;→&nbsp; Pack &nbsp;→&nbsp; Dispatch) &nbsp;━━', C_PRI)}
  <div style="font-size:12px;color:#607D8B;padding:4px 4px 12px">
    Top: bottles/units produced in each month at each stage. Bottom: where this month's RM-dispensed
    <strong>batches</strong> are right now in the pipeline. May data starts from 11 May (tracking start).
  </div>
  {monthly_summary_html()}
</div>

<!-- ════════════════════════════════════════════════════════════
     SECTION 0 — BATCH / PRODUCT LOOKUP
════════════════════════════════════════════════════════════ -->
<div class="card">
  {sec('  ━━&nbsp;&nbsp;FIND &nbsp; A &nbsp; BATCH &nbsp; OR &nbsp; PRODUCT &nbsp;━━', C_SEC)}
  <div style="padding:6px 4px 10px">
    <input id="batch-search" type="text"
           placeholder="Type a batch number or product name (e.g. EL-2430, Bonaplex)…"
           oninput="lookupBatch()"
           style="width:100%;padding:10px;font-size:14px;border:1px solid #B0BEC5;border-radius:6px;box-sizing:border-box">
  </div>
  <div id="batch-search-results"></div>
</div>

<!-- ════════════════════════════════════════════════════════════
     SECTION 1 — PRODUCT TYPE BREAKDOWN
════════════════════════════════════════════════════════════ -->
<div class="card">
  {sec('  ━━&nbsp;&nbsp;PRODUCT &nbsp; TYPE &nbsp; BREAKDOWN &nbsp;━━')}
  <div class="tbl-wrap">
    <table>
      <tr class="th-row">
        <th>PRODUCT TYPE</th>
        <th id="pt-hdr-fill">UNITS FILLED (MTD)</th>
        <th id="pt-hdr-pack">UNITS PACKED (MTD)</th>
        <th id="pt-hdr-disp">UNITS DISPATCHED (MTD)</th>
      </tr>
      <tbody id="pt-rows"></tbody>
      <tr class="tot-row">
        <td class="td-name">TOTAL</td>
        <td class="td-num" id="pt-total-fill">—</td>
        <td class="td-num" id="pt-total-pack">—</td>
        <td class="td-num" id="pt-total-disp">—</td>
      </tr>
    </table>
  </div>
</div>

<!-- ════════════════════════════════════════════════════════════
     SECTION 2 — FILLING
════════════════════════════════════════════════════════════ -->
<div class="card">
  {sec('  ━━&nbsp;&nbsp;FILLING &nbsp; PRODUCTION &nbsp;━━')}
  <div class="tile-row">
    <div class="tile"><div class="tlabel">TOTAL FILLED</div><div class="tvalue" id="f-total" style="color:{C_AMB}">—</div><div class="tsub">units filled</div></div>
    <div class="tile"><div class="tlabel">FILL RECORDS</div><div class="tvalue" id="f-rec" style="color:{C_AMB}">—</div><div class="tsub">rows logged</div></div>
    <div class="tile"><div class="tlabel">AVG UNITS / RECORD</div><div class="tvalue" id="f-avg" style="color:{C_AMB}">—</div><div class="tsub">units per entry</div></div>
    <div class="tile"><div class="tlabel">ACTIVE LINES</div><div class="tvalue" id="f-lines" style="color:{C_AMB}">—</div><div class="tsub">lines active</div></div>
  </div>
  <div class="tbl-wrap">
    <table>
      <thead id="fill-thead"><tr class="th-row"><th>FILLING LINE</th><th>PRODUCT NAME</th><th>PACK SIZE</th><th>PARTY</th><th>UNITS FILLED (MTD)</th></tr></thead>
      <tbody id="fill-line-rows"></tbody>
      <tfoot id="fill-tfoot"><tr class="tot-row"><td class="td-name" colspan="4">TOTAL ALL LINES</td><td class="td-num">—</td></tr></tfoot>
    </table>
  </div>
</div>

<!-- ════════════════════════════════════════════════════════════
     SECTION 2 — PACKING
════════════════════════════════════════════════════════════ -->
<div class="card">
  {sec('  ━━&nbsp;&nbsp;PACKING &nbsp; PRODUCTION &nbsp;━━')}
  <div class="tile-row">
    <div class="tile"><div class="tlabel">TOTAL PACKED</div><div class="tvalue" id="p-total" style="color:{C_AMB}">—</div><div class="tsub">units packed</div></div>
    <div class="tile"><div class="tlabel">FILL → PACK RATIO</div><div class="tvalue" id="p-ratio" style="color:{C_AMB}">—</div><div class="tsub">packed ÷ filled</div></div>
  </div>
  <div class="tbl-wrap">
    <table>
      <thead id="pack-thead"><tr class="th-row"><th>PACKING LINE</th><th>PRODUCT NAME</th><th>PACK SIZE</th><th>PARTY</th><th>UNITS PACKED (MTD)</th></tr></thead>
      <tbody id="pack-line-rows"></tbody>
      <tfoot id="pack-tfoot"><tr class="tot-row"><td class="td-name" colspan="4">TOTAL ALL LINES</td><td class="td-num">—</td></tr></tfoot>
    </table>
  </div>
</div>

<!-- ════════════════════════════════════════════════════════════
     SECTION 3 — DISPATCH & BSR STOCK
════════════════════════════════════════════════════════════ -->
<div class="card">
  {sec('  ━━&nbsp;&nbsp;DISPATCH &nbsp;&amp;&nbsp; BSR &nbsp; STOCK &nbsp;━━', C_ORG)}
  <div class="tile-row">
    <div class="tile"><div class="tlabel">DISPATCHED</div><div class="tvalue" id="d-total" style="color:{C_ORG}">—</div><div class="tsub">units dispatched</div></div>
    <div class="tile"><div class="tlabel">DISPATCH / FILL</div><div class="tvalue" id="d-ratio" style="color:{C_ORG}">—</div><div class="tsub">dispatched ÷ filled</div></div>
  </div>
</div>

<!-- ════════════════════════════════════════════════════════════
     SECTION 4 — STAFF
════════════════════════════════════════════════════════════ -->
<div class="card">
  {sec('  ━━&nbsp;&nbsp;STAFF &nbsp;&amp;&nbsp; ATTENDANCE &nbsp;━━')}
  <div class="tile-row">
    <div class="tile"><div class="tlabel">FEMALE WORKERS PRESENT</div><div class="tvalue" id="s-fem" style="color:{C_AMB}">—</div><div class="tsub">packing workers</div></div>
    <div class="tile"><div class="tlabel">MALE WORKERS PRESENT</div><div class="tvalue" id="s-male" style="color:{C_AMB}">—</div><div class="tsub">filling & loading</div></div>
  </div>
</div>

<!-- ════════════════════════════════════════════════════════════
     SECTION 6 — PARTY-WISE SALES
════════════════════════════════════════════════════════════ -->
<div class="card">
  {sec('  ━━&nbsp;&nbsp;PARTY-WISE &nbsp; SALES &nbsp; (Dispatched) &nbsp;━━', C_ORG)}
  <div class="tbl-wrap">
    <table>
      <thead id="party-thead"><tr class="th-row"><th>PARTY NAME</th><th>DISPATCHED (MTD)</th></tr></thead>
      <tbody id="party-rows"></tbody>
      <tfoot id="party-tfoot"><tr class="tot-row"><td class="td-name">TOTAL ALL PARTIES</td><td class="td-num">—</td></tr></tfoot>
    </table>
  </div>
</div>

<!-- ════════════════════════════════════════════════════════════
     SECTION 7 — PACKED & IN STOCK (not yet dispatched)
════════════════════════════════════════════════════════════ -->
<div class="card">
  {sec('  ━━&nbsp;&nbsp;PACKED &nbsp;&amp;&nbsp; IN &nbsp; BSR &nbsp; STOCK &nbsp; (Not &nbsp; Yet &nbsp; Dispatched) &nbsp;━━', C_SEC)}
  <div class="tile-row">
    {tile('BATCHES IN STOCK', n(len(IN_STOCK)), 'packed, awaiting dispatch', C_SEC)}
    {tile('UNITS IN STOCK', n(IN_STOCK_UNITS), 'packed & not dispatched', C_AMB)}
  </div>
  <div class="tbl-wrap">
    <table>
      <thead><tr class="th-row">
        <th>PARTY</th><th>PRODUCT</th><th>PRODUCT TYPE</th><th>BATCH</th><th>QTY PACKED (IN STOCK)</th>
      </tr></thead>
      <tbody>{batch_journey_rows()}</tbody>
    </table>
  </div>
</div>

</div><!-- /container -->

<div class="footer">
  Generated by Enicar Dashboard Generator &nbsp;|&nbsp; {generated_at}<br>
  Data source: Enicar_Dashboard_Template.xlsx &nbsp;|&nbsp;
  Run <strong>Refresh_Dashboard.command</strong> to update with latest data
</div>

<script>
const ENICAR = {DATA_JSON};
const PT = ENICAR.productTypes;
const LINES = ENICAR.lines;

// ── Helpers ──────────────────────────────────────────
const fmt = v => Math.round(v).toLocaleString('en-IN');
const pct = (a,b) => b ? (a/b*100).toFixed(1)+'%' : '0.0%';

// Fixed display order for line breakdowns:
//   Line No 1..5 first (numeric), then Flat Sachet, Stick Pack Sachet,
//   Sachet, Ointment, External.
function lineOrderKey(name) {{
  const s = (name||'').toString().trim();
  const m = s.toLowerCase().match(/^line\\s*no\\.?\\s*0*(\\d+)/);
  if (m) return [0, parseInt(m[1],10)];
  const special = {{'flat sachet':1,'stick pack sachet':2,'sachet':3,'ointment':4,'external':5}};
  const k = special[s.toLowerCase()];
  return [1, k!==undefined ? k : 99];
}}
function cmpLine(a,b) {{
  const ka=lineOrderKey(a), kb=lineOrderKey(b);
  return ka[0]-kb[0] || ka[1]-kb[1] || String(a).localeCompare(String(b));
}}

// ── Build date dropdown ───────────────────────────────
(function buildDropdown() {{
  const sel = document.getElementById('date-filter');
  const monthsAbbr = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  // Monthly Summary options at the top — one per available month, served by the server
  const monthBlocks = document.querySelectorAll('#monthly-summary-card .monthly-block');
  monthBlocks.forEach(blk => {{
    const m = blk.id.replace('monthly-block-','');
    const [y, mo] = m.split('-');
    const opt = document.createElement('option');
    opt.value = 'monthly:' + m;
    opt.text = `📊 ${{monthsAbbr[parseInt(mo)]}} ${{y}} — Monthly Summary view`;
    sel.appendChild(opt);
  }});
  // Daily date options
  const dates = new Set();
  [...ENICAR.fill, ...ENICAR.pack, ...ENICAR.disp].forEach(r => {{ if(r.date) dates.add(r.date); }});
  [...dates].sort().reverse().forEach(d => {{
    const opt = document.createElement('option');
    opt.value = d;
    const [y,m,day] = d.split('-');
    opt.text = `${{parseInt(day)}} ${{monthsAbbr[parseInt(m)]}} ${{y}}`;
    sel.appendChild(opt);
  }});
}})();

// ── Main render ───────────────────────────────────────
function applyFilter() {{
  const sel = document.getElementById('date-filter').value;

  // Monthly-Summary view: show only the chosen month's block, hide everything else
  const isMonthlyView = sel.startsWith('monthly:');
  const summaryCard = document.getElementById('monthly-summary-card');
  document.querySelectorAll('.monthly-block').forEach(b => b.style.display = 'none');
  if (isMonthlyView) {{
    const m = sel.slice('monthly:'.length);
    summaryCard.style.display = '';
    const blk = document.getElementById('monthly-block-' + m);
    if (blk) blk.style.display = '';
    // Hide the other content cards (everything except the search + monthly summary)
    document.querySelectorAll('.container > .card').forEach(card => {{
      if (card.id !== 'monthly-summary-card'
          && !card.querySelector('#batch-search')) {{
        card.style.display = 'none';
      }} else {{
        card.style.display = '';
      }}
    }});
    document.getElementById('filter-tag').textContent = 'MONTHLY VIEW';
    return;
  }} else {{
    summaryCard.style.display = 'none';
    document.querySelectorAll('.container > .card').forEach(card => {{ card.style.display = ''; }});
  }}

  const isAll = sel === 'all';
  document.getElementById('filter-tag').textContent = isAll ? 'MONTHLY TOTAL' : 'DAILY VIEW';

  const fill  = isAll ? ENICAR.fill  : ENICAR.fill.filter(r => r.date === sel);
  const pack  = isAll ? ENICAR.pack  : ENICAR.pack.filter(r => r.date === sel);
  const disp  = isAll ? ENICAR.disp  : ENICAR.disp.filter(r => r.date === sel);
  const staff = isAll ? ENICAR.staff : ENICAR.staff.filter(r => r.date === sel);
  const label = isAll ? 'MTD' : 'TODAY';

  // Update product type breakdown column headers
  document.getElementById('pt-hdr-fill').textContent = `UNITS FILLED (${{label}})`;
  document.getElementById('pt-hdr-pack').textContent = `UNITS PACKED (${{label}})`;
  document.getElementById('pt-hdr-disp').textContent = `UNITS DISPATCHED (${{label}})`;

  renderProductTypes(fill, pack, disp);
  renderFilling(fill, isAll);
  renderPacking(pack, fill, isAll);
  renderDispatch(disp, fill);
  renderStaff(staff, isAll);
  renderParties(disp, isAll);
}}

// ── Product Type Breakdown ────────────────────────────
const _FLAT_JS     = new Set(['sachet','sachets','flat sachet','flat sachets','pouch','pouch/sachet','pouch/sachets']);
const _STICK_JS    = new Set(['stick pack','stick-pack','stickpack','stick pack sachet','stick-pack sachet','stickpack sachet']);
const _OINT_JS     = new Set(['ointment','ointments','tube','tubes']);
function normPT(pt) {{
  if (!pt) return pt;
  const l = pt.toLowerCase().trim();
  if (_STICK_JS.has(l)) return 'Stick Pack Sachet';
  if (_FLAT_JS.has(l))  return 'Flat Sachet';
  if (_OINT_JS.has(l))  return 'Ointment';
  return pt.trim().replace(/\\b\\w/g, c => c.toUpperCase());  // title case
}}
function renderProductTypes(fill, pack, disp) {{
  const fByType = {{}}; fill.forEach(r => {{ const k=normPT(r.productType); if(k) fByType[k]=(fByType[k]||0)+(r.qty||0); }});
  const pByType = {{}}; pack.forEach(r => {{ const k=normPT(r.productType); if(k) pByType[k]=(pByType[k]||0)+(r.totalPacked||0); }});
  const dByType = {{}}; disp.forEach(r => {{ const k=normPT(r.productType); if(k) dByType[k]=(dByType[k]||0)+(r.qty||0); }});
  const fTot = fill.reduce((s,r)=>s+(r.qty||0),0);
  const pTot = pack.reduce((s,r)=>s+(r.totalPacked||0),0);
  const dTot = disp.reduce((s,r)=>s+(r.qty||0),0);

  let rows = ''; let i = 0;
  PT.forEach(pt => {{
    const fv = fByType[pt]||0, pv = pByType[pt]||0, dv = dByType[pt]||0;
    const bg = i++%2===0 ? '#F1F8F6':'#fff';
    rows += `<tr style="background:${{bg}}"><td class="td-name">${{pt}}</td>
      <td class="td-num" style="color:#00695C;font-weight:600">${{fmt(fv)}}</td>
      <td class="td-num" style="color:#BF360C;font-weight:600">${{fmt(pv)}}</td>
      <td class="td-num" style="color:#E65100;font-weight:600">${{fmt(dv)}}</td></tr>`;
  }});
  document.getElementById('pt-rows').innerHTML = rows || '<tr><td colspan="4" style="text-align:center;color:#90A4AE;padding:12px">No data</td></tr>';
  document.getElementById('pt-total-fill').textContent = fmt(fTot);
  document.getElementById('pt-total-pack').textContent = fmt(pTot);
  document.getElementById('pt-total-disp').textContent = fmt(dTot);
}}

// ── Filling ───────────────────────────────────────────
function renderFilling(fill, isAll) {{
  const tot = fill.reduce((s,r)=>s+(r.qty||0),0);
  const rec = fill.length;
  const lines = new Set(fill.map(r=>r.line).filter(Boolean)).size;
  document.getElementById('f-total').textContent = fmt(tot);
  document.getElementById('f-rec').textContent = rec;
  document.getElementById('f-avg').textContent = rec ? fmt(tot/rec) : '0';
  document.getElementById('f-lines').textContent = lines;

  let rows = ''; let i = 0;
  const label = isAll ? 'MTD' : 'TODAY';
  // Group by line + product + packSize + party in both MTD and Daily views.
  document.getElementById('fill-thead').innerHTML = `<tr class="th-row"><th>FILLING LINE</th><th>PRODUCT NAME</th><th>PACK SIZE</th><th>PARTY</th><th>UNITS FILLED (${{label}})</th></tr>`;
  document.getElementById('fill-tfoot').innerHTML = `<tr class="tot-row"><td class="td-name" colspan="4">TOTAL ALL LINES</td><td class="td-num">${{fmt(tot)}}</td></tr>`;
  const byKey = {{}};
  fill.forEach(r => {{
    const ln = r.line||'—', pr = r.product||'—', ps = (r.packSize===null||r.packSize===undefined||r.packSize==='')?'—':String(r.packSize), pa = r.party||'—';
    const k = ln + '|||' + pr + '|||' + ps + '|||' + pa;
    if (!byKey[k]) byKey[k] = {{line:ln, product:pr, packSize:ps, party:pa, qty:0}};
    byKey[k].qty += (r.qty||0);
  }});
  Object.values(byKey).sort((a,b) => cmpLine(a.line,b.line)
                                  || a.product.localeCompare(b.product)
                                  || String(a.packSize).localeCompare(String(b.packSize))
                                  || a.party.localeCompare(b.party)).forEach(d => {{
    const bg = i++%2===0 ? '#F1F8F6':'#fff';
    rows += `<tr style="background:${{bg}}"><td class="td-name">${{d.line}}</td><td class="td-name">${{d.product}}</td><td class="td-name" style="color:#37474F">${{d.packSize}}</td><td class="td-name" style="color:#546E7A">${{d.party}}</td><td class="td-num" style="color:#BF360C;font-weight:700">${{fmt(d.qty)}}</td></tr>`;
  }});
  document.getElementById('fill-line-rows').innerHTML = rows || '<tr><td colspan="5" style="text-align:center;color:#90A4AE;padding:12px">No data</td></tr>';
}}

// ── Packing ───────────────────────────────────────────
function renderPacking(pack, fill, isAll) {{
  const tot = pack.reduce((s,r)=>s+(r.totalPacked||0),0);
  const fTot = fill.reduce((s,r)=>s+(r.qty||0),0);
  document.getElementById('p-total').textContent = fmt(tot);
  document.getElementById('p-ratio').textContent = fTot ? (tot/fTot*100).toFixed(1)+'%' : '—';

  let rows = ''; let i = 0;
  const label = isAll ? 'MTD' : 'TODAY';
  // Group by line + product + packSize + party in both MTD and Daily views.
  document.getElementById('pack-thead').innerHTML = `<tr class="th-row"><th>PACKING LINE</th><th>PRODUCT NAME</th><th>PACK SIZE</th><th>PARTY</th><th>UNITS PACKED (${{label}})</th></tr>`;
  document.getElementById('pack-tfoot').innerHTML = `<tr class="tot-row"><td class="td-name" colspan="4">TOTAL ALL LINES</td><td class="td-num">${{fmt(tot)}}</td></tr>`;
  const byKey = {{}};
  pack.forEach(r => {{
    const ln = r.line||'—', pr = r.product||'—', ps = (r.packSize===null||r.packSize===undefined||r.packSize==='')?'—':String(r.packSize), pa = r.party||'—';
    const k = ln + '|||' + pr + '|||' + ps + '|||' + pa;
    if (!byKey[k]) byKey[k] = {{line:ln, product:pr, packSize:ps, party:pa, qty:0}};
    byKey[k].qty += (r.totalPacked||0);
  }});
  Object.values(byKey).sort((a,b) => cmpLine(a.line,b.line)
                                  || a.product.localeCompare(b.product)
                                  || String(a.packSize).localeCompare(String(b.packSize))
                                  || a.party.localeCompare(b.party)).forEach(d => {{
    const bg = i++%2===0 ? '#F1F8F6':'#fff';
    rows += `<tr style="background:${{bg}}"><td class="td-name">${{d.line}}</td><td class="td-name">${{d.product}}</td><td class="td-name" style="color:#37474F">${{d.packSize}}</td><td class="td-name" style="color:#546E7A">${{d.party}}</td><td class="td-num" style="color:#BF360C;font-weight:700">${{fmt(d.qty)}}</td></tr>`;
  }});
  document.getElementById('pack-line-rows').innerHTML = rows || '<tr><td colspan="5" style="text-align:center;color:#90A4AE;padding:12px">No data</td></tr>';
}}

// ── Dispatch ──────────────────────────────────────────
function renderDispatch(disp, fill) {{
  const tot = disp.reduce((s,r)=>s+(r.qty||0),0);
  const fTot = fill.reduce((s,r)=>s+(r.qty||0),0);
  document.getElementById('d-total').textContent = fmt(tot);
  document.getElementById('d-ratio').textContent = fTot ? (tot/fTot*100).toFixed(1)+'%' : '—';
}}

// ── Staff ─────────────────────────────────────────────
function renderStaff(staff, isAll) {{
  if (!staff.length) {{
    document.getElementById('s-fem').textContent = '—';
    document.getElementById('s-male').textContent = '—';
    return;
  }}
  const fem  = staff.reduce((s,r)=>s+(r.female||0),0);
  const male = staff.reduce((s,r)=>s+(r.male||0),0);
  if (isAll) {{
    document.getElementById('s-fem').textContent  = (fem/staff.length).toFixed(0) + ' avg';
    document.getElementById('s-male').textContent = (male/staff.length).toFixed(0) + ' avg';
  }} else {{
    document.getElementById('s-fem').textContent  = fmt(fem);
    document.getElementById('s-male').textContent = fmt(male);
  }}
}}

// ── Party-wise ────────────────────────────────────────
function renderParties(disp, isAll) {{
  const tot = disp.reduce((s,r)=>s+(r.qty||0),0);
  let rows = ''; let i = 0;
  if (isAll) {{
    // MTD: group by party only
    document.getElementById('party-thead').innerHTML = '<tr class="th-row"><th>PARTY NAME</th><th>DISPATCHED (MTD)</th></tr>';
    document.getElementById('party-tfoot').innerHTML = `<tr class="tot-row"><td class="td-name">TOTAL ALL PARTIES</td><td class="td-num">${{fmt(tot)}}</td></tr>`;
    const byParty = {{}};
    disp.forEach(r => {{ if(r.party) byParty[r.party] = (byParty[r.party]||0)+(r.qty||0); }});
    Object.entries(byParty).sort((a,b)=>b[1]-a[1]).forEach(([p,v]) => {{
      const bg = i++%2===0 ? '#FFF8F1':'#fff';
      rows += `<tr style="background:${{bg}}"><td class="td-name">${{p}}</td><td class="td-num" style="color:#E65100;font-weight:700">${{fmt(v)}}</td></tr>`;
    }});
    document.getElementById('party-rows').innerHTML = rows || '<tr><td colspan="2" style="text-align:center;color:#90A4AE;padding:12px">No data</td></tr>';
  }} else {{
    // Daily: group by party + product
    document.getElementById('party-thead').innerHTML = '<tr class="th-row"><th>PARTY NAME</th><th>PRODUCT NAME</th><th>DISPATCHED (TODAY)</th></tr>';
    document.getElementById('party-tfoot').innerHTML = `<tr class="tot-row"><td class="td-name" colspan="2">TOTAL ALL PARTIES</td><td class="td-num">${{fmt(tot)}}</td></tr>`;
    const byPartyProd = {{}};
    disp.forEach(r => {{
      const k = (r.party||'—') + '|||' + (r.product||'—');
      if (!byPartyProd[k]) byPartyProd[k] = {{party:r.party||'—', product:r.product||'—', qty:0}};
      byPartyProd[k].qty += (r.qty||0);
    }});
    Object.values(byPartyProd).sort((a,b)=>b.qty-a.qty).forEach(d => {{
      const bg = i++%2===0 ? '#FFF8F1':'#fff';
      rows += `<tr style="background:${{bg}}"><td class="td-name">${{d.party}}</td><td class="td-name">${{d.product}}</td><td class="td-num" style="color:#E65100;font-weight:700">${{fmt(d.qty)}}</td></tr>`;
    }});
    document.getElementById('party-rows').innerHTML = rows || '<tr><td colspan="3" style="text-align:center;color:#90A4AE;padding:12px">No data</td></tr>';
  }}
}}

// ── Batch / product lookup ───────────────────────────
const BATCHES = ENICAR.batches || [];
function lookupBatch() {{
  const q = (document.getElementById('batch-search').value || '').trim().toLowerCase();
  const out = document.getElementById('batch-search-results');
  if (!q) {{
    out.innerHTML = '<div style="color:#90A4AE;padding:8px 4px;font-size:12px">Type a batch number or product name to see where it is in the pipeline (filled / packed / dispatched)…</div>';
    return;
  }}
  const hits = BATCHES.filter(b =>
    ((b.batch||'').toLowerCase().includes(q)) ||
    ((b.product||'').toLowerCase().includes(q))
  );
  if (!hits.length) {{
    out.innerHTML = `<div style="color:#90A4AE;padding:8px 4px">No batches or products match "${{q}}".</div>`;
    return;
  }}
  // Sort: matching product alphabetically, then batch
  hits.sort((a,b) => (a.product||'').localeCompare(b.product||'') || (a.batch||'').localeCompare(b.batch||''));
  let rows = '';
  hits.slice(0,80).forEach((b,i) => {{
    const bg = i%2===0 ? '#F1F8F6' : '#fff';
    rows += `<tr style="background:${{bg}}">
      <td class="td-name">${{b.product||'—'}}</td>
      <td class="td-name" style="color:#607D8B">${{b.ptype||'—'}}</td>
      <td class="td-name" style="color:#37474F">${{b.packSize||'—'}}</td>
      <td class="td-name" style="font-weight:600">${{b.batch}}</td>
      <td class="td-num" style="color:#00695C">${{fmt(b.filled)}}</td>
      <td class="td-num" style="color:#BF360C">${{fmt(b.packed)}}</td>
      <td class="td-num" style="color:#E65100">${{fmt(b.dispatched)}}</td>
      <td class="td-name" style="font-size:12px">${{b.status}}</td>
    </tr>`;
  }});
  const note = hits.length>80 ? `<div style="color:#90A4AE;font-size:11px;padding:4px">Showing first 80 of ${{hits.length}} matches — refine your search.</div>` : '';
  out.innerHTML = `<div class="tbl-wrap"><table>
    <thead><tr class="th-row">
      <th>PRODUCT</th><th>TYPE</th><th>PACK SIZE</th><th>BATCH</th><th>FILLED</th><th>PACKED</th><th>DISPATCHED</th><th>STATUS</th>
    </tr></thead>
    <tbody>${{rows}}</tbody>
  </table></div>${{note}}`;
}}

// ── Init on load ──────────────────────────────────────
lookupBatch();
applyFilter();
</script>

</body>
</html>"""

# ══════════════════════════════════════════════════════════════════════════════
# WRITE OUTPUT
# ══════════════════════════════════════════════════════════════════════════════
with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'✅  Dashboard generated: {OUTPUT}')
print(f'    Period : {PERIOD}')
print(f'    Filled : {n(f_cur)} units  ({f_rec} records)')
print(f'    Packed : {n(p_cur)} units')
print(f'    Dispatched: {n(d_cur)} units')
print(f'    BSR Stock : {n(bsr_stock)} units')

# ══════════════════════════════════════════════════════════════════════════════
# AUTO-UPLOAD TO GITHUB PAGES (the live public dashboard)
# ══════════════════════════════════════════════════════════════════════════════
# Pushes the freshly generated HTML straight to the GitHub repo via the
# Contents API — no manual drag-and-drop needed. index.html is what the
# live URL serves; Enicar_Dashboard.html is kept in sync for reference.
def _push_to_github(html_text):
    import json as _json
    import base64 as _b64
    import urllib.request as _url
    import urllib.error as _urlerr

    sys.path.insert(0, HERE)
    import email_config as _cfg

    _token = getattr(_cfg, 'GITHUB_TOKEN', '')
    _user  = getattr(_cfg, 'GITHUB_USERNAME', '')
    _repo  = getattr(_cfg, 'GITHUB_REPO', '')

    if not (_token and _user and _repo):
        print('⚠️  GitHub upload skipped: token/username/repo not set in email_config.py.')
        return False

    _content_b64 = _b64.b64encode(html_text.encode('utf-8')).decode('ascii')
    _stamp = datetime.now().strftime('%d %b %Y %H:%M')
    _all_ok = True

    for _path in ('index.html', 'Enicar_Dashboard.html'):
        _api = f'https://api.github.com/repos/{_user}/{_repo}/contents/{_path}'
        _headers = {
            'Authorization': f'token {_token}',
            'Accept': 'application/vnd.github+json',
            'User-Agent': 'enicar-dashboard-bot',
        }

        # Get current file SHA (required to update an existing file)
        _sha = None
        try:
            _req = _url.Request(_api, headers=_headers, method='GET')
            with _url.urlopen(_req, timeout=30) as _r:
                _sha = _json.load(_r).get('sha')
        except _urlerr.HTTPError as _e:
            if _e.code != 404:   # 404 = file doesn't exist yet, that's fine
                print(f'⚠️  GitHub read failed for {_path}: {_e}')
                _all_ok = False
                continue

        _body = {
            'message': f'Auto-update dashboard — {_stamp}',
            'content': _content_b64,
        }
        if _sha:
            _body['sha'] = _sha

        try:
            _req = _url.Request(
                _api, headers=_headers, method='PUT',
                data=_json.dumps(_body).encode('utf-8')
            )
            with _url.urlopen(_req, timeout=30):
                pass
            print(f'🌐  GitHub updated — {_path}')
        except _urlerr.HTTPError as _e:
            print(f'⚠️  GitHub upload failed for {_path}: {_e} — {_e.read().decode("utf-8", "ignore")[:200]}')
            _all_ok = False

    print(f'    Live: https://{_user.lower()}.github.io/{_repo}/')
    return _all_ok

if os.environ.get('GITHUB_ACTIONS'):
    # Running in the cloud — the workflow itself commits the files,
    # so skip the direct API push (no personal token needed there).
    print('   (GitHub Actions detected — workflow will commit; skipping API push)')
else:
    try:
        _push_ok = _push_to_github(html)
    except Exception as _e:
        print(f'⚠️  GitHub upload skipped: {_e}')
        _push_ok = False
    # Signal failure to the caller (check_email_and_refresh.py) so it does
    # NOT save the data hash — that way a failed push is retried next run
    # instead of being silently marked as already published.
    if not _push_ok:
        sys.exit(1)
