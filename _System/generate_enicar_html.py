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


# ── Tolerant sheet-name lookup ────────────────────────────────────────────────
# Staff sometimes rename tabs in the live sheet ("\u2795 Packing Log" became
# "Packing Log" on 8 Aug 2026 and broke every build). Resolve sheet_name by
# canon match (letters/digits only, case-insensitive) so renames never break us.
def _tolerant_read_excel_install():
    import pandas as _pd, re as _re
    if getattr(_pd.read_excel, '_tolerant', False):
        return
    _orig = _pd.read_excel
    def _read_excel(io, sheet_name=0, **kw):
        if isinstance(sheet_name, str):
            try:
                names = _pd.ExcelFile(io).sheet_names
                if sheet_name not in names:
                    c = lambda s: _re.sub(r'[^a-z0-9]', '', str(s).lower())
                    m = [n for n in names if c(n) == c(sheet_name)]
                    if m:
                        sheet_name = m[0]
            except Exception:
                pass
        return _orig(io, sheet_name=sheet_name, **kw)
    _read_excel._tolerant = True
    _pd.read_excel = _read_excel
_tolerant_read_excel_install()

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
                                      'Ronak Exim pvt Ltd.', 'Ronak Exim pvt Ltd.C',
                                      'Ranak Exim pvt Ltd.', 'RONAK'],
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
    'Indoco Remedies':               ['Indoco Remedies', 'Indoco Remedies Limited',
                                      'Indoco  Remedies Limited', 'INDOCO', 'Indoco Remedies Ltd.'],
    'Kamal':                         ['KAMAL'],
    'Enicar Pharmaceutical Pvt Ltd': ['Enicar Pharmaceutical Pvt Ltd',
                                      'Enicar Pharmaceuticals pvt ltd', 'ENICAR'],
    'London United Exports Ltd':     ['London United', 'London United Exports Ltd', 'LONDON', 'LONEDON'],
    'London United Medimpex Pvt. Ltd': ['London United Medimpex', 'London United Medimpex Pvt. Ltd',
                                      'London United Medimpex Pvt Ltd'],
    'Yellow & Orange Pharmacy':      ['YELLOW ORANGE', 'Yellow & Orange Pharmacy',
                                      'Yellow & Orange Pharmacy.'],
    'Tushu Pharma':                  ['TUSHU', 'Tushu Pharma', 'TUSHU PHARMA SARL'],
    'Unichem Ghana':                 ['UNICHEM', 'Unichem Ghana', 'UNICHEM (GHANA) LIMITED'],
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
    'Nucrest':                        ['Nucrest', 'NUCREST'],
    'Biomedicare India Pvt. Ltd':     ['Biomedicare', 'Bio medicare', 'BIOMEDICARE',
                                      'BIOMEDICARE INDIA PVT. LTD'],
    'Shrey Nutraceuticals & Herbals Pvt Ltd': [
                                      'Shrey nutraceuticals & herbals pvt ltd',
                                      'Shrey Nutraceuticals & Herbals Pvt Ltd',
                                      'Shrey Nutraceuticals', 'SHREY'],
    'RPG Life Sciences Ltd':         ['RPG Life sciences', 'RPG Life Sciences',
                                      'RPG LIFE SCIENCES LTD', 'RPG Life Sciences Ltd', 'RPG'],
    'MVC':                            ['MVC'],
    'PRAHEM':                         ['PRAHEM', 'Prahem'],
    'Systopic':                       ['Systopic', 'SYSTOPIC'],
    'Leeford Healthcare Limited':     ['Leeford Healthcare Limited', 'Leeford Healthcare',
                                      'LEEFORD', 'Leeford'],
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
            if d is not None and d.year < 2024:
                d = None                     # impossible year — keep qty, drop date
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
_bj_complete = sum(1 for e in BATCH_JOURNEY if e['rank'] == 3)

# ── Per-stage dates for the Batch Journey timeline ────────────────────────────
# First/last activity date per batch per stage, plus the RM Dispensing record,
# so the lookup can show a start-to-end timeline with elapsed days.
def _stage_dates(df):
    out = {}
    for _, r in df.iterrows():
        b = r.get('Batch')
        if pd.isna(b) or not str(b).strip():
            continue
        k, d = _bkey(b), r.get('Date')
        if d is None or d.year < 2024:      # typo years (e.g. 0226) — ignore for dates
            continue
        e = out.setdefault(k, {'first': d, 'last': d})
        if d < e['first']: e['first'] = d
        if d > e['last']:  e['last']  = d
    return out

FILL_DATES = _stage_dates(fill_df)
PACK_DATES = _stage_dates(pack_df)
DISP_DATES = _stage_dates(disp_df)

def _rm_info():
    """Batch key → RM dispensing record (earliest date, customer, product, batch size)."""
    info = {}
    try:
        rm_all = pd.read_excel(TEMPLATE, sheet_name='➕ RM Dispensing Log', header=3)
        rm_all.columns = [' '.join(str(c).split()) for c in rm_all.columns]
    except Exception:
        return info
    for _, r in rm_all.iterrows():
        b = r.get('BATCH NUMBER')
        if pd.isna(b) or not str(b).strip() or str(b).strip() == '-':
            continue
        k = _bkey(b)
        d = pd.to_datetime(r.get('DISPENSING DATE'), errors='coerce')
        e = info.setdefault(k, {'date': None, 'customer': None, 'product': None, 'size': 0.0,
                                'plan_type': None, 'pack': None})
        if pd.notna(d) and (e['date'] is None or d.date() < e['date']):
            e['date'] = d.date()
        if e['plan_type'] is None and pd.notna(r.get('PLAN')):
            e['plan_type'] = str(r.get('PLAN')).strip().lower()
        if e['customer'] is None and pd.notna(r.get('CUSTOMER')):
            e['customer'] = str(r.get('CUSTOMER')).strip()
        if e['product'] is None and pd.notna(r.get('NAME OF THE PRODUCT')):
            e['product'] = str(r.get('NAME OF THE PRODUCT')).strip()
        if e['pack'] is None and pd.notna(r.get('PACK SIZE')):
            e['pack'] = str(r.get('PACK SIZE')).strip()
        e['size'] += float(pd.to_numeric(r.get('BATCH SIZE'), errors='coerce') or 0)
    return info

RM_INFO = _rm_info()

# ── Stuck batches ─────────────────────────────────────────────────────────────
# A batch is "stuck" when it has sat at one stage with no movement to the next:
#   filled but nothing packed for STUCK_FILL_DAYS+
#   packed but nothing dispatched for STUCK_PACK_DAYS+ (dispatch waits are
#   normal business, so this threshold is deliberately longer)
STUCK_FILL_DAYS = 7
STUCK_PACK_DAYS = 21
_today = date.today()

def _stuck_info(e):
    """Return (stage_label, days_waiting) if the batch is stuck, else None."""
    if e['rank'] in (0, 4):          # problems & opening stock handled elsewhere
        return None
    k = _bkey(e['batch'])
    if e['filled'] > 0 and e['packed'] == 0 and e['dispatched'] == 0:
        fd = FILL_DATES.get(k)
        if fd and (_today - fd['last']).days >= STUCK_FILL_DAYS:
            return ('Filled, waiting for packing', (_today - fd['last']).days)
    elif e['packed'] > 0 and e['dispatched'] == 0 and not e.get('auto_cleared'):
        pdte = PACK_DATES.get(k)
        if pdte and (_today - pdte['last']).days >= STUCK_PACK_DAYS:
            return ('Packed, waiting for dispatch', (_today - pdte['last']).days)
    return None

STUCK_BATCHES = []
for _e in BATCH_JOURNEY:
    _s = _stuck_info(_e)
    if _s:
        STUCK_BATCHES.append({**_e, 'stuck_stage': _s[0], 'stuck_days': _s[1]})
STUCK_BATCHES.sort(key=lambda e: -e['stuck_days'])

# ══════════════════════════════════════════════════════════════════════════════
# MONTHLY PRODUCTION PLAN  (plan vs actual)
# ──────────────────────────────────────────────────────────────────────────────
# Source of truth: a sheet tab whose name contains "PLAN" (so the Plant Head
# edits the plan directly in the Google Sheet — columns: PRODUCT NAME, PARTY,
# PLANNED QTY (UNITS), PACK SIZE, PRIORITY, REMARKS). Until that tab exists,
# falls back to the bundled plan JSON extracted from the Director's docx.
# Matching against actuals is deliberately DATE-TOLERANT: an item's status
# comes from what actually happened for that product+party from the plan
# window onward, no matter which exact day it happened.
# ══════════════════════════════════════════════════════════════════════════════
PLAN_MONTH       = '2026-08'
PLAN_TITLE       = 'AUGUST 2026 PLAN'
PLAN_WINDOW_FROM = date(2026, 8, 1)    # fill/pack/disp counted from here
PLAN_RM_FROM     = date(2026, 7, 25)   # RM dispensing may start a few days early
_PLAN_JSON       = os.path.join(HERE, 'plan_aug_2026.json')

def _pcanon(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())

# ── Party aliases (Director, 22 Aug 2026) ─────────────────────────────────────
# Legally-different names that are the SAME customer for planning purposes.
# The ⇄ "company differs" chip must NOT fire for these. Match: a party belongs
# to a group when its canon name contains any of the group's marker strings.
_PARTY_ALIAS_GROUPS = [
    ('shalina',),                                  # Shalina Laboratories ⇄ Shalina Healthcare
    ('pharmatrust',),                              # Pharmatrust Ltd ⇄ Pharmatrust Limited
    ('londonunited',),                             # London United Exports ⇄ London United Medimpex
    ('aurapharma', 'stanmar'),                     # Aura Pharmaceuticals ⇄ STANMARC ENTERPRISE (P2P)
]

def _party_group(name):
    c = _pcanon(name)
    if not c:
        return None
    for gi, markers in enumerate(_PARTY_ALIAS_GROUPS):
        if any(m in c for m in markers):
            return gi
    return None

def _same_party(a, b):
    A, B = _pcanon(a), _pcanon(b)
    if not A or not B:
        return False
    if A in B or B in A:
        return True
    ga, gb = _party_group(a), _party_group(b)
    return ga is not None and ga == gb

# Plan-sheet short party names → canonical dashboard customers
_PLAN_PARTY_SHORT = {
    'luex': 'London United Exports Ltd', 'p&g': 'Procter & Gamble',
    'aura/stanmark': 'Aura Pharmaceuticals Pvt Ltd', 'workcell': 'Workcell Solution',
    'careth corporation': 'Careth Corporation', 'zieva/enicar': 'Enicar Pharmaceutical Pvt Ltd',
    'laso': 'Laso', 'bliss': 'Bliss GVS', 'galaxy': 'Galaxy Pharma',
    'ronak': 'Ronak Exim pvt Ltd', 'tushu': 'Tushu Pharma', 'unique': 'Unique Pharma',
    'sapphire': 'Sapphire Lifescience Pvt. Ltd.', 'indoco': 'Indoco Remedies',
    'shalina': 'Shalina Laboratories Pvt Ltd', 'rpg': 'RPG Life Sciences Ltd',
    'group': 'Group Pharma', 'pharmatrust': 'Pharmatrust Ltd', 'socomed': 'Socomed Pharma',
    'lesanto': 'Lesanto Laboratories', 'nelpa': 'Nelpa Lifescience',
    'blue map': 'Blue Map Pharmachem', 'kamal': 'Kamal', 'medico': 'Medico',
}
def _plan_party_canon(p):
    raw = ' '.join(str(p or '').strip().split())
    short = _PLAN_PARTY_SHORT.get(raw.lower())
    if short:
        return normalise_party(short)
    n = normalise_party(raw)
    if n != raw:
        return n
    # last resort: prefix-contains against canonical names
    pk = _pkey(raw)
    for canon_name in _PARTY_GROUPS:
        if pk and (pk in _pkey(canon_name) or _pkey(canon_name).startswith(pk)):
            return canon_name
    return raw

def _load_plan():
    """Returns (items, source_label). Sheet tab wins; bundled JSON is fallback."""
    try:
        import openpyxl as _ox
        _sheets = pd.ExcelFile(TEMPLATE).sheet_names
        _tab = next((s for s in _sheets if 'PLAN' in s.upper() and 'DISPENS' not in s.upper()), None)
    except Exception:
        _tab = None
    if _tab:
        try:
            pdf = pd.read_excel(TEMPLATE, sheet_name=_tab, header=0)
            pdf.columns = [' '.join(str(c).split()).upper() for c in pdf.columns]
            def _c(*names):
                return next((c for c in pdf.columns if any(n in c for n in names)), None)
            cp, cparty = _c('PRODUCT'), _c('PARTY', 'CUSTOMER')
            cq, cpack, cprio = _c('QTY', 'PLANNED'), _c('PACK'), _c('PRIORITY')
            # Store-writable columns (optional — added by RM / Plant Head)
            cstat = _c('RM STATUS', 'STATUS')
            cdate = _c('DISPENSE ON', 'DISPENSE BY', 'TARGET', 'RM DATE', 'DISPENSE DATE')
            cbatch, cby = _c('BATCH'), _c('UPDATED BY', 'BY')
            crem = _c('REMARK', 'NOTE')
            def _num(v):
                x = pd.to_numeric(v, errors='coerce')
                return 0.0 if pd.isna(x) else float(x)
            def _txt(v):
                return '' if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v).strip()
            items = []
            for _, r in pdf.iterrows():
                prod = r.get(cp)
                if pd.isna(prod) or not str(prod).strip():
                    continue
                _d = pd.to_datetime(r.get(cdate), dayfirst=True, errors='coerce') if cdate else None
                items.append({'product': str(prod).strip(),
                              'party': _txt(r.get(cparty)),
                              'planned_units': _num(r.get(cq)),
                              'pack': _txt(r.get(cpack)),
                              'priority': int(_num(r.get(cprio))) or None,
                              'rm_status': _txt(r.get(cstat)) if cstat else '',
                              'rm_date': (_d.date() if _d is not None and pd.notna(_d) else None),
                              'rm_batch': _txt(r.get(cbatch)) if cbatch else '',
                              'updated_by': _txt(r.get(cby)) if cby else '',
                              'remark': _txt(r.get(crem)) if crem else ''})
            if items:
                return items, f'sheet tab “{_tab}” (live — edit there)'
        except Exception as _e:
            print(f'  Plan tab unreadable ({_e}) — using bundled plan.')
    try:
        _pj = json.load(open(_PLAN_JSON))
        _fb = [{**i, 'rm_status': '', 'rm_date': None, 'rm_batch': '',
                'updated_by': '', 'remark': ''} for i in _pj.get('items', [])]
        return _fb, 'bundled plan (docx 31 Jul) — add an “AUG PLAN” tab to the sheet to edit live'
    except Exception:
        return [], 'no plan found'

_journey_by_key = {_bkey(e['batch']): e for e in BATCH_JOURNEY}

def _packnum(v):
    """Leading number of a pack size — "15 ml", "15.0", "200 ml (Angola)" → 15.0/200.0."""
    m = re.search(r'\d+(?:\.\d+)?', str(v or ''))
    return float(m.group()) if m else None


def _batch_pack_actuals():
    """(batch key) → {pack number or None: filled/packed/dispatched}.

    Pack size is a property of the FILLING / PACKING / DISPATCH row, not of the
    batch. One bulk batch routinely fills two bottle sizes — EL-2509 Ferrolife
    (1,515 L) made 5,000 x 100 ml AND 5,000 x 200 ml — and RM's own PACK SIZE
    column records only one of them. So a plan line must be credited from the
    logs, per pack, or the second pack size looks "not started" while its stock
    is already dispatched (Director, 26 Aug 2026).
    """
    out = {}
    def add(df, field, qcol):
        if df is None or len(df) == 0 or qcol not in df.columns:
            return
        for b, pk, q in zip(df['Batch'], df['PackSize'], df[qcol]):
            if b is None or (not isinstance(b, str) and pd.isna(b)):
                continue
            k = _bkey(b)
            if not k or k == '-':
                continue
            v = pd.to_numeric(q, errors='coerce')
            e = out.setdefault(k, {}).setdefault(
                _packnum(pk), {'filled': 0.0, 'packed': 0.0, 'dispatched': 0.0})
            e[field] += 0.0 if pd.isna(v) else float(v)
    add(fill_df, 'filled',     'Qty')
    add(pack_df, 'packed',     'TotalPacked')
    add(disp_df, 'dispatched', 'Qty')
    return out

BATCH_PACK = _batch_pack_actuals()


def _rm_batches_in_window():
    """RM batches that belong to this month's plan — the authoritative link
    between the plan and real production. Batch number and COMPANY NAME both
    come from RM: under loan-licence / P2P arrangements the plan may carry the
    marketing company while RM carries the manufacturing licence holder.

    A batch qualifies if its RM dispense falls in the window, OR if the real
    work (filling / packing / dispatch) happened during the plan month. RM can
    legitimately be dispensed well before production starts — NP-1060 Becatone-L
    was dispensed 20 Jul but filled, packed and dispatched in August, and used
    to show as "Not started" because only the RM date was considered
    (Director, 20 Aug 2026)."""
    def _worked_this_month(k):
        for src in (FILL_DATES, PACK_DATES, DISP_DATES):
            d = src.get(k)
            if d and d.get('last') and d['last'] >= PLAN_WINDOW_FROM:
                return True
        return False
    out = []
    for k, v in RM_INFO.items():
        if not v.get('date'):
            continue
        if v['date'] < PLAN_RM_FROM and not _worked_this_month(k):
            continue
        out.append({'key': k, 'batch': (_journey_by_key.get(k, {}) or {}).get('batch', k),
                    'product': v.get('product') or '', 'customer': v.get('customer') or '',
                    'date': v['date'], 'size': v.get('size') or 0,
                    'pack': v.get('pack') or ''})
    # ── Production with NO RM record still belongs on the plan card ──
    # MinMin PS (REM387, 50,000 filled + 49,150 dispatched) and P&G Polybion
    # 150 ml (6230C85001/2, 6233C85001 — 98,500 filled) showed "Not started"
    # because plan linking only ever looked at the RM log. A batch that was
    # filled or packed inside the plan month is this month's work whether or
    # not the store remembered its RM row (Director, 26 Aug 2026). Product,
    # party and dates come from the production logs; the plan line gets a
    # flag so the missing RM entry still gets chased, not hidden.
    _seen = {e['key'] for e in out}
    _cand = {}
    for _df in (fill_df, pack_df):
        for _b, _p, _pa, _d in zip(_df['Batch'], _df['Product'], _df['Party'], _df['Date']):
            if _b is None or (not isinstance(_b, str) and pd.isna(_b)):
                continue
            _k = _bkey(_b)
            if not _k or _k in _seen or _k in ('-', 'NAN', 'NA'):
                continue
            if _k in RM_INFO and RM_INFO[_k].get('date'):
                continue          # has an RM record — the window rules above own it
            e = _cand.setdefault(_k, {'batch': str(_b).strip(), 'product': None,
                                      'customer': None, 'date': None})
            if _p is not None and e['product'] is None:
                e['product'] = str(_p)
            if _pa is not None and e['customer'] is None:
                e['customer'] = str(_pa)
            if e['date'] is None or _d < e['date']:
                e['date'] = _d
    for _k, e in _cand.items():
        _fd = FILL_DATES.get(_k)
        _first = (_fd and _fd['first']) or e['date']
        if _first is None or _first < PLAN_WINDOW_FROM:
            continue              # previous-month work / opening stock
        out.append({'key': _k, 'batch': e['batch'], 'product': e['product'] or '',
                    'customer': e['customer'] or '', 'date': e['date'],
                    'size': 0, 'pack': '', 'rm_missing': True})
    return out

def _build_plan_view():
    raw_items, source = _load_plan()
    if not raw_items:
        return [], source, {}, []
    # ── Consolidate June–July carry-over into the same plan (Director, 22 Aug):
    # every row carries a month tag. Pending items whose product is ALREADY on
    # the AUG plan just badge that row ("also pending from JUN/JUL"); the rest
    # join the table as their own rows so the one card is the whole workload.
    for it in raw_items:
        # A sheet row whose REMARKS says "carry-over from JUL/JUN" keeps its
        # original month chip — so the Director can paste the pending list
        # straight into the AUG PLAN tab and the segregation survives.
        _rl = (it.get('remark') or '').lower()
        if 'carr' in _rl and 'jun' in _rl:
            it['month'] = 'JUN'
        elif 'carr' in _rl and 'jul' in _rl:
            it['month'] = 'JUL'
        else:
            it['month'] = 'AUG'
    _carry_badge = {}
    try:
        _pending = json.load(open(os.path.join(HERE, 'pending_plan_jun_jul.json'))).get('items', [])
    except Exception:
        _pending = []
    def _pp_match(a, b):
        A, B = _pcanon(a), _pcanon(b)
        return A and B and (A == B or (len(A) >= 5 and len(B) >= 5 and (A in B or B in A)))
    for _pe in _pending:
        _mon = 'JUN' if str(_pe.get('plan_date') or '')[5:7] == '06' else 'JUL'
        # Match against EVERY sheet row, not just AUG ones. Once the Director
        # pastes the carry-over list into the plan tab those rows carry a JUN/JUL
        # chip, and an AUG-only test re-injected them as phantom duplicates —
        # and where PARTY was blank the two copies merged and DOUBLED the
        # planned quantity (METRON-F showed 17,100 for 8,550). 26 Aug 2026.
        _hit = next((it for it in raw_items
                     if _pp_match(it['product'], _pe['product'])), None)
        if _hit is not None:
            k = _pcanon(_hit['product'])
            e = _carry_badge.setdefault(k, {'months': set(), 'units': 0})
            e['months'].add(_mon)
            e['units'] += int(_pe.get('pending') or 0)
        else:
            raw_items.append({'product': _pe['product'], 'party': '',
                              'planned_units': float(_pe.get('pending') or 0),
                              'pack': _pe.get('pack') or '', 'priority': None,
                              'rm_status': '', 'rm_date': None, 'rm_batch': '',
                              'updated_by': '', 'remark': f"carry-over from {_mon}",
                              'month': _mon})
    # Merge duplicate product+party+pack lines (lots) so actuals aren't double-counted
    merged = {}
    for it in raw_items:
        key = (_pcanon(it['product']), _pcanon(_plan_party_canon(it['party'])), _pcanon(it.get('pack')))
        m = merged.setdefault(key, {**it, 'party_canon': _plan_party_canon(it['party']),
                                    'planned_units': 0, 'lots': 0, 'prio_list': []})
        m['planned_units'] += float(it.get('planned_units') or 0)
        m['lots'] += 1
        if it.get('priority'):
            m['prio_list'].append(it['priority'])
        # Store-written fields: keep the most informative value across merged lots
        for fld in ('rm_status', 'rm_batch', 'updated_by', 'remark'):
            if it.get(fld) and not m.get(fld):
                m[fld] = it[fld]
        if it.get('rm_date') and not m.get('rm_date'):
            m['rm_date'] = it['rm_date']
    items = list(merged.values())
    for it in items:
        _cb = _carry_badge.get(_pcanon(it['product']))
        if it.get('month') == 'AUG' and _cb:
            it['carry_months'] = '+'.join(sorted(_cb['months']))
            it['carry_units'] = _cb['units']
    for it in items:
        it['priority'] = min(it['prio_list']) if it['prio_list'] else None

    rmw = _rm_batches_in_window()
    used_keys = set()
    _blank_claimed = set()   # batches whose no-pack log rows are already credited
    _brand_line_count = {}
    _exact_plan_names = {_pcanon(it['product']) for it in items if _pcanon(it['product'])}

    # Product aliases the Director confirmed are the SAME item (22 Aug 2026).
    # "MinMin PS" is the physician-sample pack of Minmin Tonic.
    # Brands packed as a MIXED carton — one sachet of each flavour per carton.
    _MIXED_CARTON = {'kifarujelly'}
    _PRODUCT_ALIASES = [frozenset(('minminps', 'minmintonic')),
                        frozenset(('unidaktgeloral', 'unidaktcream'))]

    def _brandkey(s):
        """Product name with a parenthetical flavour and a trailing strength
        token removed: "Kifaru Jelly 100mg" and "KIFARU JELLY (Strawberry
        Flavour)" both become "kifarujelly". Nothing else is stripped, so
        Polybion Lc / Polybion Active and Allerzy-DC / -X stay distinct."""
        t = re.sub(r'\([^)]*\)', ' ', str(s or ''))
        t = re.sub(r'\s+\d+(?:\.\d+)?\s*(mg|ml|gm|g|mcg|%)?\s*$', ' ', t, flags=re.I)
        return _pcanon(t)

    for _it in items:
        _k = _brandkey(_it['product'])
        _brand_line_count[_k] = _brand_line_count.get(_k, 0) + 1

    def prod_match(a, b):
        A, B = _pcanon(a), _pcanon(b)
        KA, KB = _brandkey(a), _brandkey(b)
        # Brand-level match only when the plan does NOT split that brand by
        # flavour itself. Bedisyl 100 Jelly is planned as three separate lines
        # (banana / mango / pineapple), so each must keep its own batch;
        # Kifaru Jelly is one line covering both flavours, so it takes all.
        if (KA and KA == KB and len(KA) >= 5
                and _brand_line_count.get(KA, 0) == 1):
            return True
        if A and A == B:           # exact name — covers short ones like PA-C
            return True
        if frozenset((A, B)) in _PRODUCT_ALIASES:
            return True
        return len(A) >= 4 and len(B) >= 4 and (A in B or B in A)

    # Every pack size each product is planned in — used to spot production
    # logged in a pack NO line of that product plans (usually a typo in the
    # log's Pack Size cell: Amylase plan 200 ml vs logs 150 ml).
    _packs_by_canon = {}
    for _it in items:
        _packs_by_canon.setdefault(_pcanon(_it['product']), set()).add(_packnum(_it.get('pack')))
    _orphan_claimed = set()

    for it in items:
        _me = _pcanon(it['product'])
        # ── Link plan → RM ──
        # If the store wrote a batch number in the plan row, that is definitive
        # (it also rescues items whose product wording differs from the plan).
        # Otherwise match on PRODUCT ONLY — party may legitimately differ.
        wanted = {_bkey(x) for x in re.split(r'[,/;]+', it.get('rm_batch') or '') if x.strip()}
        if wanted:
            hits = [b for b in rmw if b['key'] in wanted]
            it['linked_by'] = 'batch written by store'
            for w in wanted:                       # batch typed but not in RM log yet
                if not any(b['key'] == w for b in rmw):
                    it.setdefault('missing_batches', []).append(w)
        else:
            hits = [b for b in rmw if prod_match(it['product'], b['product'])]
            # "Gelucid" is a substring of "Gelucid O", so the loose match handed
            # each line the other's batches and double-counted both. A batch
            # whose RM name is EXACTLY another plan line's product belongs to
            # that line alone (26 Aug 2026). Brand-level matches such as
            # "Kifaru Jelly 100mg" -> "KIFARU JELLY (Strawberry Flavour)" are
            # untouched: the RM name is not itself a plan line.
            hits = [b for b in hits
                    if _pcanon(b['product']) == _me
                    or _pcanon(b['product']) not in _exact_plan_names]
            it['linked_by'] = 'product name + pack size'
        # ── Credit each batch to this line PER PACK SIZE ──
        # Same product planned in two packs must not share the same units
        # (Alaize 15 ml vs 200 ml double-counted), but one bulk batch may
        # genuinely serve both packs (Ferrolife EL-2509). So the quantities come
        # from the log rows for THIS pack; a batch that produced only other
        # packs is not this line's batch at all. Rows where the log left the
        # pack blank go to the first line that claims the batch.
        _want = _packnum(it.get('pack'))
        batches = []
        for b in hits:
            sl = BATCH_PACK.get(b['key'], {})
            j = _journey_by_key.get(b['key'], {})
            if _want is None or not sl:
                # No pack on the plan row, or nothing produced yet: fall back to
                # whole-batch figures, but don't attach an RM batch whose own
                # pack size contradicts the line.
                _rp = _packnum(b.get('pack'))
                if _want is not None and not sl and _rp is not None and _rp != _want:
                    continue
                qty = {'filled': float(j.get('filled') or 0),
                       'packed': float(j.get('packed') or 0),
                       'dispatched': float(j.get('dispatched') or 0)}
            else:
                mine = sl.get(_want)
                blank = sl.get(None) if b['key'] not in _blank_claimed else None
                if blank is not None:
                    _blank_claimed.add(b['key'])
                # Production logged in a pack size that NO line of this product
                # plans — a Pack Size typo in one of the logs (Bonaplex EL-2438
                # packed rows say 200ML, the only plan line is 250 ml; Amylase
                # logs say 150 ml, the plan says 200 ml). Credit it HERE, once,
                # with a flag, instead of letting the units vanish — but only
                # when the log's own product name is exactly this line's.
                extra = None
                if _pcanon(b['product']) == _me:
                    _own = _packs_by_canon.get(_me, set())
                    for _pk, _q in sl.items():
                        if _pk is None or _pk in _own:
                            continue
                        if (b['key'], _pk) in _orphan_claimed:
                            continue
                        _orphan_claimed.add((b['key'], _pk))
                        extra = {f: (extra or {}).get(f, 0.0) + _q[f]
                                 for f in ('filled', 'packed', 'dispatched')}
                        it.setdefault('pack_notes', []).append(
                            f"{b['batch']} logged as {_pk:g} (plan: {it.get('pack')})")
                if mine is None and blank is None and extra is None:
                    if not wanted:          # produced in other packs only
                        continue
                    mine = {'filled': 0.0, 'packed': 0.0, 'dispatched': 0.0}
                qty = {f: (mine or {}).get(f, 0.0) + (blank or {}).get(f, 0.0)
                       + (extra or {}).get(f, 0.0)
                       for f in ('filled', 'packed', 'dispatched')}
            used_keys.add(b['key'])
            _fd, _pd2, _dd2 = FILL_DATES.get(b['key']), PACK_DATES.get(b['key']), DISP_DATES.get(b['key'])
            batches.append({
                'batch': b['batch'], 'rm_date': b['date'], 'rm_customer': b['customer'],
                'rm_product': b['product'], 'size': b['size'], 'rm_missing': b.get('rm_missing', False), **qty,
                'status': j.get('status') or 'RM dispensed — no production yet',
                'f_first': _fd and _fd['first'], 'f_last': _fd and _fd['last'],
                'p_first': _pd2 and _pd2['first'], 'p_last': _pd2 and _pd2['last'],
                'd_first': _dd2 and _dd2['first'], 'd_last': _dd2 and _dd2['last'],
            })
        batches.sort(key=lambda x: x['rm_date'])
        it['batches'] = batches
        # Mixed-carton products: one carton holds one sachet of EACH flavour, so
        # the plan quantity counts CARTONS. Kifaru Jelly is planned 2,00,000 and
        # made as 2,00,000 strawberry + 2,00,000 banana on separate batches —
        # adding the flavours together would read 4,00,000 against a 2,00,000
        # plan (Director, 26 Aug 2026). Credit the per-flavour average instead.
        _grp = {}
        for _b in batches:
            _grp.setdefault(_pcanon(_b.get('rm_product')), []).append(_b)
        _div = len(_grp) if (_brandkey(it['product']) in _MIXED_CARTON and len(_grp) > 1) else 1
        it['mixed_carton'] = _div > 1
        it['filled']     = sum(b['filled'] for b in batches) / _div
        it['packed']     = sum(b['packed'] for b in batches) / _div
        it['dispatched'] = sum(b['dispatched'] for b in batches) / _div
        # Company name is taken from RM; flag when the plan shows a different one
        rm_customers = sorted({b['rm_customer'] for b in batches if b['rm_customer']})
        it['rm_customers'] = rm_customers
        # RM dispensing name is the authoritative company (Director, 22 Aug 2026):
        # display it whenever batches exist; the plan's own party becomes a small
        # "plan: X" note only when it genuinely names a different company.
        it['display_party'] = (' / '.join(rm_customers) if rm_customers
                               else it['party_canon'])
        it['party_differs'] = (bool(rm_customers) and bool(_pcanon(it['party_canon']))
                               and not any(_same_party(it['party_canon'], c)
                                           for c in rm_customers))

        # Reconcile what the store wrote against what the logs actually show
        _st = (it.get('rm_status') or '').strip().lower()
        it['flag'] = ''
        if _st.startswith('dispensed') and not batches:
            it['flag'] = 'Marked dispensed, but no RM log entry found — please add the RM row'
        elif batches and _st and not _st.startswith(('dispens', 'done', 'complete')):
            it['flag'] = f'RM has already dispensed {len(batches)} batch(es) — status can be updated'
        if it.get('missing_batches'):
            it['flag'] = ('Batch ' + ', '.join(it['missing_batches'])
                          + ' written here is not in the RM log yet')
        _no_rm = [b['batch'] for b in batches if b.get('rm_missing')]
        _notes = []
        if _no_rm:
            _notes.append('no RM Dispensing entry for ' + ', '.join(_no_rm))
        if it.get('pack_notes'):
            _notes.append('pack size mismatch — ' + '; '.join(it['pack_notes'][:3]))
        if _notes:
            it['flag'] = (it['flag'] + ' · ' if it['flag'] else '') + ' · '.join(_notes)

        plan_q = it['planned_units'] or 0
        if not batches:
            it['status'], it['srank'] = ('🟣 ' + it['rm_status'], 0) if it.get('rm_status') else ('⚪ Not started', 0)
        elif plan_q and it['dispatched'] >= plan_q * 0.95:
            it['status'], it['srank'] = '✅ Done (dispatched)', 5
        elif it['dispatched'] > 0:
            it['status'], it['srank'] = '🟠 Dispatching', 4
        elif it['packed'] > 0:
            it['status'], it['srank'] = '🟡 Packed', 3
        elif it['filled'] > 0:
            it['status'], it['srank'] = '🔵 Filling', 2
        else:
            it['status'], it['srank'] = '🟤 RM dispensed', 1
        it['pct'] = min(100.0, (it['filled'] / plan_q * 100) if plan_q else 0)

    # ── Dispensing schedule ──────────────────────────────────────────────
    # Anything the Plant Head / store gave a target date to is an instruction:
    # it sorts to the top of the plan, soonest first, and drives the schedule
    # card. Plant time (IST) decides what "today" means, not the build server.
    _ist_today = (datetime.utcnow() + timedelta(hours=5, minutes=30)).date()
    for it in items:
        td = it.get('rm_date')
        it['due_days'] = (td - _ist_today).days if td else None
        st = (it.get('rm_status') or '').lower()
        it['is_done'] = bool(it['batches']) or st.startswith(('dispensed', 'done', 'complete'))
        if td and not it['is_done']:
            d = it['due_days']
            it['due_bucket'] = ('overdue' if d < 0 else 'today' if d == 0
                                else 'tomorrow' if d == 1 else 'later')
        else:
            it['due_bucket'] = ''
    _BUCKET_ORDER = {'overdue': 0, 'today': 1, 'tomorrow': 2, 'later': 3, '': 4}
    items.sort(key=lambda x: (_BUCKET_ORDER[x['due_bucket']],
                              x['due_days'] if x['due_days'] is not None else 9999,
                              x['priority'] or 9, -(x['planned_units'] or 0)))
    # RM batches dispensed in the window that no plan line claims.
    # July carry-over is NOT off-plan. Two exclusions:
    #   1. dispensed before the plan month started → last month's plan
    #      (the pre-month RM window exists only to credit early starts of
    #       matched plan items, never to raise alarms);
    #   2. filling already began before the plan month → last month's work
    #      even if the RM date reads later (catches mistyped future dates).
    def _is_prev_month_work(b):
        if b['date'] < PLAN_WINDOW_FROM:
            return True
        fd = FILL_DATES.get(b['key'])
        return bool(fd and fd['first'] < PLAN_WINDOW_FROM)
    # Plan rows often carry only the base/brand name — "Kifaru Jelly 100mg"
    # must cover "KIFARU JELLY (Strawberry Flavour)". Same brand anywhere on
    # the plan ⇒ never an off-plan alarm (Director, 6 Aug 2026).
    def _brand_of(p):
        tok = str(p or '').strip().split()[0] if str(p or '').strip() else ''
        return _pcanon(re.sub(r'[-–]?\d+$', '', tok))
    _plan_brands = {b for b in (_brand_of(it['product']) for it in items) if len(b) >= 4}
    # Products whose RM was dispensed in the PREVIOUS month were on that
    # month's plan and are carrying over — deliberately not repeated on this
    # month's plan, so new batches of them are never off-plan (Director, 6 Aug).
    _prev_end = PLAN_WINDOW_FROM - timedelta(days=1)
    _prev_start = _prev_end.replace(day=1)
    _prev_prods, _prev_brands = set(), set()
    for _v in RM_INFO.values():
        if _v.get('date') and _prev_start <= _v['date'] <= _prev_end and _v.get('product'):
            _prev_prods.add(_pcanon(_v['product']))
            _prev_brands.add(_brand_of(_v['product']))
    def _carried_over(p):
        c = _pcanon(p)
        if c and c in _prev_prods:
            return True
        if len(c) >= 4 and any(len(x) >= 4 and (c in x or x in c) for x in _prev_prods):
            return True
        b = _brand_of(p)
        return len(b) >= 4 and b in _prev_brands
    # RM's PLAN column classifies each dispense: only "regular" production can
    # be off-plan — "trial" (TLB etc.) and "additional" are deliberate extras.
    def _is_regular(b):
        return (RM_INFO.get(b['key'], {}).get('plan_type') or '') == 'regular'
    off_plan = sorted([b for b in rmw
                       if b['key'] not in used_keys and not _is_prev_month_work(b)
                       and _is_regular(b)
                       and _brand_of(b['product']) not in _plan_brands
                       and not _carried_over(b['product'])],
                      key=lambda b: b['date'])
    # ── Self-check (26 Aug 2026): no batch may be credited to plan lines for
    # MORE than the logs say it produced. This is the invariant the Alaize /
    # Gelucid / Kifaru bugs all broke — if a future name or pack pattern breaks
    # it again, say so loudly in the build log instead of shipping wrong numbers.
    _credit = {}
    for _it in items:
        for _b in _it['batches']:
            _c = _credit.setdefault(_bkey(_b['batch']), {'filled': 0.0, 'packed': 0.0, 'dispatched': 0.0})
            for _f in ('filled', 'packed', 'dispatched'):
                _c[_f] += _b[_f]
    _div_by = {}
    for _it in items:                     # mixed-carton lines credit the average
        if _it.get('mixed_carton'):
            for _b in _it['batches']:
                _div_by[_bkey(_b['batch'])] = True
    for _k, _c in _credit.items():
        _j = _journey_by_key.get(_k, {})
        for _f in ('filled', 'packed', 'dispatched'):
            _act = float(_j.get(_f) or 0)
            if _c[_f] > _act + 1 and not _div_by.get(_k):
                print(f'⚠️  PLAN SELF-CHECK: batch {_k} credited {_c[_f]:,.0f} {_f} '
                      f'but logs show only {_act:,.0f} — plan figures may double-count!')

    _today_op = date.today()
    for b in off_plan:
        b['future'] = b['date'] > _today_op
    summary = {
        'items': len(items),
        'units': sum(x['planned_units'] or 0 for x in items),
        'started': sum(1 for x in items if x['srank'] > 0),
        'done': sum(1 for x in items if x['srank'] == 5),
        'filled': sum(x['filled'] for x in items),
        'batches': sum(len(x['batches']) for x in items),
        'off_plan': len(off_plan),
        'written': sum(1 for x in items if x.get('rm_status')),
        'flags': sum(1 for x in items if x.get('flag')),
        'next': sum(1 for x in items if 'next' in (x.get('rm_status') or '').lower()
                    or 'tomorrow' in (x.get('rm_status') or '').lower()),
        'overdue': sum(1 for x in items if x['due_bucket'] == 'overdue'),
        'due_today': sum(1 for x in items if x['due_bucket'] == 'today'),
        'due_tomorrow': sum(1 for x in items if x['due_bucket'] == 'tomorrow'),
        'due_later': sum(1 for x in items if x['due_bucket'] == 'later'),
        'packed_sum': sum(x['packed'] for x in items),
        'disp_sum': sum(x['dispatched'] for x in items),
        'rm_started': sum(1 for x in items if x['batches']),
        'carry_items': sum(1 for x in items if x.get('month') != 'AUG'),
        'carry_units': sum((x['planned_units'] or 0) for x in items if x.get('month') != 'AUG'),
    }
    return items, source, summary, off_plan

# ── Pending production plan (June–July 2026 carry-over) ───────────────────────
# A static snapshot the Director supplied (22 Aug 2026): products PLANNED in
# Jun–Jul that still had unproduced quantity when the plan was exported. We
# cross-check each against real production so "still pending" reflects reality:
# a pending line is CLEARED once that product was filled on or after its plan
# date (matched on product name; pack size shown for context).
def _build_pending_plan():
    path = os.path.join(HERE, 'pending_plan_jun_jul.json')
    try:
        data = json.load(open(path))
    except Exception:
        return None
    # product canon → most recent filling / dispatch date across all logs
    last_fill, last_disp = {}, {}
    for _, r in fill_df.iterrows():
        c = _pcanon(r.get('Product'))
        if c and r.get('Date'):
            if c not in last_fill or r['Date'] > last_fill[c]:
                last_fill[c] = r['Date']
    for _, r in disp_df.iterrows():
        c = _pcanon(r.get('Product'))
        if c and r.get('Date'):
            if c not in last_disp or r['Date'] > last_disp[c]:
                last_disp[c] = r['Date']

    def _match(canon_prod):
        # exact, then containment either way (>=5 chars to avoid junk matches)
        if canon_prod in last_fill:
            return last_fill[canon_prod], last_disp.get(canon_prod)
        for c, d in last_fill.items():
            if len(canon_prod) >= 5 and (canon_prod in c or c in canon_prod):
                return d, last_disp.get(c)
        return None, None

    items = []
    for it in data.get('items', []):
        c = _pcanon(it['product'])
        pdate = date.fromisoformat(it['plan_date']) if it.get('plan_date') else None
        fdate, ddate = _match(c)
        # cleared = produced on/after the plan date (fresh production, not the
        # earlier run that this line was already pending against)
        done = bool(fdate and pdate and fdate >= pdate)
        items.append({**it, 'plan_dt': pdate, 'last_fill': fdate,
                      'last_disp': ddate, 'done': done})
    # still-pending first, then by plan date
    items.sort(key=lambda x: (x['done'], x['plan_dt'] or date.max))
    still = [i for i in items if not i['done']]
    summary = {
        'total': len(items), 'done': sum(1 for i in items if i['done']),
        'pending': len(still),
        'units_total': sum(i['pending'] for i in items),
        'units_pending': sum(i['pending'] for i in still),
    }
    return {'items': items, 'summary': summary, 'source': data.get('source', '')}

PENDING_PLAN = _build_pending_plan()

PLAN_ITEMS, PLAN_SOURCE, PLAN_SUMMARY, PLAN_OFF = _build_plan_view()

# ══════════════════════════════════════════════════════════════════════════════
# NAME CONSISTENCY  —  same batch, different product name across logs
# ──────────────────────────────────────────────────────────────────────────────
# Director's rule (31 Jul 2026): whenever filling / packing / dispatch use a
# different product name for the same batch number, it must be verified and
# corrected. RM is the source of truth for the correct name.
# Severity: a different flavour/variant is a real error; a short form or a
# misspelling is a naming-discipline fix.
# ══════════════════════════════════════════════════════════════════════════════
_NAME_SOP = {'SL05406', 'SL05407'}      # one batch legitimately split Red/White

def _collect_names():
    out = {}
    def grab(df, bcol_names, prod_names, log):
        bc = next((c for c in df.columns if str(c).strip().lower() in
                   [x.lower() for x in bcol_names]), None)
        pcx = next((c for c in df.columns if 'product' in str(c).lower()), None)
        if bc is None or pcx is None:
            return
        for _, r in df.iterrows():
            b, p = r.get(bc), r.get(pcx)
            if pd.isna(b) or not str(b).strip() or str(b).strip() == '-':
                continue
            if pd.isna(p) or not str(p).strip():
                continue
            out.setdefault(_bkey(b), {}).setdefault(log, set()).add(str(p).strip())
    try:
        _rm_raw = pd.read_excel(TEMPLATE, sheet_name='➕ RM Dispensing Log', header=3)
        _rm_raw.columns = [' '.join(str(c).split()) for c in _rm_raw.columns]
        _rm_raw = _rm_raw.rename(columns={'BATCH NUMBER': 'Batch No.',
                                          'NAME OF THE PRODUCT': 'Product Name'})
        grab(_rm_raw, ['batch no.'], ['product name'], 'RM')
    except Exception:
        pass
    grab(fill_df.rename(columns={'Batch': 'Batch No.', 'Product': 'Product Name'}),
         ['batch no.'], ['product name'], 'Filling')
    grab(pack_df.rename(columns={'Batch': 'Batch No.', 'Product': 'Product Name'}),
         ['batch no.'], ['product name'], 'Packing')
    grab(disp_df.rename(columns={'Batch': 'Batch No.', 'Product': 'Product Name'}),
         ['batch no.'], ['product name'], 'Dispatch')
    return out

_FLAVOUR_WORDS = ['mango', 'mint', 'orange', 'banana', 'chocolate', 'cherry', 'honey',
                  'grapes', 'pineapple', 'mixfruit', 'vanilla', 'strawberry', 'cocktail',
                  'melon', 'lemon']

def _flavours(name):
    c = _pcanon(name)
    return {w for w in _FLAVOUR_WORDS if w in c}

def _name_conflicts():
    """Director's rule (updated 31 Jul 2026): staff type names manually, so
    routine spelling differences between logs are NOT worth showing. Surface a
    batch only when:
      (a) filling / packing / dispatch AGREE with each other but RM says
          something different  →  the RM entry itself needs verifying, or
      (b) the names differ by FLAVOUR / VARIANT anywhere  →  one entry is on
          the wrong product and must be verified before correcting."""
    data = _collect_names()
    out = []
    for k, logs in data.items():
        if k in _NAME_SOP or len(logs) < 2:
            continue
        def clusters_of(names):
            from difflib import SequenceMatcher
            cs = sorted({_pcanon(x) for x in names}, key=len)
            cl = []
            for x in cs:
                if not any(x in c or c in x
                           or SequenceMatcher(None, x, c).ratio() >= 0.82
                           for c in cl):
                    cl.append(x)
            return cl
        allnames = set()
        for s in logs.values():
            allnames |= s
        # (b) flavour / variant conflict anywhere → always show
        flav = [_flavours(x) for x in allnames if _flavours(x)]
        differing_flavour = len({frozenset(f) for f in flav}) > 1 if flav else False
        if differing_flavour:
            out.append({'batch': (_journey_by_key.get(k, {}) or {}).get('batch', k),
                        'logs': {lg: sorted(v) for lg, v in logs.items()},
                        'rm': (sorted(logs['RM'])[0] if logs.get('RM') else None),
                        'sev': '🔴 Different flavour / variant — verify which is right', 'srank': 0})
            continue
        # (a) production logs agree with each other, RM differs → verify RM
        down = set()
        for lg in ('Filling', 'Packing', 'Dispatch'):
            down |= logs.get(lg, set())
        if not down or not logs.get('RM'):
            continue
        if len(clusters_of(down)) != 1:
            continue                      # staff spelling noise — hidden by rule
        if len(clusters_of(down | logs['RM'])) > 1:
            out.append({'batch': (_journey_by_key.get(k, {}) or {}).get('batch', k),
                        'logs': {lg: sorted(v) for lg, v in logs.items()},
                        'rm': sorted(logs['RM'])[0],
                        'sev': '🟠 All production logs agree — RM name differs, verify the RM entry',
                        'srank': 1})
    out.sort(key=lambda x: (x['srank'], x['batch']))
    return out

NAME_CONFLICTS = _name_conflicts()



def name_conflict_html():
    if not NAME_CONFLICTS:
        return ''
    crit = sum(1 for c in NAME_CONFLICTS if c['srank'] == 0)
    rows = ''
    for i, c in enumerate(NAME_CONFLICTS):
        bg = '#FFF8F1' if i % 2 == 0 else '#FFFFFF'
        per_log = '<br>'.join(
            f'<span style="color:#90A4AE">{lg}:</span> {", ".join(v)}'
            for lg, v in c['logs'].items())
        rows += (f'<tr style="background:{bg}">'
                 f'<td class="td-name" style="font-weight:600;white-space:nowrap">{c["batch"]}</td>'
                 f'<td class="td-name" style="font-size:12px;line-height:1.6">{per_log}</td>'
                 f'<td class="td-name" style="font-weight:600;color:{C_GRN}">{c["rm"] or "— (no RM record)"}</td>'
                 f'<td class="td-name" style="font-size:12px;white-space:nowrap">{c["sev"]}</td></tr>')
    return f'''
<details class="card" id="nameconf-card">
  <summary>{sec(f'  ━━&nbsp;&nbsp;⚠ NAME &nbsp; MISMATCH &nbsp; — &nbsp; SAME &nbsp; BATCH, &nbsp; DIFFERENT &nbsp; PRODUCT &nbsp; NAME &nbsp; ({len(NAME_CONFLICTS)}) &nbsp;━━', C_AMB)}</summary>
  <div style="font-size:12px;color:#607D8B;padding:8px 16px 0">
    Routine spelling variations from manual typing are <strong>hidden</strong>. A batch appears here only when
    <strong>filling, packing and dispatch all agree with each other but RM says something different</strong>
    (verify the RM entry), or when names differ by <strong>flavour / variant</strong> ({crit} such — one entry is
    on the wrong product and must be verified before correcting).
  </div>
  <div class="tbl-wrap" style="padding-top:10px">
    <table style="min-width:720px">
      <thead><tr class="th-row">
        <th>BATCH</th><th>NAME USED IN EACH LOG</th><th>CORRECT NAME (RM)</th><th>ACTION</th>
      </tr></thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
</details>'''

def _plan_batch_detail(it, idx):
    """Hidden row shown when a plan line is clicked — the verification view."""
    if not it['batches']:
        return (f'<tr id="plan-d-{idx}" style="display:none"><td colspan="11" '
                f'style="background:#FAFDFC;padding:10px 16px;color:#90A4AE;font-size:12px">'
                f'No RM dispensing recorded yet for this item. Once the store dispenses it, the '
                f'batch number appears here automatically and filling / packing / dispatch follow it.'
                f'</td></tr>')
    def _dr(a, z):
        if not a:
            return ''
        s2 = a.strftime('%d %b') if a == z or not z else f'{a.strftime("%d %b")}–{z.strftime("%d %b")}'
        return f'<div style="font-size:10px;color:#90A4AE">{s2}</div>'
    rows = ''
    for b in it['batches']:
        rows += (f'<tr>'
                 f'<td class="td-name" style="font-weight:600">{b["batch"]}</td>'
                 f'<td class="td-name">{b["rm_date"].strftime("%d %b")}</td>'
                 f'<td class="td-name" style="color:#546E7A">{b["rm_customer"] or "—"}</td>'
                 f'<td class="td-num">{n(b["size"]) if b["size"] else "—"}</td>'
                 f'<td class="td-num" style="color:{C_SEC}">{n(b["filled"]) if b["filled"] else "—"}{_dr(b.get("f_first"), b.get("f_last"))}</td>'
                 f'<td class="td-num" style="color:{C_AMB}">{n(b["packed"]) if b["packed"] else "—"}{_dr(b.get("p_first"), b.get("p_last"))}</td>'
                 f'<td class="td-num" style="color:{C_ORG}">{n(b["dispatched"]) if b["dispatched"] else "—"}{_dr(b.get("d_first"), b.get("d_last"))}</td>'
                 f'<td class="td-name" style="font-size:12px">{b["status"]}</td>'
                 f'<td class="td-num"><button onclick="event.stopPropagation();jumpToBatch(\'{b["batch"]}\')" '
                 f'style="border:1px solid {C_SEC};background:#fff;color:{C_SEC};border-radius:5px;'
                 f'padding:3px 8px;font-size:10px;font-weight:700;cursor:pointer">🔍 JOURNEY</button></td></tr>')
    warn = ''
    if it['party_differs']:
        warn = (f'<div style="background:#FFF3E0;border-left:4px solid {C_ORG};padding:7px 10px;'
                f'margin-bottom:8px;font-size:12px;color:#E65100">Company differs: plan says '
                f'<strong>{it["party_canon"]}</strong>, RM says <strong>{", ".join(it["rm_customers"])}</strong>'
                f' — normal for loan-licence / P2P work. The RM name is treated as correct.</div>')
    return (f'<tr id="plan-d-{idx}" style="display:none"><td colspan="11" style="background:#FAFDFC;padding:10px 14px">'
            f'{warn}'
            f'<div style="font-size:12px;font-weight:700;color:{C_PRI};margin-bottom:5px">'
            f'Batches dispensed by RM for this plan item — verified chain</div>'
            f'<table style="width:100%;min-width:720px"><thead><tr class="th-row">'
            f'<th>BATCH</th><th>RM DATE</th><th>COMPANY (FROM RM)</th><th>BATCH SIZE</th>'
            f'<th>FILLED (DATES)</th><th>PACKED (DATES)</th><th>DISPATCHED (DATES)</th><th>STATUS</th><th></th>'
            f'</tr></thead><tbody>{rows}</tbody></table></td></tr>')

def dispense_schedule_html():
    """The store's work queue: every plan item with a target dispense date that
    is not yet dispensed, grouped Overdue / Today / Tomorrow / Coming up."""
    due = [x for x in PLAN_ITEMS if x['due_bucket']]
    if not due:
        return ''
    groups = [('overdue',  '🔴 OVERDUE — was due earlier',  '#FFEBEE', C_AMB),
              ('today',    '🟠 DISPENSE TODAY',             '#FFF3E0', C_ORG),
              ('tomorrow', '🟡 DISPENSE TOMORROW',          '#FFFDE7', '#F9A825'),
              ('later',    '⚪ COMING UP',                  '#FAFAFA', '#78909C')]
    blocks = ''
    for key, label, bg, col_ in groups:
        rows_g = [x for x in due if x['due_bucket'] == key]
        if not rows_g:
            continue
        trs = ''
        for it in rows_g:
            when = it['rm_date'].strftime('%d %b') if it['rm_date'] else '—'
            st = it.get('rm_status') or '—'
            by = it.get('updated_by') or ''
            trs += (f'<tr style="background:#fff">'
                    f'<td class="td-name" style="font-weight:600">{it["product"]}</td>'
                    f'<td class="td-name" style="color:#546E7A">{it["party_canon"]}</td>'
                    f'<td class="td-num" style="font-weight:700">{n(it["planned_units"])}</td>'
                    f'<td class="td-name">{it.get("pack") or "—"}</td>'
                    f'<td class="td-num" style="font-weight:700;color:{col_}">{when}</td>'
                    f'<td class="td-name" style="font-size:12px">{st}'
                    + (f'<div style="font-size:10px;color:#B0BEC5">{by}</div>' if by else '')
                    + '</td></tr>')
        blocks += (f'<div style="background:{bg};border-radius:8px;padding:10px 12px;margin:0 14px 10px">'
                   f'<div style="font-weight:700;color:{col_};font-size:13px;margin-bottom:6px">'
                   f'{label} ({len(rows_g)})</div>'
                   f'<div class="tbl-wrap" style="padding:0"><table style="min-width:640px">'
                   f'<thead><tr class="th-row"><th>PRODUCT</th><th>COMPANY</th><th>QTY</th>'
                   f'<th>PACK</th><th>DISPENSE ON</th><th>STATUS / BY</th></tr></thead>'
                   f'<tbody>{trs}</tbody></table></div></div>')
    return f'''
<details class="card" open id="dispense-card">
  <summary>{sec(f'  ━━&nbsp;&nbsp;📋 RM &nbsp; DISPENSING &nbsp; SCHEDULE &nbsp; — &nbsp; ORDERED &nbsp; BY &nbsp; PLANT &nbsp; HEAD &nbsp;({len(due)})&nbsp;━━', C_ORG)}</summary>
  <div style="font-size:12px;color:#607D8B;padding:8px 16px 8px">
    Items the Plant Head / store have given a <strong>DISPENSE ON</strong> date in the AUG PLAN tab and
    that RM has not dispensed yet. An item disappears from this list automatically the moment its RM
    dispensing is logged. Overdue first, then today, then tomorrow.
  </div>
  {blocks}
</details>'''


def pending_plan_html():
    """Card: June–July plan items that still had unproduced quantity, with a
    live 'produced since?' status so the Director can see what is truly left."""
    if not PENDING_PLAN or not PENDING_PLAN['items']:
        return ''
    s = PENDING_PLAN['summary']
    def _tile(lbl, val, col):
        return (f'<div class="tile"><div class="tile-v" style="color:{col}">{val}</div>'
                f'<div class="tile-l">{lbl}</div></div>')
    tiles = (_tile('PENDING ITEMS', s['pending'], C_ORG)
             + _tile('UNITS PENDING', f"{s['units_pending']:,}", C_ORG)
             + _tile('CLEARED SINCE', s['done'], C_SEC)
             + _tile('TOTAL LINES', s['total'], C_PRI))
    rows = ''
    for it in PENDING_PLAN['items']:
        done = it['done']
        if done:
            badge = (f'<span style="background:#E8F5E9;color:#2E7D32;border-radius:3px;'
                     f'padding:1px 6px;font-size:11px;font-weight:700">✅ produced '
                     f'{it["last_fill"].strftime("%d %b")}</span>')
        elif it['last_fill']:
            badge = (f'<span style="background:#FFF8E1;color:#F57F17;border-radius:3px;'
                     f'padding:1px 6px;font-size:11px;font-weight:700">⚪ pending · last run '
                     f'{it["last_fill"].strftime("%d %b")}</span>')
        else:
            badge = ('<span style="background:#FBE9E7;color:#BF360C;border-radius:3px;'
                     'padding:1px 6px;font-size:11px;font-weight:700">⚪ not yet produced</span>')
        bg = '#FAFDFC' if done else '#fff'
        rows += (f'<tr data-done="{1 if done else 0}" style="background:{bg}">'
                 f'<td class="td-name">{it["plan_dt"].strftime("%d %b") if it["plan_dt"] else "—"}</td>'
                 f'<td class="td-name" style="font-weight:600">{it["product"]}</td>'
                 f'<td class="td-name" style="color:#546E7A">{it["pack"]}</td>'
                 f'<td class="td-num" style="font-weight:700">{it["pending"]:,}</td>'
                 f'<td class="td-name">{badge}</td></tr>')
    chips = (
        '<span class="chip active" onclick="event.stopPropagation();pendFilter(this,\'all\')">ALL</span>'
        '<span class="chip" onclick="event.stopPropagation();pendFilter(this,\'pending\')">STILL PENDING</span>'
        '<span class="chip" onclick="event.stopPropagation();pendFilter(this,\'done\')">CLEARED SINCE</span>')
    return f'''
<details class="card" id="pending-card">
  <summary>{sec('  ━━&nbsp;&nbsp;⏳ PENDING FROM JUNE–JULY PLAN &nbsp;—&nbsp; CARRY-OVER &nbsp;━━', C_ORG)}</summary>
  <div style="font-size:12px;color:#607D8B;padding:8px 16px 0">
    Products that were <strong>planned in June–July</strong> but still had unproduced quantity when the plan
    was exported. Each line is checked against real production: it shows <strong>✅ produced</strong> once that
    product was filled on or after its plan date, otherwise it stays <strong>⚪ pending</strong>
    (with the last time it ran, if ever). Source: {PENDING_PLAN['source']}.
  </div>
  <div class="tile-row">{tiles}</div>
  <div class="chip-row" style="padding:0 16px 10px">{chips}</div>
  <div class="tbl-wrap">
    <table style="min-width:640px">
      <thead><tr class="th-row"><th>PLAN DATE</th><th>PRODUCT</th><th>PACK</th>
        <th>PENDING</th><th>STATUS</th></tr></thead>
      <tbody id="pending-rows">{rows}</tbody>
    </table>
  </div>
</details>'''


def plan_section_html():
    if not PLAN_ITEMS:
        return ''
    s = PLAN_SUMMARY
    pct_started = (s['started'] / s['items'] * 100) if s['items'] else 0
    tiles = (
        tile('PLAN ITEMS', n(s['items']), f'{n(s["units"])} units planned in total', C_PRI)
      + tile('STARTED', f"{s['started']} / {s['items']}", f'{pct_started:.0f}% begun · {s["batches"]} RM batches linked', C_SEC)
      + tile('COMPLETED', n(s['done']), 'items fully dispatched (≥95% of plan)', C_GRN)
      + tile('UNITS FILLED VS PLAN', n(s['filled']), f'of {n(s["units"])} planned', C_AMB)
      + tile('MARKED BY RM / PLANT HEAD', n(s['written']),
             (f'{s["next"]} marked to dispense next · {s["flags"]} need attention'
              if s['written'] else 'add the RM STATUS column to the plan tab to use this'), '#7B1FA2')
      + tile('DISPENSING SCHEDULE', n(s['overdue'] + s['due_today'] + s['due_tomorrow'] + s['due_later']),
             f'{s["overdue"]} overdue · {s["due_today"]} today · {s["due_tomorrow"]} tomorrow', C_ORG)
      + tile('↩ JUN–JUL CARRY-OVER', n(s.get('carry_items', 0)),
             f'{n(s.get("carry_units", 0))} units still pending from earlier plans', '#F57F17')
    )
    rows = ''
    for i, it in enumerate(PLAN_ITEMS):
        bg = ('#FFEBEE' if it['due_bucket'] == 'overdue' else
              '#FFF3E0' if it['due_bucket'] == 'today' else
              '#FFFDE7' if it['due_bucket'] == 'tomorrow' else
              '#F1F8F6' if i % 2 == 0 else '#FFFFFF')
        prio = it['priority'] or '—'
        lots = f' <span style="color:#90A4AE;font-size:10px">({it["lots"]} lots)</span>' if it['lots'] > 1 else ''
        nb = len(it['batches'])
        badge = (f'<span style="background:{C_LBG};color:{C_PRI};border-radius:3px;padding:1px 5px;'
                 f'font-size:10px;font-weight:700">{nb} batch{"es" if nb != 1 else ""}</span>') if nb else ''
        pflag = (f' <span title="The plan tab names a different company — RM name is treated as correct" '
                 f'style="background:#FFF3E0;color:#E65100;border-radius:3px;padding:1px 5px;'
                 f'font-size:10px;font-weight:700">⇄ plan: {it["party_canon"]}</span>') if it['party_differs'] else ''
        # what the store / Plant Head wrote in the sheet
        _stt = it.get('rm_status') or ''
        _sl = _stt.lower()
        _scol = (C_GRN if _sl.startswith(('dispens', 'done', 'complete')) else
                 C_ORG if ('next' in _sl or 'tomorrow' in _sl) else
                 '#7B1FA2' if _sl else '#90A4AE')
        wrote = (f'<span style="background:#F3E5F5;color:{_scol};border-radius:3px;padding:2px 7px;'
                 f'font-size:11px;font-weight:700">{_stt}</span>') if _stt else '<span style="color:#CFD8DC">—</span>'
        if it.get('rm_date'):
            wrote += f'<div style="font-size:10px;color:#90A4AE">{it["rm_date"].strftime("%d %b")}</div>'
        if it.get('updated_by'):
            wrote += f'<div style="font-size:10px;color:#B0BEC5">{it["updated_by"]}</div>'
        flagchip = (f'<div style="font-size:10px;color:{C_AMB};margin-top:3px">⚠ {it["flag"]}</div>'
                    ) if it.get('flag') else ''
        def _mini(pct_v, col, lbl, tip):
            hpx = 0 if pct_v <= 0 else max(3, min(22, round(22 * pct_v / 100)))
            return (f'<div title="{tip}" style="display:flex;flex-direction:column;align-items:center;gap:1px">'
                    f'<div style="width:15px;height:22px;background:#ECEFF1;border-radius:3px;'
                    f'display:flex;align-items:flex-end;overflow:hidden">'
                    f'<div style="width:100%;height:{hpx}px;background:{col}"></div></div>'
                    f'<span style="font-size:8px;color:#90A4AE;font-weight:700">{lbl}</span></div>')
        _pq = it['planned_units'] or 0
        _pctf = (it['filled'] / _pq * 100) if _pq else 0
        _pctp = (it['packed'] / _pq * 100) if _pq else 0
        _pctd = (it['dispatched'] / _pq * 100) if _pq else 0
        bar = ('<div style="display:flex;gap:3px">'
               + _mini(100 if it['batches'] else 0, C_PRI, 'RM',
                       f'RM dispensed: {len(it["batches"])} batch(es)' if it['batches'] else 'RM: not dispensed yet')
               + _mini(_pctf, C_SEC, 'F', f'Filled {it["filled"]:,.0f} of {_pq:,.0f} ({_pctf:.0f}%)')
               + _mini(_pctp, C_AMB, 'P', f'Packed {it["packed"]:,.0f} of {_pq:,.0f} ({_pctp:.0f}%)')
               + _mini(_pctd, C_ORG, 'D', f'Dispatched {it["dispatched"]:,.0f} of {_pq:,.0f} ({_pctd:.0f}%)')
               + '</div>')
        _isnext = 1 if ('next' in _sl or 'tomorrow' in _sl) else 0
        _mon = it.get('month', 'AUG')
        _mcol = {'AUG': ('#E0F2F1', '#00695C'), 'JUL': ('#FFF8E1', '#F57F17'),
                 'JUN': ('#FBE9E7', '#BF360C')}[_mon if _mon in ('AUG', 'JUL', 'JUN') else 'AUG']
        _mchip = (f'<span style="background:{_mcol[0]};color:{_mcol[1]};border-radius:3px;'
                  f'padding:1px 5px;font-size:10px;font-weight:800;margin-right:5px">{_mon}</span>')
        _carry = ''
        if it.get('carry_months'):
            _carry = (f' <span title="This product also had {it["carry_units"]:,} units pending '
                      f'from the {it["carry_months"]} plan — carried into this month" '
                      f'style="background:#FFF8E1;color:#F57F17;border-radius:3px;padding:1px 5px;'
                      f'font-size:10px;font-weight:700">↩ carried from {it["carry_months"]} '
                      f'({it["carry_units"]:,})</span>')
        rows += (f'<tr style="background:{bg};cursor:pointer" data-prio="{it["priority"] or 0}" '
                 f'data-srank="{it["srank"]}" data-next="{_isnext}" data-flag="{1 if it.get("flag") else 0}" '
                 f'data-month="{_mon}" '
                 f'data-company="{(it.get("display_party") or "").lower()}" '
                 f'onclick="togglePlan({i})" title="Click to verify the RM batches behind this item">'
                 f'<td class="td-num" style="font-weight:700;color:{C_PRI}">{prio}</td>'
                 f'<td class="td-name">{_mchip}{it["product"]}{lots} {badge}{_carry}</td>'
                 f'<td class="td-name" style="color:#546E7A">{it.get("display_party") or "—"}{pflag}</td>'
                 f'<td class="td-name" style="color:#37474F">{it.get("pack") or "—"}</td>'
                 f'<td class="td-num" style="font-weight:700">{n(it["planned_units"])}</td>'
                 f'<td class="td-num" style="color:{C_SEC}">{n(it["filled"]) if it["filled"] else "—"}</td>'
                 f'<td class="td-num" style="color:{C_AMB}">{n(it["packed"]) if it["packed"] else "—"}</td>'
                 f'<td class="td-num" style="color:{C_ORG}">{n(it["dispatched"]) if it["dispatched"] else "—"}</td>'
                 f'<td>{bar}</td>'
                 f'<td class="td-name" style="font-size:12px;white-space:nowrap">{it["status"]}{flagchip}</td>'
                 f'<td class="td-name">{wrote}</td>'
                 f'</tr>')
        rows += _plan_batch_detail(it, i)
    _sc = {k: sum(1 for x in PLAN_ITEMS if x['srank'] == k) for k in range(6)}
    chips = ''.join(
        f'<span class="chip" onclick="event.stopPropagation();planFilter(this,{p})">{lbl}</span>'
        for p, lbl in [(-1, 'ALL'), (-6, 'AUG PLAN'), (-7, '↩ JUN–JUL CARRY-OVER'),
                       (1, 'PRIORITY 1'), (2, 'PRIORITY 2'), (3, 'PRIORITY 3'),
                       (4, 'PRIORITY 4'), (-2, f'NOT STARTED ({_sc[0]})'),
                       (11, f'🟤 RM DISPENSED ({_sc[1]})'),
                       (12, f'🔵 FILLING ({_sc[2]})'),
                       (13, f'🟡 PACKED ({_sc[3]})'),
                       (14, f'🟠 DISPATCHING ({_sc[4]})'),
                       (15, f'✅ COMPLETED ({_sc[5]})'),
                       (-4, '🟣 DISPENSE NEXT'), (-5, '⚠ NEEDS ATTENTION')])
    off = ''
    if PLAN_OFF:
        _futwarn = (' <span style="background:#FBE9E7;color:#BF360C;border-radius:3px;'
                    'padding:1px 5px;font-size:10px;font-weight:700">⚠ future date — check RM entry</span>')
        orows = ''.join(
            f'<tr><td class="td-name" style="font-weight:600">{b["batch"]}</td>'
            f'<td class="td-name">{b["date"].strftime("%d %b")}{_futwarn if b.get("future") else ""}</td>'
            f'<td class="td-name">{b["product"] or "—"}</td>'
            f'<td class="td-name" style="color:#546E7A">{b["customer"] or "—"}</td></tr>'
            for b in PLAN_OFF)
        off = (f'<div style="padding:10px 16px 0;font-size:12px;font-weight:700;color:{C_AMB}">'
               f'⚠ Dispensed by RM but not matched to any plan line ({len(PLAN_OFF)}) — '
               f'either extra production, or the product name differs from the plan. '
               f'Batches whose filling already began before {PLAN_WINDOW_FROM.strftime("%d %b")} are '
               f'last month\'s work and are not listed here.</div>'
               f'<div class="tbl-wrap"><table style="min-width:520px"><thead><tr class="th-row">'
               f'<th>BATCH</th><th>RM DATE</th><th>PRODUCT (RM)</th><th>COMPANY (RM)</th></tr></thead>'
               f'<tbody>{orows}</tbody></table></div>')
    return f'''
<details class="card" id="plan-card">
  <summary>{sec(f'  ━━&nbsp;&nbsp;🗓 {PLAN_TITLE} &nbsp;—&nbsp; PLANNED &nbsp; vs &nbsp; ACTUAL &nbsp;━━', C_PRI)}</summary>
  <div style="font-size:12px;color:#607D8B;padding:8px 16px 0">
    Plan source: <strong>{PLAN_SOURCE}</strong>, consolidated with the <strong>June–July pending plan</strong> —
    the month chip on each row shows which plan it belongs to, and AUG items that were also pending
    earlier carry an <strong>↩ carried from</strong> tag. <strong>Click any row</strong> to see the RM batches behind it
    and verify the chain. Plan lines carry no batch number — the link is made on <strong>product name</strong>,
    and the <strong>batch number and company name are taken from RM</strong> (under loan-licence / P2P the plan
    company can differ; that is marked ⇄, not treated as an error). Filling, packing and dispatch are then
    tracked by batch number, so shifting dates never break the tracking.
    <br><strong>RM store / Mr. Verma write directly in the plan tab</strong> (RM STATUS, RM DATE, BATCH NO.,
    UPDATED BY) — those entries show in the last column, are reconciled against the logs, and any change is
    emailed to the store team automatically.
  </div>
  <div class="tile-row">{tiles}</div>
  <div style="padding:2px 18px 12px">
    <div style="font-weight:700;color:{C_PRI};font-size:12px;margin-bottom:5px">HOW THE WHOLE PLAN IS MOVING (units, % of {n(s['units'])} planned)</div>
    <table style="width:100%;font-size:12px;border-collapse:collapse">
      <tr><td style="width:110px;color:#546E7A;padding:2px 0">RM started</td>
          <td><div style="display:inline-block;background:{C_PRI};color:#fff;padding:2px 8px;border-radius:4px;min-width:60px;width:{(s['rm_started']/s['items']*100) if s['items'] else 0:.0f}%;box-sizing:border-box">{s['rm_started']} of {s['items']} items</div></td></tr>
      <tr><td style="color:#546E7A;padding:2px 0">Filled</td>
          <td><div style="display:inline-block;background:{C_SEC};color:#fff;padding:2px 8px;border-radius:4px;min-width:60px;width:{min(100,(s['filled']/s['units']*100) if s['units'] else 0):.0f}%;box-sizing:border-box">{n(s['filled'])} &nbsp;({(s['filled']/s['units']*100) if s['units'] else 0:.1f}%)</div></td></tr>
      <tr><td style="color:#546E7A;padding:2px 0">Packed</td>
          <td><div style="display:inline-block;background:{C_AMB};color:#fff;padding:2px 8px;border-radius:4px;min-width:60px;width:{min(100,(s['packed_sum']/s['units']*100) if s['units'] else 0):.0f}%;box-sizing:border-box">{n(s['packed_sum'])} &nbsp;({(s['packed_sum']/s['units']*100) if s['units'] else 0:.1f}%)</div></td></tr>
      <tr><td style="color:#546E7A;padding:2px 0">Dispatched</td>
          <td><div style="display:inline-block;background:{C_ORG};color:#fff;padding:2px 8px;border-radius:4px;min-width:60px;width:{min(100,(s['disp_sum']/s['units']*100) if s['units'] else 0):.0f}%;box-sizing:border-box">{n(s['disp_sum'])} &nbsp;({(s['disp_sum']/s['units']*100) if s['units'] else 0:.1f}%)</div></td></tr>
    </table>
  </div>
  <div style="padding:0 16px 8px">
    <input id="plan-search" type="search" placeholder="🔍 Search the plan — product, customer, pack, batch…"
           oninput="planSearch(this.value)" onclick="event.stopPropagation()"
           style="width:100%;max-width:430px;padding:7px 12px;border:1px solid #B0BEC5;border-radius:6px;font-size:13px;color:#37474F">
    <span id="plan-search-count" style="font-size:12px;color:#607D8B;margin-left:8px"></span>
  </div>
  <div class="chip-row" style="padding:0 16px 10px">{chips}</div>
  <div class="tbl-wrap">
    <table style="min-width:880px">
      <thead><tr class="th-row">
        <th>P</th><th>PRODUCT</th><th>CUSTOMER</th><th>PACK</th><th>PLANNED</th>
        <th>FILLED</th><th>PACKED</th><th>DISPATCHED</th><th>PIPELINE (RM→F→P→D)</th><th>TRACKED STATUS</th>
        <th>RM / PLANT HEAD ENTRY</th>
      </tr></thead>
      <tbody id="plan-rows">{rows}</tbody>
    </table>
  </div>
  {off}
</details>'''

# ── Monthly summary (RM → Fill → Pack → Disp) ─────────────────────────────────
# Tracking-start cutoff: dispenses before this date were from the pre-system
# era and are excluded so they don't pollute the monthly numbers.
TRACKING_START = date(2026, 5, 11)

def _monthly_summary():
    try:
        rm_all = pd.read_excel(TEMPLATE, sheet_name='➕ RM Dispensing Log', header=3)
        rm_all.columns = [' '.join(str(c).split()) for c in rm_all.columns]  # collapse newlines
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

        # Drill-downs for the type table:
        #   Bottle → per pack size
        #   Pouches / Ointment / External → per BRAND name (first word of the
        #   product, so Redsun Mango/Mint/... all fold into "Redsun" — no
        #   flavour bifurcation, per the Director).
        def _brand(pname):
            s2 = str(pname or '').strip()
            if not s2 or s2.lower() == 'nan':
                return '(name n/a)'
            w = re.split(r'[\s(]+', s2)[0].strip(' .-,')
            w = re.sub(r'-\d+$', '', w)          # COMIT-100 → COMIT
            return w.title() if w else '(name n/a)'
        pt_drill = {pt: {} for pt in PRODUCT_TYPES}
        for df_x, qty_col, ptype_col, stage in [
            (f_m, 'Qty',         'ProductType', 'f'),
            (p_m, 'TotalPacked', 'ProdType',    'p'),
            (d_m, 'Qty',         'ProductType', 'd'),
        ]:
            for _, r in df_x.iterrows():
                pt = r.get(ptype_col)
                if pt not in pt_drill: continue
                q = float(pd.to_numeric(r.get(qty_col), errors='coerce') or 0)
                if not q: continue
                sz = r.get('PackSize')
                sz_n = (re.sub(r'\s+', '', str(sz)).upper()
                        if sz is not None and str(sz).strip() not in ('', 'nan', 'None') else '')
                if pt == 'Bottle':
                    key = sz_n or '(size n/a)'
                else:
                    # brand + size kept separately so multi-size brands can be split
                    key = (_brand(r.get('Product')), sz_n)
                pt_drill[pt].setdefault(key, {'f': 0.0, 'p': 0.0, 'd': 0.0})
                pt_drill[pt][key][stage] += q

        # Flatten non-bottle drills: one line per brand — but when a brand comes
        # in DIFFERENT pack sizes, one line per size (each size has its own rate).
        for _pt in list(pt_drill):
            if _pt == 'Bottle':
                continue
            per_brand = {}
            for (br, sz_n), v in pt_drill[_pt].items():
                per_brand.setdefault(br, {}).setdefault(sz_n, {'f': 0.0, 'p': 0.0, 'd': 0.0})
                for st2 in ('f', 'p', 'd'):
                    per_brand[br][sz_n][st2] += v[st2]
            flat = {}
            for br, sizes in per_brand.items():
                if len(sizes) == 1:
                    sz_n = next(iter(sizes))
                    label = f'{br} ({sz_n})' if sz_n else br
                    flat[label] = sizes[sz_n]
                else:
                    for sz_n, v in sizes.items():
                        flat[f'{br} — {sz_n or "size n/a"}'] = v
            pt_drill[_pt] = flat

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
            'pt_drill':   pt_drill,
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

def _stage_json(dates_map, key):
    d = dates_map.get(key)
    return {'first': safe(d['first']), 'last': safe(d['last'])} if d else None

def _rm_json(key):
    r = RM_INFO.get(key)
    if not r:
        return None
    return {'date': safe(r['date']), 'customer': r['customer'],
            'product': r['product'], 'size': float(r['size'])}

_stuck_by_key = { _bkey(s['batch']): (s['stuck_stage'], s['stuck_days']) for s in STUCK_BATCHES }

batch_rows = []
for e in BATCH_JOURNEY:
    k = _bkey(e['batch'])
    st = _stuck_by_key.get(k)
    batch_rows.append(
        {'batch':e['batch'], 'product':e['product'], 'ptype':e['ptype'],
         'party':e.get('party'), 'packSize':e.get('packsize'),
         'filled':float(e['filled']), 'packed':float(e['packed']), 'dispatched':float(e['dispatched']),
         'status':e['status'],
         'rm': _rm_json(k),
         'fillD': _stage_json(FILL_DATES, k),
         'packD': _stage_json(PACK_DATES, k),
         'dispD': _stage_json(DISP_DATES, k),
         'stuck': {'stage': st[0], 'days': st[1]} if st else None})

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
    # Month-scoped (the *_m dicts) so the server-rendered default view matches
    # the single-month scope everywhere else on the page.
    rows = ''
    for i, pt in enumerate(PRODUCT_TYPES):
        fv = fill_by_type_m.get(pt, 0)
        pv = pack_by_type_m.get(pt, 0)
        dv = disp_by_type_m.get(pt, 0)
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

# ── Server-rendered initial content ──────────────────────────────────────────
# The tiles and tables below used to ship as "—" placeholders that JavaScript
# filled in on load. Any viewer that blocks or fails to run JS (Drive/mail
# previews, PDF export, old WebViews) saw a wall of dashes. The generator
# already knows every number, so render the full-period view straight into the
# HTML; the JS date-filter simply re-renders on top of it.
_NO_DATA = ('<tr><td colspan="{cols}" style="text-align:center;color:#90A4AE;'
            'padding:14px">No data loaded for this period ({period}).</td></tr>')

def _line_order_key(name):
    """Same display order as the JS cmpLine(): Line No 1..5, then specials."""
    s = str(name or '').strip()
    m = re.match(r'^line\s*no\.?\s*0*(\d+)', s, re.I)
    if m: return (0, int(m.group(1)), '')
    special = {'flat sachet': 1, 'stick pack sachet': 2, 'sachet': 3,
               'ointment': 4, 'external': 5}
    return (1, special.get(s.lower(), 99), s.lower())

def grouped_stage_rows(df, qty_field):
    """Filling/Packing detail rows grouped by Line+Product+PackSize+Party —
    mirrors the JS renderFilling/renderPacking grouping exactly.
    Scoped to the default single-month view so glance and detail agree."""
    d = filt(df, _glance_start, _glance_end)
    if not len(d):
        return _NO_DATA.format(cols=5, period=_glance_month_label)
    g = {}
    for _, r in d.iterrows():
        ln = str(r.get('Line') or '—'); pr = str(r.get('Product') or '—')
        ps = r.get('PackSize'); ps = '—' if ps is None or (isinstance(ps, float) and pd.isna(ps)) else str(ps)
        pa = str(r.get('Party') or '—')
        q = float(pd.to_numeric(r.get(qty_field), errors='coerce') or 0)
        g[(ln, pr, ps, pa)] = g.get((ln, pr, ps, pa), 0) + q
    rows = ''
    for i, ((ln, pr, ps, pa), q) in enumerate(sorted(
            g.items(), key=lambda kv: (_line_order_key(kv[0][0]), kv[0][1].lower(),
                                       kv[0][2], kv[0][3].lower()))):
        bg = '#F1F8F6' if i % 2 == 0 else '#FFFFFF'
        rows += (f'<tr style="background:{bg}"><td class="td-name">{ln}</td>'
                 f'<td class="td-name">{pr}</td>'
                 f'<td class="td-name" style="color:#37474F">{ps}</td>'
                 f'<td class="td-name" style="color:#546E7A">{pa}</td>'
                 f'<td class="td-num" style="color:{C_AMB};font-weight:700">{n(q)}</td></tr>')
    return rows

def party_mtd_rows():
    """Party-wise dispatch rows (party + qty), month-scoped like the default view."""
    if not len(_g_disp):
        return _NO_DATA.format(cols=2, period=_glance_month_label)
    rows = ''
    for i, (p, v) in enumerate(party_month.items()):
        if not str(p).strip():
            continue
        bg = '#FFF8F1' if i % 2 == 0 else '#FFFFFF'
        rows += (f'<tr style="background:{bg}"><td class="td-name">{p}</td>'
                 f'<td class="td-num" style="color:{C_ORG};font-weight:700">{n(v)}</td></tr>')
    return rows or _NO_DATA.format(cols=2, period=_glance_month_label)

def stuck_batch_rows():
    if not STUCK_BATCHES:
        return ('<tr><td colspan="6" style="text-align:center;color:#90A4AE;padding:14px">'
                'No stuck batches — everything filled is moving to packing, and everything '
                f'packed is dispatching within {STUCK_PACK_DAYS} days. ✅</td></tr>')
    rows = ''
    for i, s in enumerate(STUCK_BATCHES):
        bg = '#FFF8F1' if i % 2 == 0 else '#FFFFFF'
        qty = s['packed'] if s['packed'] > 0 else s['filled']
        rows += (f'<tr style="background:{bg}">'
                 f'<td class="td-name" style="font-weight:600">{s["batch"]}</td>'
                 f'<td class="td-name">{s["product"] or "—"}</td>'
                 f'<td class="td-name" style="color:#546E7A">{s["party"] or "—"}</td>'
                 f'<td class="td-name">{s["stuck_stage"]}</td>'
                 f'<td class="td-num" style="color:{C_AMB};font-weight:700">{s["stuck_days"]} days</td>'
                 f'<td class="td-num">{n(qty)}</td></tr>')
    return rows

# ── Director summary (top-of-page headline layer) ─────────────────────────────
_pipeline_units = IN_STOCK_UNITS   # packed, waiting in the store
_attention_n = len(STUCK_BATCHES) + _bj_problems

def _attention_lines():
    """Up to three plain-language attention items for the summary layer."""
    items = []
    if STUCK_BATCHES:
        w = STUCK_BATCHES[0]
        items.append(f'Longest wait: batch <strong>{w["batch"]}</strong> '
                     f'({w["product"] or "unknown product"}) has been '
                     f'{w["stuck_stage"].lower()} for <strong>{w["stuck_days"]} days</strong>.')
        if len(STUCK_BATCHES) > 1:
            items.append(f'{len(STUCK_BATCHES)} batches in total are waiting longer than '
                         f'normal ({STUCK_FILL_DAYS}+ days to pack, {STUCK_PACK_DAYS}+ days to dispatch) '
                         '— full list in "Batches needing attention" below.')
    if _bj_problems:
        items.append(f'{_bj_problems} batch record(s) look inconsistent (dispatched or packed '
                     'with no filling record, or shipped more than was made) — likely a batch-number '
                     'typo in one of the logs. Searchable in "Find a batch" below.')
    if not items:
        items.append('Nothing unusual — production is flowing and no batch is waiting longer than normal.')
    return items

# All unit/batch cards in the glance show ONE month only. Default = the latest
# month that actually has dispatches (a brand-new month with zero rows yet would
# read as an alarming "0"); the JS date-filter re-points every card at whichever
# month the selected day belongs to.
_disp_months  = sorted({d.strftime('%Y-%m') for d in cur(disp_df)['Date']})
_g_mkey       = _disp_months[-1] if _disp_months else f'{YEAR}-{MONTH:02d}'
_g_y, _g_m    = int(_g_mkey[:4]), int(_g_mkey[5:7])
_glance_start = date(_g_y, _g_m, 1)
_glance_end   = date(_g_y, _g_m, calendar.monthrange(_g_y, _g_m)[1])
_g_fill       = filt(fill_df, _glance_start, _glance_end)
_g_pack       = filt(pack_df, _glance_start, _glance_end)
_g_disp       = filt(disp_df, _glance_start, _glance_end)
f_month       = _g_fill['Qty'].sum()
p_month       = _g_pack['TotalPacked'].sum()
d_month       = _g_disp['Qty'].sum()
d_month_cust  = sum(1 for v in _g_disp.groupby('Party')['Qty'].sum().values if v > 0)

# Month-scoped batch counts, defined to match the batch-journey statuses:
#   completed = dispatched this month AND has fill+pack records (full journey)
#   pipeline  = fill/pack activity this month, nothing dispatched yet
_journey_by_key = {_bkey(e['batch']): e for e in BATCH_JOURNEY}
_g_disp_keys = {_bkey(b) for b in _g_disp['Batch'].dropna() if str(b).strip()}
comp_month = sum(1 for k in _g_disp_keys
                 if k in _journey_by_key
                 and _journey_by_key[k]['filled'] > 0 and _journey_by_key[k]['packed'] > 0)
_PIPE_STATUSES = ('Filled & packed (in stock)', 'Filled only')
_g_act_keys = {_bkey(b) for b in list(_g_fill['Batch'].dropna()) + list(_g_pack['Batch'].dropna())
               if str(b).strip()}
pipe_month = sum(1 for k in _g_act_keys
                 if _journey_by_key.get(k, {}).get('status') in _PIPE_STATUSES)
_glance_month_label = _glance_start.strftime('%B %Y').upper()

# ── ONE PERIOD STORY ──────────────────────────────────────────────────────────
# The DEFAULT page view is one month (the glance month above) so the summary
# cards and the detail sections always agree. The date filter offers each month,
# "both months", and single days; the JS re-renders every section to the same
# scope. Month-scoped values below are what the server renders initially.
f_rec_m   = len(_g_fill)
f_avg_m   = f_month / f_rec_m if f_rec_m else 0
f_lines_m = _g_fill['Line'].nunique()
p_ratio_m = p_month / f_month if f_month else 0
d_ratio_m = d_month / f_month if f_month else 0
_g_staff  = filt(staff_df, _glance_start, _glance_end)
s_fem_m   = _g_staff['Female'].mean() if len(_g_staff) else 0
s_male_m  = _g_staff['Male'].mean()   if len(_g_staff) else 0
party_month = _g_disp.groupby('Party')['Qty'].sum().sort_values(ascending=False)
fill_by_type_m = _g_fill.groupby('ProductType')['Qty'].sum()
pack_by_type_m = _g_pack.groupby('ProdType')['TotalPacked'].sum()
disp_by_type_m = _g_disp.groupby('ProductType')['Qty'].sum()

# ── ONE STOCK STORY ───────────────────────────────────────────────────────────
# Two clearly-named stock figures (both all-time, independent of the filter):
#   packed stock  = packed, not yet dispatched  (matches the stock list section)
#   filling WIP   = filled, not yet packed (work in progress on the floor)
WIP_UNITS = sum(max(e['filled'] - e['packed'], 0) for e in BATCH_JOURNEY)

def _glance_tile(label, vid, value, sid, subtext, color):
    return (f'<div class="tile"><div class="tlabel">{label}</div>'
            f'<div class="tvalue" id="{vid}" style="color:{color}">{value}</div>'
            f'<div class="tsub" id="{sid}">{subtext}</div></div>')

def director_summary_html():
    cards = (
        _glance_tile('UNITS FILLED', 'glance-fill', n(f_month),
                     'glance-fill-sub', 'filled this month', C_SEC)
      + _glance_tile('UNITS PACKED', 'glance-pack', n(p_month),
                     'glance-pack-sub', 'packed this month', C_AMB)
      + _glance_tile('UNITS DISPATCHED', 'glance-disp', n(d_month),
                     'glance-disp-sub', f'sent to {d_month_cust} customers this month', C_ORG)
      + _glance_tile('BATCHES COMPLETED', 'glance-comp', n(comp_month),
                     'glance-comp-sub', 'made, packed & dispatched this month', C_GRN)
      + _glance_tile('BATCHES IN THE PIPELINE', 'glance-pipe', n(pipe_month),
                     'glance-pipe-sub', 'filled or packed this month, awaiting dispatch', C_SEC)
      + tile('READY IN THE STORE', n(len(IN_STOCK)),
             f'{n(IN_STOCK_UNITS)} packed units in the BSR (Bonded Store Room) awaiting dispatch — list below', C_SEC)
    )
    notes = ''.join(f'<li style="margin:3px 0">{t}</li>' for t in _attention_lines())
    return f'''
<div class="card">
  {sec('  ━━&nbsp;&nbsp;AT &nbsp; A &nbsp; GLANCE &nbsp;━━', C_PRI)}
  <div style="font-size:12px;color:#607D8B;padding:8px 16px 0">
    Month shown: <strong id="glance-month-note" style="color:{C_PRI}">{_glance_month_label}</strong>
    — the cards follow the date picked in the filter above ("Ready in the store" is today's stock, always current).
  </div>
  <div class="tile-row">{cards}</div>
  <ul style="margin:0 22px 14px 34px;font-size:12.5px;color:#37474F;line-height:1.5">{notes}</ul>
</div>'''

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
    drill = s.get('pt_drill', {})
    mid = s['m'].replace('-', '')
    rows = ''
    for i, p in enumerate(PRODUCT_TYPES):
        bg = '#F1F8F6' if i % 2 == 0 else '#FFFFFF'
        sub = drill.get(p) or {}
        has_drill = bool(sub)
        gid = f'{mid}-{i}'
        caret = (f' <span id="ptc-{gid}" style="color:#90A4AE;font-size:10px">▸ click for '
                 f'{"sizes" if p == "Bottle" else "brands"}</span>') if has_drill else ''
        click = (f' style="background:{bg};cursor:pointer" onclick="togglePTx(\'{gid}\')"'
                 if has_drill else f' style="background:{bg}"')
        rows += (f'<tr{click}><td class="td-name">{p}{caret}</td>'
                 f'<td class="td-num" style="color:{C_SEC};font-weight:600">{n(pt[p]["f"])}</td>'
                 f'<td class="td-num" style="color:{C_AMB};font-weight:600">{n(pt[p]["p"])}</td>'
                 f'<td class="td-num" style="color:{C_ORG};font-weight:600">{n(pt[p]["d"])}</td></tr>')
        if has_drill:
            if p == 'Bottle':
                def _sk(kv):
                    m2 = re.search(r'([\d.]+)', kv[0])
                    return float(m2.group(1)) if m2 else 9999
                entries = sorted(sub.items(), key=_sk)
            else:
                entries = sorted(sub.items(), key=lambda kv: -(kv[1]['f'] + kv[1]['p'] + kv[1]['d']))
            for lbl, v in entries:
                rows += (f'<tr class="ptx-{gid}" style="display:none;background:#FAFDFC">'
                         f'<td class="td-name" style="padding-left:30px;color:#546E7A;font-size:12px">↳ {lbl}</td>'
                         f'<td class="td-num" style="color:{C_SEC};font-size:12px">{n(v["f"]) if v["f"] else "—"}</td>'
                         f'<td class="td-num" style="color:{C_AMB};font-size:12px">{n(v["p"]) if v["p"] else "—"}</td>'
                         f'<td class="td-num" style="color:{C_ORG};font-size:12px">{n(v["d"]) if v["d"] else "—"}</td></tr>')

    pt_html = f'''
    <div style="border-top:1px solid #B0BEC5;padding-top:10px;margin-top:14px">
      <div style="font-weight:700;color:{C_PRI};font-size:13px;margin-bottom:6px">Product type breakdown — this month
        <span style="font-weight:400;font-size:11px;color:#90A4AE">— click Bottle for sizes; pouches, ointment &amp; external for brand names</span></div>
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

# Company logo, embedded as a data URI so the page stays a single file.
try:
    LOGO_URI = ('data:image/png;base64,'
                + base64.b64encode(open(os.path.join(HERE, 'logo.png'), 'rb').read()).decode('ascii'))
except Exception:
    LOGO_URI = ''

# Always stamp in Indian time — the cloud runner is UTC and confused the team
from datetime import timezone as _tz, timedelta as _tdelta
generated_at = datetime.now(_tz(_tdelta(hours=5, minutes=30))).strftime('%d %b %Y, %I:%M %p IST')

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="900">
<title>Enicar Dashboard — {PERIOD}</title>
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;
         background:#ECEFF1; color:#263238; -webkit-font-smoothing:antialiased; }}

  /* ── Header ── */
  .header {{ background:linear-gradient(135deg,#00332B 0%,{C_PRI} 55%,{C_SEC} 100%); color:#fff; padding:18px 28px 14px; }}
  .header h1 {{ font-size:32px; font-weight:900; letter-spacing:3px; line-height:1; }}
  .header-sub {{ font-size:13px; color:#B2DFDB; margin-top:4px; }}
  .period-bar {{ background:{C_SEC}; color:#fff; text-align:center; padding:8px;
                 font-size:13px; font-weight:700; letter-spacing:1px; }}

  /* ── Layout ── */
  .container {{ max-width:1280px; margin:0 auto; padding:16px; }}
  .card {{ background:#fff; border-radius:12px; box-shadow:0 2px 10px rgba(0,38,32,0.10);
           margin-bottom:16px; overflow:hidden; }}

  /* ── Section header ── */
  .sec-hdr {{ color:#fff; font-size:13px; font-weight:700; padding:9px 16px;
              letter-spacing:1px; }}

  /* ── Tiles ── */
  .tile-row {{ display:flex; gap:10px; padding:14px 14px 10px; flex-wrap:wrap; }}
  .tile {{ flex:1; min-width:150px; background:#fff; border:1px solid #E0F2F1;
           border-radius:10px; padding:12px 14px; text-align:center;
           transition:transform .15s ease, box-shadow .15s ease; }}
  .tile:hover {{ transform:translateY(-2px); box-shadow:0 6px 16px rgba(0,38,32,0.12); }}
  .tlabel {{ font-size:9px; font-weight:700; color:{C_SEC}; text-transform:uppercase;
             letter-spacing:0.8px; margin-bottom:6px; }}
  .tvalue {{ font-size:26px; font-weight:700; line-height:1.1; }}
  .tsub {{ font-size:9px; color:#90A4AE; margin-top:5px; font-style:italic; }}

  /* ── Tables ── */
  /* Horizontal scroll on phones — tables keep readable column widths and the
     wrapper scrolls sideways instead of squishing 5-8 columns into 375px. */
  .tbl-wrap {{ padding:0 14px 14px; overflow-x:auto; -webkit-overflow-scrolling:touch; }}
  .tbl-wrap table {{ min-width:600px; }}
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

  /* ── Collapsible detail sections ── */
  details.card > summary {{ list-style:none; cursor:pointer; user-select:none; }}
  details.card > summary::-webkit-details-marker {{ display:none; }}
  details.card > summary .sec-hdr {{ position:relative; }}
  details.card > summary .sec-hdr::after {{ content:'▸ show'; position:absolute; right:16px;
      font-weight:400; font-size:11px; opacity:0.85; }}
  details.card[open] > summary .sec-hdr::after {{ content:'▾ hide'; }}
  .expand-bar {{ text-align:right; margin:-6px 0 10px; }}
  .expand-bar button {{ background:#fff; border:1.5px solid {C_SEC}; color:{C_SEC};
      border-radius:6px; padding:5px 14px; font-size:12px; font-weight:700; cursor:pointer; }}
  .noscript-note {{ background:#FFF3E0; border-left:4px solid {C_ORG}; padding:10px 16px;
      font-size:12px; color:{C_ORG}; margin:0 0 12px; border-radius:6px; }}

  /* ── Batch journey timeline ── */
  .bj-timeline {{ display:flex; align-items:stretch; gap:0; flex-wrap:wrap; padding:10px 14px 14px; }}
  .bj-stage {{ flex:1; min-width:140px; border:1px solid #CFD8DC; border-radius:8px;
      padding:10px 12px; background:#fff; }}
  .bj-stage.missing {{ background:#FAFAFA; border-style:dashed; color:#90A4AE; }}
  .bj-stage .bj-name {{ font-size:10px; font-weight:700; letter-spacing:0.6px;
      text-transform:uppercase; color:{C_SEC}; margin-bottom:4px; }}
  .bj-stage.missing .bj-name {{ color:#90A4AE; }}
  .bj-qty {{ font-size:17px; font-weight:700; color:#263238; }}
  .bj-date {{ font-size:11px; color:#607D8B; margin-top:3px; }}
  .bj-gap {{ display:flex; flex-direction:column; justify-content:center; align-items:center;
      padding:0 6px; color:#78909C; font-size:10px; min-width:52px; }}
  .bj-gap .arrow {{ font-size:16px; color:#B0BEC5; }}

  /* ── Scope chips + calendar picker ── */
  .chip-row {{ display:flex; gap:8px; flex-wrap:wrap; align-items:center; }}
  .chip {{ border:1.5px solid {C_SEC}; color:{C_SEC}; background:#fff; border-radius:16px;
           padding:7px 14px; font-size:12.5px; font-weight:700; cursor:pointer; user-select:none;
           transition:background .12s ease, color .12s ease; }}
  .chip:hover {{ background:{C_LBG}; }}
  .chip.active {{ background:{C_SEC}; color:#fff; }}
  .chip.cal {{ border-style:dashed; }}
  .cal-panel {{ background:#fff; border:1.5px solid {C_LBG}; border-radius:10px;
                box-shadow:0 4px 14px rgba(0,0,0,0.12); padding:12px; margin-top:8px;
                max-width:330px; }}
  .cal-head {{ display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; }}
  .cal-head button {{ border:none; background:{C_LBG}; color:{C_PRI}; border-radius:6px;
                      width:34px; height:30px; font-size:15px; cursor:pointer; }}
  .cal-title {{ font-weight:700; color:{C_PRI}; font-size:14px; }}
  .cal-grid {{ display:grid; grid-template-columns:repeat(7,1fr); gap:3px; }}
  .cal-dow {{ text-align:center; font-size:10px; color:#90A4AE; font-weight:700; padding:3px 0; }}
  .cal-day {{ text-align:center; padding:7px 0; font-size:13px; border-radius:6px; color:#B0BEC5; }}
  .cal-day.has {{ color:#263238; font-weight:600; cursor:pointer; background:#F1F8F6; position:relative; }}
  .cal-day.has::after {{ content:''; position:absolute; bottom:2px; left:50%; transform:translateX(-50%);
                         width:4px; height:4px; border-radius:50%; background:{C_GRN}; }}
  .cal-day.has:hover {{ background:{C_LBG}; }}
  .cal-day.sel {{ background:{C_SEC}; color:#fff; }}
  .cal-day.sel::after {{ background:#fff; }}
  .day-nav {{ border:1.5px solid {C_SEC}; background:#fff; color:{C_SEC}; border-radius:6px;
              width:30px; height:28px; font-weight:700; cursor:pointer; }}
  .msum-switch {{ display:flex; gap:6px; padding:0 4px 8px; flex-wrap:wrap; }}
  .msum-switch .chip {{ font-size:11px; padding:5px 11px; }}

  /* ── Date Filter Bar ── */
  .filter-bar {{ position:sticky; top:0; z-index:50; box-shadow:0 2px 6px rgba(0,0,0,0.06);
                 background:#fff; border-bottom:2px solid {C_LBG}; padding:10px 28px;
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
  <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap">
    {f'<img src="{LOGO_URI}" alt="Enicar Pharmaceuticals" style="height:52px;width:auto;background:#fff;border-radius:8px;padding:4px 8px">' if LOGO_URI else ''}
    <div>
      <h1>ENICAR</h1>
      <div class="header-sub">PRODUCTION DASHBOARD &nbsp;|&nbsp; Generated {generated_at}</div>
    </div>
  </div>
</div>
<div class="period-bar">PRODUCTION &nbsp; DASHBOARD &nbsp;&nbsp;|&nbsp;&nbsp; {PERIOD}</div>

<div class="filter-bar" id="filter-bar">
  <div class="chip-row" id="scope-chips"></div>
  <span class="filter-tag" id="filter-tag">{_glance_month_label}</span>
  <span id="day-nav-box" style="display:none">
    <button class="day-nav" onclick="stepDay(-1)" title="Previous day">◀</button>
    <button class="day-nav" onclick="stepDay(1)" title="Next day">▶</button>
  </span>
  <div id="cal-panel" class="cal-panel" style="display:none"></div>
</div>

<div class="container">

<noscript><div class="noscript-note">Interactive date filtering is off (JavaScript disabled
in this viewer) — the numbers below show {_glance_month_label}.</div></noscript>

<!-- ════════════════════════════════════════════════════════════
     SECTION A — AT A GLANCE (director summary layer)
════════════════════════════════════════════════════════════ -->
{director_summary_html()}


<div class="expand-bar"><button id="expand-all-btn" onclick="toggleAllDetails()">▸ Expand all detail sections</button></div>

<!-- ════════════════════════════════════════════════════════════
     SECTION 00 — MONTHLY SUMMARY (RM → Fill → Pack → Disp)
     Hidden by default — shown only when the date-filter is set to
     a "Monthly view: YYYY-MM" option. JS toggles individual blocks.
════════════════════════════════════════════════════════════ -->
<div class="card" id="monthly-summary-card" style="display:none">
  {sec('  ━━&nbsp;&nbsp;MONTHLY &nbsp; SUMMARY &nbsp; (RM &nbsp;→&nbsp; Fill &nbsp;→&nbsp; Pack &nbsp;→&nbsp; Dispatch) &nbsp;━━', C_PRI)}
  <div class="msum-switch" id="msum-switch" style="padding-top:8px"></div>
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
     SECTION 0b — MONTHLY PRODUCTION PLAN (planned vs actual)
════════════════════════════════════════════════════════════ -->
{dispense_schedule_html()}

{plan_section_html()}

{name_conflict_html()}

<!-- ════════════════════════════════════════════════════════════
     SECTION 1 — PRODUCT TYPE BREAKDOWN
════════════════════════════════════════════════════════════ -->
<details class="card">
  <summary>{sec('  ━━&nbsp;&nbsp;PRODUCT &nbsp; TYPE &nbsp; BREAKDOWN &nbsp;━━')}</summary>
  <div class="tbl-wrap">
    <table>
      <tr class="th-row">
        <th>PRODUCT TYPE</th>
        <th id="pt-hdr-fill">UNITS FILLED ({_glance_month_label})</th>
        <th id="pt-hdr-pack">UNITS PACKED ({_glance_month_label})</th>
        <th id="pt-hdr-disp">UNITS DISPATCHED ({_glance_month_label})</th>
      </tr>
      <tbody id="pt-rows">{product_type_rows()}</tbody>
      <tr class="tot-row">
        <td class="td-name">TOTAL</td>
        <td class="td-num" id="pt-total-fill">{n(f_month)}</td>
        <td class="td-num" id="pt-total-pack">{n(p_month)}</td>
        <td class="td-num" id="pt-total-disp">{n(d_month)}</td>
      </tr>
    </table>
  </div>
</details>

<!-- ════════════════════════════════════════════════════════════
     SECTION 2 — FILLING
════════════════════════════════════════════════════════════ -->
<details class="card">
  <summary>{sec('  ━━&nbsp;&nbsp;FILLING &nbsp; PRODUCTION &nbsp;━━')}</summary>
  <div class="tile-row">
    <div class="tile"><div class="tlabel">TOTAL FILLED</div><div class="tvalue" id="f-total" style="color:{C_AMB}">{n(f_month)}</div><div class="tsub">units filled</div></div>
    <div class="tile"><div class="tlabel">FILL RECORDS</div><div class="tvalue" id="f-rec" style="color:{C_AMB}">{f_rec_m}</div><div class="tsub">rows logged</div></div>
    <div class="tile"><div class="tlabel">AVG UNITS / RECORD</div><div class="tvalue" id="f-avg" style="color:{C_AMB}">{n(f_avg_m)}</div><div class="tsub">units per entry</div></div>
    <div class="tile"><div class="tlabel">ACTIVE LINES</div><div class="tvalue" id="f-lines" style="color:{C_AMB}">{f_lines_m}</div><div class="tsub">lines active</div></div>
  </div>
  <div class="tbl-wrap">
    <table>
      <thead id="fill-thead"><tr class="th-row"><th>FILLING LINE</th><th>PRODUCT NAME</th><th>PACK SIZE</th><th>PARTY</th><th>UNITS FILLED ({_glance_month_label})</th></tr></thead>
      <tbody id="fill-line-rows">{grouped_stage_rows(fill_df, 'Qty')}</tbody>
      <tfoot id="fill-tfoot"><tr class="tot-row"><td class="td-name" colspan="4">TOTAL ALL LINES</td><td class="td-num">{n(f_month)}</td></tr></tfoot>
    </table>
  </div>
</details>

<!-- ════════════════════════════════════════════════════════════
     SECTION 2 — PACKING
════════════════════════════════════════════════════════════ -->
<details class="card">
  <summary>{sec('  ━━&nbsp;&nbsp;PACKING &nbsp; PRODUCTION &nbsp;━━')}</summary>
  <div class="tile-row">
    <div class="tile"><div class="tlabel">TOTAL PACKED</div><div class="tvalue" id="p-total" style="color:{C_AMB}">{n(p_month)}</div><div class="tsub">units packed</div></div>
    <div class="tile"><div class="tlabel">FILL → PACK RATIO</div><div class="tvalue" id="p-ratio" style="color:{C_AMB}">{pct(p_ratio_m)}</div><div class="tsub">packed ÷ filled</div></div>
  </div>
  <div class="tbl-wrap">
    <table>
      <thead id="pack-thead"><tr class="th-row"><th>PACKING LINE</th><th>PRODUCT NAME</th><th>PACK SIZE</th><th>PARTY</th><th>UNITS PACKED ({_glance_month_label})</th></tr></thead>
      <tbody id="pack-line-rows">{grouped_stage_rows(pack_df, 'TotalPacked')}</tbody>
      <tfoot id="pack-tfoot"><tr class="tot-row"><td class="td-name" colspan="4">TOTAL ALL LINES</td><td class="td-num">{n(p_month)}</td></tr></tfoot>
    </table>
  </div>
</details>

<!-- ════════════════════════════════════════════════════════════
     SECTION 3 — DISPATCH & BSR STOCK
════════════════════════════════════════════════════════════ -->
<details class="card">
  <summary>{sec('  ━━&nbsp;&nbsp;DISPATCH &nbsp;&amp;&nbsp; BSR &nbsp; STOCK &nbsp;━━', C_ORG)}</summary>
  <div class="tile-row">
    <div class="tile"><div class="tlabel">DISPATCHED</div><div class="tvalue" id="d-total" style="color:{C_ORG}">{n(d_month)}</div><div class="tsub">units dispatched (follows the date filter)</div></div>
    <div class="tile"><div class="tlabel">DISPATCH / FILL</div><div class="tvalue" id="d-ratio" style="color:{C_ORG}">{pct(d_ratio_m)}</div><div class="tsub">dispatched ÷ filled</div></div>
    {tile('PACKED STOCK IN BSR', n(IN_STOCK_UNITS),
          'packed, not yet dispatched — today, all batches; full list in the stock section below', C_SEC)}
    {tile('FILLED, AWAITING PACKING', n(WIP_UNITS),
          'work in progress on the floor — filled but not yet packed (today, all batches)', C_SEC)}
  </div>
  {f'<div style="font-size:11px;color:#90A4AE;padding:0 16px 12px">Note: {AUTO_CLEARED_COUNT} old batch(es) with small leftovers ({n(AUTO_CLEARED_TOTAL)} units, untouched 60+ days after their last dispatch) are auto-cleared from stock as samples/shrinkage.</div>' if AUTO_CLEARED_TOTAL else ''}
</details>

<!-- ════════════════════════════════════════════════════════════
     SECTION 4 — STAFF
════════════════════════════════════════════════════════════ -->
<details class="card">
  <summary>{sec('  ━━&nbsp;&nbsp;STAFF &nbsp;&amp;&nbsp; ATTENDANCE &nbsp;━━')}</summary>
  <div class="tile-row">
    <div class="tile"><div class="tlabel">FEMALE WORKERS PRESENT</div><div class="tvalue" id="s-fem" style="color:{C_AMB}">{s_fem_m:.0f} avg</div><div class="tsub">packing workers</div></div>
    <div class="tile"><div class="tlabel">MALE WORKERS PRESENT</div><div class="tvalue" id="s-male" style="color:{C_AMB}">{s_male_m:.0f} avg</div><div class="tsub">filling & loading</div></div>
  </div>
</details>

<!-- ════════════════════════════════════════════════════════════
     SECTION 6 — PARTY-WISE SALES
════════════════════════════════════════════════════════════ -->
<details class="card">
  <summary>{sec('  ━━&nbsp;&nbsp;PARTY-WISE &nbsp; SALES &nbsp; (Dispatched) &nbsp;━━', C_ORG)}</summary>
  <div class="tbl-wrap">
    <table>
      <thead id="party-thead"><tr class="th-row"><th>PARTY NAME</th><th>DISPATCHED ({_glance_month_label})</th></tr></thead>
      <tbody id="party-rows">{party_mtd_rows()}</tbody>
      <tfoot id="party-tfoot"><tr class="tot-row"><td class="td-name">TOTAL ALL PARTIES</td><td class="td-num">{n(d_month)}</td></tr></tfoot>
    </table>
  </div>
</details>

<!-- ════════════════════════════════════════════════════════════
     SECTION 7 — BATCHES NEEDING ATTENTION (stuck in the pipeline)
════════════════════════════════════════════════════════════ -->
<details class="card" id="stuck-card">
  <summary>{sec(f'  ━━&nbsp;&nbsp;BATCHES &nbsp; NEEDING &nbsp; ATTENTION &nbsp; ({len(STUCK_BATCHES)} &nbsp; stuck) &nbsp;━━', C_AMB)}</summary>
  <div style="font-size:12px;color:#607D8B;padding:8px 16px 0">
    A batch appears here when it has stopped moving: filled but nothing packed for
    <strong>{STUCK_FILL_DAYS}+ days</strong>, or packed but nothing dispatched for
    <strong>{STUCK_PACK_DAYS}+ days</strong> (dispatch normally waits on customer schedules,
    so its threshold is longer). Days count from the batch's last activity at that stage.
  </div>
  <div class="tbl-wrap" style="padding-top:10px">
    <table>
      <thead><tr class="th-row">
        <th>BATCH</th><th>PRODUCT</th><th>CUSTOMER</th><th>WHERE IT IS STUCK</th><th>WAITING</th><th>UNITS WAITING</th>
      </tr></thead>
      <tbody>{stuck_batch_rows()}</tbody>
    </table>
  </div>
</details>

<!-- ════════════════════════════════════════════════════════════
     SECTION 8 — PACKED & IN STOCK (not yet dispatched)
════════════════════════════════════════════════════════════ -->
<details class="card">
  <summary>{sec('  ━━&nbsp;&nbsp;PACKED &nbsp;&amp;&nbsp; IN &nbsp; BSR &nbsp; STOCK &nbsp; (Not &nbsp; Yet &nbsp; Dispatched) &nbsp;━━', C_SEC)}</summary>
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
</details>

</div><!-- /container -->

<div class="footer">
  Generated by Enicar Dashboard Generator &nbsp;|&nbsp; {generated_at}<br>
  Data source: Enicar_Dashboard_Template.xlsx &nbsp;|&nbsp;
  Updates automatically every 15 minutes from the live production sheet &nbsp;|&nbsp;
  <a href="https://raw.githubusercontent.com/EnicarPharmaceuticals/enicar-dashboard/main/archive/enicar_latest.json"
     style="color:#00695C;font-weight:700">⬇ Full data (JSON)</a> — daily snapshots archived since 04 Aug 2026
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

// ── Scope helpers ─────────────────────────────────────
// ONE PERIOD STORY: the page shows exactly one scope at a time — a single
// month (default: the latest month with dispatches), "both months", or a
// single day. Every section, including the At-a-Glance cards, follows it.
const DEFAULT_M = '{_g_mkey}';
const MONTHS_ABBR = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function dataMonths() {{
  const s = new Set();
  [...ENICAR.fill, ...ENICAR.pack, ...ENICAR.disp].forEach(r => {{ if (r.date) s.add(r.date.slice(0,7)); }});
  return [...s].sort().reverse();
}}
function monthName(mkey, full) {{
  const [y, m] = mkey.split('-');
  const names = full
    ? ['','JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER']
    : MONTHS_ABBR.map(x => x.toUpperCase());
  return names[parseInt(m)] + ' ' + y;
}}
// Resolve a filter selection into {{rows-per-log, short label, tag, isDaily}}
function resolveScope(sel) {{
  if (sel === 'all') {{
    const label = dataMonths().slice().reverse().map(m => monthName(m, false)).join(' + ');
    return {{ match: r => true, label: label, tag: 'BOTH MONTHS', isDaily: false }};
  }}
  if (sel.startsWith('month:')) {{
    const mk = sel.slice(6);
    return {{ match: r => r.date && r.date.startsWith(mk), label: monthName(mk, false),
             tag: monthName(mk, true), isDaily: false }};
  }}
  // daily
  const [y, m, d] = sel.split('-');
  return {{ match: r => r.date === sel, label: `${{parseInt(d)}} ${{MONTHS_ABBR[parseInt(m)].toUpperCase()}}`,
           tag: 'DAILY VIEW', isDaily: true }};
}}

// ── Scope state + chips + calendar UI ─────────────────
// The user picks a scope via chips (months / both / monthly summary) or the
// click-a-date calendar. CURRENT_SEL holds the active selection.
let CURRENT_SEL = 'month:' + DEFAULT_M;
let CAL_MONTH   = DEFAULT_M;            // month the calendar grid is showing
let CAL_OPEN    = false;

const ACTIVE_DATES = (() => {{           // days that actually have data
  const s = new Set();
  [...ENICAR.fill, ...ENICAR.pack, ...ENICAR.disp].forEach(r => {{ if (r.date) s.add(r.date); }});
  return s;
}})();

function setScope(v) {{
  CURRENT_SEL = v;
  if (!v.startsWith('month') && !v.startsWith('all') && !v.startsWith('monthly:')) CAL_MONTH = v.slice(0,7);
  applyFilter();
  renderFilterUI();
}}

function stepDay(delta) {{
  // ◀ ▶ through days that have data
  const days = [...ACTIVE_DATES].sort();
  let i = days.indexOf(CURRENT_SEL);
  if (i === -1) {{ setScope(days[days.length-1]); return; }}
  i = Math.min(Math.max(i + delta, 0), days.length - 1);
  setScope(days[i]);
}}

function toggleCal() {{ CAL_OPEN = !CAL_OPEN; renderFilterUI(); }}

function renderFilterUI() {{
  const box = document.getElementById('scope-chips');
  if (!box) return;
  const chips = [];
  dataMonths().forEach(mk => {{
    const [y, mo] = mk.split('-');
    chips.push(`<span class="chip ${{CURRENT_SEL==='month:'+mk?'active':''}}" onclick="setScope('month:${{mk}}')">${{MONTHS_ABBR[parseInt(mo)].toUpperCase()}} ${{y}}</span>`);
  }});
  chips.push(`<span class="chip ${{CURRENT_SEL==='all'?'active':''}}" onclick="setScope('all')">BOTH MONTHS</span>`);
  chips.push(`<span class="chip ${{CURRENT_SEL.startsWith('monthly:')?'active':''}}" onclick="openSummary()">📊 MONTHLY SUMMARY</span>`);
  const isDaily = !CURRENT_SEL.startsWith('month') && CURRENT_SEL !== 'all' && !CURRENT_SEL.startsWith('monthly:');
  chips.push(`<span class="chip cal ${{isDaily?'active':''}}" onclick="toggleCal()">📅 ${{isDaily ? prettyDate(CURRENT_SEL) : 'PICK A DAY'}}</span>`);
  box.innerHTML = chips.join('');
  document.getElementById('day-nav-box').style.display = isDaily ? '' : 'none';
  renderCalendar();
}}

function openSummary() {{
  // Monthly summary of the month currently in scope (falls back to latest)
  const mk = CURRENT_SEL.startsWith('month:') ? CURRENT_SEL.slice(6)
           : (!CURRENT_SEL.startsWith('all') && !CURRENT_SEL.startsWith('monthly:')) ? CURRENT_SEL.slice(0,7)
           : DEFAULT_M;
  const blk = document.getElementById('monthly-block-' + mk);
  setScope('monthly:' + (blk ? mk : DEFAULT_M));
}}

function renderCalendar() {{
  const panel = document.getElementById('cal-panel');
  if (!panel) return;
  if (!CAL_OPEN) {{ panel.style.display = 'none'; return; }}
  panel.style.display = '';
  const [y, m] = CAL_MONTH.split('-').map(Number);
  const first = new Date(y, m-1, 1);
  const daysInMonth = new Date(y, m, 0).getDate();
  const startDow = first.getDay();                     // 0 = Sunday
  let html = `<div class="cal-head">
    <button onclick="calNav(-1)">◀</button>
    <span class="cal-title">${{MONTHS_ABBR[m]}} ${{y}}</span>
    <button onclick="calNav(1)">▶</button>
  </div><div class="cal-grid">`;
  ['S','M','T','W','T','F','S'].forEach(d => html += `<div class="cal-dow">${{d}}</div>`);
  for (let i = 0; i < startDow; i++) html += '<div class="cal-day"></div>';
  for (let d = 1; d <= daysInMonth; d++) {{
    const iso = `${{y}}-${{String(m).padStart(2,'0')}}-${{String(d).padStart(2,'0')}}`;
    const has = ACTIVE_DATES.has(iso);
    const sel = CURRENT_SEL === iso;
    html += `<div class="cal-day ${{has?'has':''}} ${{sel?'sel':''}}" ${{has?`onclick="CAL_OPEN=false; setScope('${{iso}}')"`:''}}>${{d}}</div>`;
  }}
  html += '</div><div style="font-size:10px;color:#90A4AE;margin-top:6px">Green dot = day with production data. Tap a day to view it.</div>';
  panel.innerHTML = html;
}}

function calNav(delta) {{
  let [y, m] = CAL_MONTH.split('-').map(Number);
  m += delta; if (m < 1) {{ m = 12; y--; }} if (m > 12) {{ m = 1; y++; }}
  CAL_MONTH = `${{y}}-${{String(m).padStart(2,'0')}}`;
  renderCalendar();
}}

// ── Main render ───────────────────────────────────────
function applyFilter() {{
  const sel = CURRENT_SEL;

  // Monthly-Summary view: show only the chosen month's block, hide everything else
  const isMonthlyView = sel.startsWith('monthly:');
  const summaryCard = document.getElementById('monthly-summary-card');
  document.querySelectorAll('.monthly-block').forEach(b => b.style.display = 'none');
  if (isMonthlyView) {{
    const m = sel.slice('monthly:'.length);
    summaryCard.style.display = '';
    const blk = document.getElementById('monthly-block-' + m);
    if (blk) blk.style.display = '';
    // Month switcher inside the summary card
    const sw = document.getElementById('msum-switch');
    if (sw) {{
      sw.innerHTML = [...document.querySelectorAll('#monthly-summary-card .monthly-block')].map(b => {{
        const mk = b.id.replace('monthly-block-','');
        const [y, mo] = mk.split('-');
        return `<span class="chip ${{mk===m?'active':''}}" onclick="setScope('monthly:${{mk}}')">${{MONTHS_ABBR[parseInt(mo)].toUpperCase()}} ${{y}}</span>`;
      }}).join('');
    }}
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

  const scope = resolveScope(sel);
  document.getElementById('filter-tag').textContent = scope.tag;
  updateGlance(sel);

  const fill  = ENICAR.fill.filter(scope.match);
  const pack  = ENICAR.pack.filter(scope.match);
  const disp  = ENICAR.disp.filter(scope.match);
  const staff = ENICAR.staff.filter(scope.match);
  const label = scope.label;

  // Update product type breakdown column headers
  document.getElementById('pt-hdr-fill').textContent = `UNITS FILLED (${{label}})`;
  document.getElementById('pt-hdr-pack').textContent = `UNITS PACKED (${{label}})`;
  document.getElementById('pt-hdr-disp').textContent = `UNITS DISPATCHED (${{label}})`;

  renderProductTypes(fill, pack, disp);
  renderFilling(fill, label, scope.isDaily);
  renderPacking(pack, fill, label, scope.isDaily);
  renderDispatch(disp, fill);
  renderStaff(staff, scope.isDaily);
  renderParties(disp, scope.isDaily, label);
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
function renderFilling(fill, label, isDaily) {{
  const tot = fill.reduce((s,r)=>s+(r.qty||0),0);
  const rec = fill.length;
  const lines = new Set(fill.map(r=>r.line).filter(Boolean)).size;
  document.getElementById('f-total').textContent = fmt(tot);
  document.getElementById('f-rec').textContent = rec;
  document.getElementById('f-avg').textContent = rec ? fmt(tot/rec) : '0';
  document.getElementById('f-lines').textContent = lines;

  let rows = ''; let i = 0;
  // Month/both views: grouped by line+product+packSize+party.
  // Daily view: also grouped by BATCH, with the batch's running total across
  // all days — so today's row shows yesterday's progress too.
  const batchCol = isDaily ? '<th>BATCH</th>' : '';
  const totCol   = isDaily ? '<th>BATCH TOTAL (ALL DAYS)</th>' : '';
  const nCols    = isDaily ? 7 : 5;
  document.getElementById('fill-thead').innerHTML = `<tr class="th-row"><th>FILLING LINE</th><th>PRODUCT NAME</th><th>PACK SIZE</th><th>PARTY</th>${{batchCol}}<th>UNITS FILLED (${{label}})</th>${{totCol}}</tr>`;
  document.getElementById('fill-tfoot').innerHTML = `<tr class="tot-row"><td class="td-name" colspan="${{nCols-1-(isDaily?1:0)}}">TOTAL ALL LINES</td><td class="td-num">${{fmt(tot)}}</td>${{isDaily?'<td></td>':''}}</tr>`;
  const byKey = {{}};
  fill.forEach(r => {{
    const ln = r.line||'—', pr = r.product||'—', ps = (r.packSize===null||r.packSize===undefined||r.packSize==='')?'—':String(r.packSize), pa = r.party||'—';
    const bt = isDaily ? (r.batch||'—') : '';
    const k = ln + '|||' + pr + '|||' + ps + '|||' + pa + '|||' + bt;
    if (!byKey[k]) byKey[k] = {{line:ln, product:pr, packSize:ps, party:pa, batch:bt, qty:0}};
    byKey[k].qty += (r.qty||0);
  }});
  Object.values(byKey).sort((a,b) => cmpLine(a.line,b.line)
                                  || a.product.localeCompare(b.product)
                                  || String(a.packSize).localeCompare(String(b.packSize))
                                  || a.party.localeCompare(b.party)).forEach(d => {{
    const bg = i++%2===0 ? '#F1F8F6':'#fff';
    let extra1 = '', extra2 = '';
    if (isDaily) {{
      const j = JOURNEY_BY_KEY[bkeyJS(d.batch)];
      const bTot = j ? j.filled : 0;
      extra1 = `<td class="td-name" style="font-weight:600">${{d.batch}}</td>`;
      extra2 = `<td class="td-num" style="color:#00695C;font-weight:700">${{bTot?fmt(bTot):'—'}}</td>`;
    }}
    rows += `<tr style="background:${{bg}}"><td class="td-name">${{d.line}}</td><td class="td-name">${{d.product}}</td><td class="td-name" style="color:#37474F">${{d.packSize}}</td><td class="td-name" style="color:#546E7A">${{d.party}}</td>${{extra1}}<td class="td-num" style="color:#BF360C;font-weight:700">${{fmt(d.qty)}}</td>${{extra2}}</tr>`;
  }});
  document.getElementById('fill-line-rows').innerHTML = rows || `<tr><td colspan="${{nCols}}" style="text-align:center;color:#90A4AE;padding:12px">No data</td></tr>`;
}}

// ── Packing ───────────────────────────────────────────
function renderPacking(pack, fill, label, isDaily) {{
  const tot = pack.reduce((s,r)=>s+(r.totalPacked||0),0);
  const fTot = fill.reduce((s,r)=>s+(r.qty||0),0);
  document.getElementById('p-total').textContent = fmt(tot);
  document.getElementById('p-ratio').textContent = fTot ? (tot/fTot*100).toFixed(1)+'%' : '—';

  let rows = ''; let i = 0;
  // Month/both views: grouped by line+product+packSize+party.
  // Daily view: also grouped by BATCH with the batch's running packed total.
  const batchCol = isDaily ? '<th>BATCH</th>' : '';
  const totCol   = isDaily ? '<th>BATCH TOTAL (ALL DAYS)</th>' : '';
  const nCols    = isDaily ? 7 : 5;
  document.getElementById('pack-thead').innerHTML = `<tr class="th-row"><th>PACKING LINE</th><th>PRODUCT NAME</th><th>PACK SIZE</th><th>PARTY</th>${{batchCol}}<th>UNITS PACKED (${{label}})</th>${{totCol}}</tr>`;
  document.getElementById('pack-tfoot').innerHTML = `<tr class="tot-row"><td class="td-name" colspan="${{nCols-1-(isDaily?1:0)}}">TOTAL ALL LINES</td><td class="td-num">${{fmt(tot)}}</td>${{isDaily?'<td></td>':''}}</tr>`;
  const byKey = {{}};
  pack.forEach(r => {{
    const ln = r.line||'—', pr = r.product||'—', ps = (r.packSize===null||r.packSize===undefined||r.packSize==='')?'—':String(r.packSize), pa = r.party||'—';
    const bt = isDaily ? (r.batch||'—') : '';
    const k = ln + '|||' + pr + '|||' + ps + '|||' + pa + '|||' + bt;
    if (!byKey[k]) byKey[k] = {{line:ln, product:pr, packSize:ps, party:pa, batch:bt, qty:0}};
    byKey[k].qty += (r.totalPacked||0);
  }});
  Object.values(byKey).sort((a,b) => cmpLine(a.line,b.line)
                                  || a.product.localeCompare(b.product)
                                  || String(a.packSize).localeCompare(String(b.packSize))
                                  || a.party.localeCompare(b.party)).forEach(d => {{
    const bg = i++%2===0 ? '#F1F8F6':'#fff';
    let extra1 = '', extra2 = '';
    if (isDaily) {{
      const j = JOURNEY_BY_KEY[bkeyJS(d.batch)];
      const bTot = j ? j.packed : 0;
      extra1 = `<td class="td-name" style="font-weight:600">${{d.batch}}</td>`;
      extra2 = `<td class="td-num" style="color:#00695C;font-weight:700">${{bTot?fmt(bTot):'—'}}</td>`;
    }}
    rows += `<tr style="background:${{bg}}"><td class="td-name">${{d.line}}</td><td class="td-name">${{d.product}}</td><td class="td-name" style="color:#37474F">${{d.packSize}}</td><td class="td-name" style="color:#546E7A">${{d.party}}</td>${{extra1}}<td class="td-num" style="color:#BF360C;font-weight:700">${{fmt(d.qty)}}</td>${{extra2}}</tr>`;
  }});
  document.getElementById('pack-line-rows').innerHTML = rows || `<tr><td colspan="${{nCols}}" style="text-align:center;color:#90A4AE;padding:12px">No data</td></tr>`;
}}

// ── Dispatch ──────────────────────────────────────────
function renderDispatch(disp, fill) {{
  const tot = disp.reduce((s,r)=>s+(r.qty||0),0);
  const fTot = fill.reduce((s,r)=>s+(r.qty||0),0);
  document.getElementById('d-total').textContent = fmt(tot);
  document.getElementById('d-ratio').textContent = fTot ? (tot/fTot*100).toFixed(1)+'%' : '—';
}}

// ── Staff ─────────────────────────────────────────────
function renderStaff(staff, isDaily) {{
  if (!staff.length) {{
    document.getElementById('s-fem').textContent = '—';
    document.getElementById('s-male').textContent = '—';
    return;
  }}
  const fem  = staff.reduce((s,r)=>s+(r.female||0),0);
  const male = staff.reduce((s,r)=>s+(r.male||0),0);
  if (!isDaily) {{
    document.getElementById('s-fem').textContent  = (fem/staff.length).toFixed(0) + ' avg';
    document.getElementById('s-male').textContent = (male/staff.length).toFixed(0) + ' avg';
  }} else {{
    document.getElementById('s-fem').textContent  = fmt(fem);
    document.getElementById('s-male').textContent = fmt(male);
  }}
}}

// ── Party-wise ────────────────────────────────────────
function renderParties(disp, isDaily, label) {{
  const tot = disp.reduce((s,r)=>s+(r.qty||0),0);
  let rows = ''; let i = 0;
  if (!isDaily) {{
    // Month / both-months: group by party only
    document.getElementById('party-thead').innerHTML = `<tr class="th-row"><th>PARTY NAME</th><th>DISPATCHED (${{label}})</th></tr>`;
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
    document.getElementById('party-thead').innerHTML = `<tr class="th-row"><th>PARTY NAME</th><th>PRODUCT NAME</th><th>DISPATCHED (${{label}})</th></tr>`;
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

// ── At-a-Glance: the cards follow the SAME scope as every detail section ──
const bkeyJS = s => String(s || '').replace(/\\s+/g, '').toUpperCase();
const JOURNEY_BY_KEY = {{}};
(ENICAR.batches || []).forEach(b => JOURNEY_BY_KEY[bkeyJS(b.batch)] = b);
const PIPE_STATUSES = new Set(['Filled & packed (in stock)', 'Filled only']);

function updateGlance(sel) {{
  if (!document.getElementById('glance-disp')) return;
  // A selected DAY scopes the glance to that day's whole month (daily batch
  // counts would be meaningless); month and both-months scopes pass through.
  const gsel = (!sel.startsWith('month:') && sel !== 'all') ? 'month:' + sel.slice(0, 7) : sel;
  const scope = resolveScope(gsel);
  const phrase = gsel === 'all' ? 'across both months' : 'this month';
  const noteLabel = gsel === 'all' ? scope.label + ' (BOTH MONTHS)' : scope.tag;
  const fillR = ENICAR.fill.filter(scope.match), packR = ENICAR.pack.filter(scope.match), dispR = ENICAR.disp.filter(scope.match);

  const fTot = fillR.reduce((s, r) => s + (r.qty || 0), 0);
  const pTot = packR.reduce((s, r) => s + (r.totalPacked || 0), 0);
  const dTot = dispR.reduce((s, r) => s + (r.qty || 0), 0);
  const parties = new Set(dispR.filter(r => r.party && (r.qty || 0) > 0).map(r => r.party));

  // Completed = dispatched this month with a full journey (filled + packed records).
  const dispKeys = new Set(dispR.filter(r => r.batch).map(r => bkeyJS(r.batch)));
  let comp = 0;
  dispKeys.forEach(k => {{
    const b = JOURNEY_BY_KEY[k];
    if (b && b.filled > 0 && b.packed > 0) comp++;
  }});
  // Pipeline = fill/pack activity this month, nothing dispatched yet.
  const actKeys = new Set([...fillR, ...packR].filter(r => r.batch).map(r => bkeyJS(r.batch)));
  let pipe = 0;
  actKeys.forEach(k => {{
    const b = JOURNEY_BY_KEY[k];
    if (b && PIPE_STATUSES.has(b.status)) pipe++;
  }});

  document.getElementById('glance-month-note').textContent = noteLabel;
  document.getElementById('glance-fill').textContent = fmt(fTot);
  document.getElementById('glance-fill-sub').textContent = `filled ${{phrase}}`;
  document.getElementById('glance-pack').textContent = fmt(pTot);
  document.getElementById('glance-pack-sub').textContent = `packed ${{phrase}}`;
  document.getElementById('glance-disp').textContent = fmt(dTot);
  document.getElementById('glance-disp-sub').textContent = `sent to ${{parties.size}} customers ${{phrase}}`;
  document.getElementById('glance-comp').textContent = fmt(comp);
  document.getElementById('glance-comp-sub').textContent = `made, packed & dispatched ${{phrase}}`;
  document.getElementById('glance-pipe').textContent = fmt(pipe);
  document.getElementById('glance-pipe-sub').textContent = `filled or packed ${{phrase}}, awaiting dispatch`;
}}

// ── Expand / collapse all detail sections ────────────
function toggleAllDetails() {{
  const all = [...document.querySelectorAll('details.card')];
  const anyClosed = all.some(d => !d.open);
  all.forEach(d => d.open = anyClosed);
  document.getElementById('expand-all-btn').textContent =
    anyClosed ? '▾ Collapse all detail sections' : '▸ Expand all detail sections';
}}

// ── Batch / product lookup + journey timeline ─────────
const BATCHES = ENICAR.batches || [];
let LOOKUP_HITS = [];   // hits of the current search, referenced by row index

// Every product name a batch appears under, across ALL logs (RM, filling,
// packing, dispatch). The search matches any of them — so a batch filled as
// "Anticid Plus" but packed under a different name is still found either way.
const NAMES_BY_KEY = {{}};
function _addName(batch, name) {{
  if (!batch || !name) return;
  const k = bkeyJS(batch);
  (NAMES_BY_KEY[k] = NAMES_BY_KEY[k] || new Set()).add(String(name).trim());
}}
[...ENICAR.fill, ...ENICAR.pack, ...ENICAR.disp].forEach(r => _addName(r.batch, r.product));
BATCHES.forEach(b => {{ _addName(b.batch, b.product); if (b.rm) _addName(b.batch, b.rm.product); }});
function namesFor(batch) {{ return [...(NAMES_BY_KEY[bkeyJS(batch)] || [])]; }}
function nameConflict(batch) {{
  // distinct after loose normalisation (case/punct) — spelling variants don't count
  const canon = new Set(namesFor(batch).map(s => s.toLowerCase().replace(/[^a-z0-9]/g,'')));
  const list = [...canon].sort((a,b)=>a.length-b.length);
  const distinct = [];
  list.forEach(p => {{ if (!distinct.some(d => p.includes(d) || d.includes(p))) distinct.push(p); }});
  return distinct.length > 1;
}}

const MONTHS_SHORT = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function prettyDate(iso) {{
  if (!iso) return null;
  const [y,m,d] = iso.split('-');
  return `${{parseInt(d)}} ${{MONTHS_SHORT[parseInt(m)]}} ${{y}}`;
}}
function daysBetween(isoA, isoB) {{
  if (!isoA || !isoB) return null;
  return Math.round((new Date(isoB) - new Date(isoA)) / 86400000);
}}

function bjTimeline(b) {{
  // Build the four lifecycle stages. A stage with no record says so explicitly —
  // we never invent a link that isn't in the logs.
  const stages = [
    {{name:'Raw material dispensed', present: !!(b.rm && b.rm.date),
      date: b.rm ? b.rm.date : null, endDate: null,
      qty: b.rm && b.rm.size ? fmt(b.rm.size)+' (batch size)' : null,
      missing:'No record in the RM Dispensing Log'}},
    {{name:'Filling', present: !!b.fillD,
      date: b.fillD ? b.fillD.first : null, endDate: b.fillD ? b.fillD.last : null,
      qty: b.filled ? fmt(b.filled)+' units' : null,
      missing:'No record in the Filling Log'}},
    {{name:'Packing', present: !!b.packD,
      date: b.packD ? b.packD.first : null, endDate: b.packD ? b.packD.last : null,
      qty: b.packed ? fmt(b.packed)+' units' : null,
      missing:'No record in the Packing Log'}},
    {{name:'Dispatch', present: !!b.dispD,
      date: b.dispD ? b.dispD.first : null, endDate: b.dispD ? b.dispD.last : null,
      qty: b.dispatched ? fmt(b.dispatched)+' units' : null,
      missing:'No record in the Dispatch Log'}},
  ];
  let html = '<div class="bj-timeline">';
  let prevDate = null;
  stages.forEach((s, i) => {{
    if (i > 0) {{
      const gap = (s.present && prevDate) ? daysBetween(prevDate, s.date) : null;
      html += `<div class="bj-gap"><span class="arrow">→</span>${{gap!==null ? `<span>${{gap}} day${{gap===1?'':'s'}}</span>` : ''}}</div>`;
    }}
    if (s.present) {{
      const range = (s.endDate && s.endDate !== s.date)
        ? `${{prettyDate(s.date)}} → ${{prettyDate(s.endDate)}}`
        : prettyDate(s.date) || 'date unknown';
      html += `<div class="bj-stage"><div class="bj-name">${{s.name}}</div>
        <div class="bj-qty">${{s.qty || '—'}}</div>
        <div class="bj-date">${{range}}</div></div>`;
      if (s.date) prevDate = s.date;
    }} else {{
      html += `<div class="bj-stage missing"><div class="bj-name">${{s.name}}</div>
        <div style="font-size:12px;margin-top:4px">${{s.missing}}</div></div>`;
    }}
  }});
  html += '</div>';
  const extras = [];
  if (b.stuck) extras.push(`<span style="background:#FBE9E7;color:#BF360C;border-radius:4px;padding:3px 8px;font-weight:700">⏸ Stuck — ${{b.stuck.stage.toLowerCase()}} for ${{b.stuck.days}} days</span>`);
  if (b.rm && b.rm.customer) extras.push(`<span style="color:#607D8B">RM log customer: <strong>${{b.rm.customer}}</strong></span>`);
  extras.push(`<span style="color:#607D8B">Status: ${{b.status}}</span>`);
  return html + `<div style="padding:0 14px 12px;font-size:12px;display:flex;gap:14px;flex-wrap:wrap;align-items:center">${{extras.join('')}}</div>`;
}}

function toggleJourney(idx) {{
  const row = document.getElementById('bj-detail-' + idx);
  if (row) row.style.display = row.style.display === 'none' ? '' : 'none';
}}

function lookupBatch() {{
  const q = (document.getElementById('batch-search').value || '').trim().toLowerCase();
  const out = document.getElementById('batch-search-results');
  if (!q) {{
    out.innerHTML = '<div style="color:#90A4AE;padding:8px 4px;font-size:12px">Type a batch number or product name, then click a result row to see its full journey (raw material → filling → packing → dispatch) with dates and waiting time at each step…</div>';
    return;
  }}
  const hits = BATCHES.filter(b =>
    ((b.batch||'').toLowerCase().includes(q)) ||
    namesFor(b.batch).some(nm => nm.toLowerCase().includes(q))
  );
  if (!hits.length) {{
    out.innerHTML = `<div style="color:#90A4AE;padding:8px 4px">No batches or products match "${{q}}".</div>`;
    return;
  }}
  // Sort: matching product alphabetically, then batch
  hits.sort((a,b) => (a.product||'').localeCompare(b.product||'') || (a.batch||'').localeCompare(b.batch||''));
  LOOKUP_HITS = hits;
  const openFirst = hits.length === 1;   // exactly one match → show its journey immediately
  let rows = '';
  hits.slice(0,80).forEach((b,i) => {{
    const bg = i%2===0 ? '#F1F8F6' : '#fff';
    const stuckChip = b.stuck ? ` <span style="background:#FBE9E7;color:#BF360C;border-radius:3px;padding:1px 5px;font-size:10px;font-weight:700">⏸ ${{b.stuck.days}}d</span>` : '';
    const conflictChip = nameConflict(b.batch)
      ? ` <span style="background:#FFF3E0;color:#E65100;border-radius:3px;padding:1px 5px;font-size:10px;font-weight:700" title="This batch appears under different product names in different logs: ${{namesFor(b.batch).join(' / ')}}">⚠ 2 names</span>`
      : '';
    rows += `<tr style="background:${{bg}};cursor:pointer" onclick="toggleJourney(${{i}})" title="Click to show / hide the batch journey">
      <td class="td-name">${{b.product||'—'}}${{conflictChip}}</td>
      <td class="td-name" style="color:#607D8B">${{b.ptype||'—'}}</td>
      <td class="td-name" style="color:#37474F">${{b.packSize||'—'}}</td>
      <td class="td-name" style="font-weight:600">${{b.batch}}${{stuckChip}}</td>
      <td class="td-num" style="color:#00695C">${{fmt(b.filled)}}</td>
      <td class="td-num" style="color:#BF360C">${{fmt(b.packed)}}</td>
      <td class="td-num" style="color:#E65100">${{fmt(b.dispatched)}}</td>
      <td class="td-name" style="font-size:12px">${{b.status}}</td>
    </tr>
    <tr id="bj-detail-${{i}}" style="display:${{openFirst?'':'none'}}"><td colspan="8" style="background:#FAFDFC;border-top:1px solid #E0F2F1;padding:0">${{bjTimeline(b)}}</td></tr>`;
  }});
  const note = hits.length>80 ? `<div style="color:#90A4AE;font-size:11px;padding:4px">Showing first 80 of ${{hits.length}} matches — refine your search.</div>` : '';

  // Day-by-day activity for everything matched — so "Anticid" shows exactly
  // what was filled/packed/dispatched on the 22nd, the 23rd, and so on.
  const hitKeys = new Set(hits.map(b => bkeyJS(b.batch)));
  const rowMatch = r => (r.batch && hitKeys.has(bkeyJS(r.batch)))
                     || ((r.product || '').toLowerCase().includes(q));
  const byDay = {{}};
  const bump = (date, field, v) => {{
    if (!date || !v) return;
    (byDay[date] = byDay[date] || {{f:0,p:0,d:0}})[field] += v;
  }};
  ENICAR.fill.forEach(r => {{ if (rowMatch(r)) bump(r.date, 'f', r.qty||0); }});
  ENICAR.pack.forEach(r => {{ if (rowMatch(r)) bump(r.date, 'p', r.totalPacked||0); }});
  ENICAR.disp.forEach(r => {{ if (rowMatch(r)) bump(r.date, 'd', r.qty||0); }});
  const days = Object.keys(byDay).sort().reverse();
  let actHtml = '';
  if (days.length) {{
    let tf=0, tp=0, td=0;
    const dayRows = days.map((d,i) => {{
      const e = byDay[d]; tf+=e.f; tp+=e.p; td+=e.d;
      return `<tr style="background:${{i%2===0?'#F1F8F6':'#fff'}}">
        <td class="td-name">${{prettyDate(d)}}</td>
        <td class="td-num" style="color:#00695C;font-weight:600">${{e.f?fmt(e.f):'—'}}</td>
        <td class="td-num" style="color:#BF360C;font-weight:600">${{e.p?fmt(e.p):'—'}}</td>
        <td class="td-num" style="color:#E65100;font-weight:600">${{e.d?fmt(e.d):'—'}}</td></tr>`;
    }}).join('');
    actHtml = `<div style="font-weight:700;color:#004D40;font-size:13px;padding:12px 14px 0">📅 Day-by-day activity for this search</div>
    <div class="tbl-wrap" style="padding-top:6px"><table style="min-width:420px">
      <thead><tr class="th-row"><th>DATE</th><th>FILLED</th><th>PACKED</th><th>DISPATCHED</th></tr></thead>
      <tbody>${{dayRows}}</tbody>
      <tfoot><tr class="tot-row"><td class="td-name">TOTAL</td>
        <td class="td-num">${{fmt(tf)}}</td><td class="td-num">${{fmt(tp)}}</td><td class="td-num">${{fmt(td)}}</td></tr></tfoot>
    </table></div>`;
  }}

  out.innerHTML = `<div class="tbl-wrap"><table>
    <thead><tr class="th-row">
      <th>PRODUCT</th><th>TYPE</th><th>PACK SIZE</th><th>BATCH</th><th>FILLED</th><th>PACKED</th><th>DISPATCHED</th><th>STATUS</th>
    </tr></thead>
    <tbody>${{rows}}</tbody>
  </table></div>${{note}}${{actHtml}}`;
}}

// ── Product-type drill-down (monthly summary) ─────────
function togglePTx(gid) {{
  document.querySelectorAll('.ptx-' + gid).forEach(tr => {{
    tr.style.display = tr.style.display === 'none' ? '' : 'none';
  }});
  const c = document.getElementById('ptc-' + gid);
  if (c) c.textContent = c.textContent.startsWith('▸') ?
    c.textContent.replace('▸', '▾') : c.textContent.replace('▾', '▸');
}}

// ── Jump from a plan batch to its full journey in search ──
function jumpToBatch(b) {{
  const inp = document.getElementById('batch-search');
  if (!inp) return;
  inp.value = b;
  lookupBatch();
  if (inp.scrollIntoView) inp.scrollIntoView({{behavior: 'smooth', block: 'center'}});
}}

// ── Plan card: click a row to verify its RM batches ───
function togglePlan(i) {{
  const d = document.getElementById('plan-d-' + i);
  if (d) d.style.display = (d.style.display === 'none' ? '' : 'none');
}}

// ── Plan card: priority / status filter chips ─────────
function pendFilter(el, mode) {{
  document.querySelectorAll('#pending-card .chip-row .chip').forEach(c => c.classList.remove('active'));
  if (el) el.classList.add('active');
  document.querySelectorAll('#pending-rows tr[data-done]').forEach(tr => {{
    const done = tr.getAttribute('data-done') === '1';
    let show = true;
    if (mode === 'pending') show = !done;
    else if (mode === 'done') show = done;
    tr.style.display = show ? '' : 'none';
  }});
}}
// NOT STARTED is a work-planning view — the Director reads it company by
// company to decide what the store dispenses next, so it sorts alphabetically
// by company instead of by priority (26 Aug 2026). Every other chip keeps the
// original priority/due-date order.
let _planOrigOrder = null;
function _planSortByCompany(on) {{
  const tb = document.getElementById('plan-rows');
  if (!tb) return;
  if (!_planOrigOrder) _planOrigOrder = Array.from(tb.children);
  if (!on) {{ _planOrigOrder.forEach(n => tb.appendChild(n)); return; }}
  const pairs = [];
  for (let i = 0; i < _planOrigOrder.length; i++) {{
    const r = _planOrigOrder[i];
    if (r.nodeType !== 1 || !r.hasAttribute('data-prio')) continue;
    const d = _planOrigOrder[i + 1];
    pairs.push([r, (d && d.id && d.id.indexOf('plan-d-') === 0) ? d : null]);
  }}
  pairs.sort((a, b) => {{
    const ca = a[0].getAttribute('data-company') || '';
    const cb = b[0].getAttribute('data-company') || '';
    if (!ca && cb) return 1;            // blank company sorts last
    if (ca && !cb) return -1;
    return ca.localeCompare(cb);
  }});
  pairs.forEach(([r, d]) => {{ tb.appendChild(r); if (d) tb.appendChild(d); }});
}}

let _planChipP = -1, _planQ = '';
function _planApply() {{
  let shown = 0, total = 0;
  document.querySelectorAll('#plan-rows tr[data-prio]').forEach(tr => {{
    const prio = parseInt(tr.getAttribute('data-prio') || '0');
    const srank = parseInt(tr.getAttribute('data-srank') || '0');
    const isnext = tr.getAttribute('data-next') === '1';
    const isflag = tr.getAttribute('data-flag') === '1';
    const p = _planChipP;
    let show = true;
    if (p >= 11 && p <= 15) show = (srank === p - 10);
    else if (p >= 1) show = (prio === p);
    else if (p === -2) show = (srank === 0);
    else if (p === -3) show = (srank > 0 && srank < 5);
    else if (p === -4) show = isnext;
    else if (p === -5) show = isflag;
    else if (p === -6) show = (tr.getAttribute('data-month') === 'AUG');
    else if (p === -7) show = (tr.getAttribute('data-month') !== 'AUG');
    const det = tr.nextElementSibling;                 // paired detail row
    if (show && _planQ) {{
      // match the visible row text, or the batch numbers inside its detail row
      show = tr.textContent.toLowerCase().includes(_planQ)
          || !!(det && det.id && det.id.startsWith('plan-d-')
                && det.textContent.toLowerCase().includes(_planQ));
    }}
    total++; if (show) shown++;
    tr.style.display = show ? '' : 'none';
    if (det && det.id && det.id.startsWith('plan-d-')) det.style.display = 'none';
  }});
  const c = document.getElementById('plan-search-count');
  if (c) c.textContent = _planQ ? `${{shown}} of ${{total}} plan lines` : '';
}}
function planFilter(chipEl, p) {{
  document.querySelectorAll('#plan-card .chip-row .chip').forEach(c => c.classList.remove('active'));
  if (chipEl) chipEl.classList.add('active');
  _planChipP = p;
  _planSortByCompany(p === -2);        // NOT STARTED → alphabetical by company
  _planApply();
}}
function planSearch(q) {{
  _planQ = (q || '').trim().toLowerCase();
  _planApply();
}}

// ── Init on load ──────────────────────────────────────
lookupBatch();
applyFilter();
renderFilterUI();
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
