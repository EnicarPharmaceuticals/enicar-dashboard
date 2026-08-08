#!/usr/bin/env python3
"""
archive_data.py — permanent daily snapshots of all production data
──────────────────────────────────────────────────────────────────
Management (1 Aug 2026): "we need the data for our future project."
The Google Sheet is a LIVE document — rows get edited, corrected and
eventually cleared, so history only survives if we snapshot it.

Every run (guarded to once per day):
  1. Exports the complete structured dataset — every row of RM Dispensing,
     Filling, Packing, Dispatch, Staff, the current plan tab, and the computed
     per-batch journey — as JSON.
  2. Saves locally:   Enicar Report/_Archive/enicar_YYYY-MM-DD.json.gz
     (plus enicar_latest.json, uncompressed, always current)
  3. Pushes to GitHub: archive/enicar_YYYY-MM-DD.json.gz + archive/enicar_latest.json
     so the future project can consume the data straight from a URL:
     https://raw.githubusercontent.com/EnicarPharmaceuticals/enicar-dashboard/main/archive/enicar_latest.json

Runs from the local Mac's 5-minute refresh loop (check_email_and_refresh.py) —
the Mac has the GitHub token; the cloud workflow cannot commit new paths.
Manual run:  python3 archive_data.py [--force]
"""

import os, sys, json, gzip, re, base64
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

HERE  = os.path.dirname(os.path.abspath(__file__))
ROOT  = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
XLSX  = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')
ARCH  = os.path.join(ROOT, '_Archive')
STAMP = os.path.join(ARCH, '.last_archive')


def bk(b): return re.sub(r'\s+', '', str(b)).upper()


def export_all():
    import pandas as pd

    def load(sheet, header=3):
        df = pd.read_excel(XLSX, sheet_name=sheet, header=header)
        df.columns = [' '.join(str(c).split()) for c in df.columns]
        return df

    def rows(df):
        out = []
        for _, r in df.iterrows():
            rec = {}
            for c, v in r.items():
                if str(c).startswith('Unnamed'):
                    continue
                if pd.isna(v):
                    continue
                if hasattr(v, 'strftime'):
                    v = v.strftime('%Y-%m-%d')
                elif isinstance(v, float) and v == int(v):
                    v = int(v)
                rec[str(c)] = v
            if rec:
                out.append(rec)
        return out

    data = {
        'snapshot_date': date.today().isoformat(),
        'generated_at': datetime.now().isoformat(timespec='seconds'),
        'source': 'Enicar_Dashboard_Template.xlsx (live Google Sheet)',
        'sheets': {},
    }
    # log tabs are matched by canon name, not the ➕ prefix — staff rename tabs
    _LOGS = {'fillinglog': 'Filling Log', 'packinglog': 'Packing Log',
             'dispatchlog': 'Dispatch Log', 'rmdispensinglog': 'RM Dispensing Log',
             'stafflog': 'Staff Log'}
    def _canon_tab(s): return re.sub(r'[^a-z0-9]', '', str(s).lower())
    sheet_names = pd.ExcelFile(XLSX).sheet_names
    for s in sheet_names:
        if _canon_tab(s) in _LOGS:
            data['sheets'][_LOGS[_canon_tab(s)]] = rows(load(s))
        elif 'PLAN' in s.upper() and 'DISPENS' not in s.upper():
            data['sheets'][s] = rows(load(s, header=0))

    # per-batch journey summary (cross-log join, same keying as the dashboard)
    J = {}
    def agg(sheet, bcols, qcol, stage, parse_pack=False):
        df = data['sheets'].get(sheet, [])
        for r in df:
            b = next((r[c] for c in bcols if c in r), None)
            if not b or str(b).strip() in ('', '-'):
                continue
            k = bk(b)
            e = J.setdefault(k, {'batch': str(b).strip(), 'filled': 0, 'packed': 0,
                                 'dispatched': 0, 'rm_dispensed': False})
            if stage == 'rm':
                e['rm_dispensed'] = True
                continue
            if parse_pack:
                q = 0.0
                for c, v in r.items():
                    if any(x in c.lower() for x in ('carton', 'sleeve', 'naked')):
                        try:
                            q += float(v)
                        except Exception:
                            s2 = str(v)
                            if '=' in s2:
                                try: q += float(s2.split('=')[-1].replace(',', '').strip())
                                except Exception: pass
                e[stage] += q
            else:
                try:
                    e[stage] += float(r.get(qcol) or 0)
                except Exception:
                    pass
    agg('RM Dispensing Log', ['BATCH NUMBER'], None, 'rm')
    agg('Filling Log', ['Batch No.'], 'Qty Filled (Units)', 'filled')
    agg('Packing Log', ['Batch No.'], None, 'packed', parse_pack=True)
    agg('Dispatch Log', ['Batch No.'], 'Qty Dispatched (Units)', 'dispatched')
    data['batch_journey'] = sorted(J.values(), key=lambda x: x['batch'])
    data['totals'] = {
        'batches': len(J),
        'filled': sum(e['filled'] for e in J.values()),
        'packed': sum(e['packed'] for e in J.values()),
        'dispatched': sum(e['dispatched'] for e in J.values()),
        'rows': {k: len(v) for k, v in data['sheets'].items()},
    }
    return data


def push_to_repo(path_in_repo, content_bytes, msg):
    sys.path.insert(0, HERE)
    import email_config as cfg
    import urllib.request
    H = {'Authorization': f'token {cfg.GITHUB_TOKEN}',
         'Accept': 'application/vnd.github+json', 'User-Agent': 'enicar-dashboard-bot'}
    url = (f'https://api.github.com/repos/{cfg.GITHUB_USERNAME}/'
           f'{cfg.GITHUB_REPO}/contents/{path_in_repo}')
    sha = None
    try:
        d = json.load(urllib.request.urlopen(urllib.request.Request(url, headers=H), timeout=30))
        sha = d['sha']
    except Exception:
        pass
    body = {'message': msg, 'content': base64.b64encode(content_bytes).decode()}
    if sha:
        body['sha'] = sha
    urllib.request.urlopen(urllib.request.Request(
        url, headers=H, method='PUT', data=json.dumps(body).encode()), timeout=60)


def main():
    force = '--force' in sys.argv
    today = date.today().isoformat()
    os.makedirs(ARCH, exist_ok=True)
    try:
        if not force and open(STAMP).read().strip() == today:
            return 0                      # already archived today
    except Exception:
        pass

    print(f'── Data archive · {today} ──')
    data = export_all()
    js = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    gz = gzip.compress(js.encode('utf-8'))
    daily = os.path.join(ARCH, f'enicar_{today}.json.gz')
    open(daily, 'wb').write(gz)
    open(os.path.join(ARCH, 'enicar_latest.json'), 'w').write(js)
    print(f'  local: {daily} ({len(gz)//1024} KB gz, {len(js)//1024} KB raw) — '
          f'{data["totals"]["batches"]} batches, rows={data["totals"]["rows"]}')

    try:
        push_to_repo(f'archive/enicar_{today}.json.gz', gz, f'Data archive {today}')
        push_to_repo('archive/enicar_latest.json', js.encode('utf-8'), f'Data archive latest ({today})')
        print('  pushed to repo: archive/enicar_latest.json + dated snapshot')
    except Exception as e:
        print(f'  (repo push failed, local archive kept: {e})')
        return 0
    open(STAMP, 'w').write(today)
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        print(f'  (archive unavailable: {e})')
        sys.exit(0)
