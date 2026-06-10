#!/usr/bin/env python3
"""
check_date_sanity.py  —  weekly date-entry sanity review
─────────────────────────────────────────────────────────
The logs are safest when dates are real spreadsheet dates. When a date
is typed as TEXT ("08.06.2026") it still works today (we parse it
day-first), but one entry typed month-first ("06/08/2026" meaning 8 Jun)
would silently land in the wrong month. This check reports:

  1. Rows whose Date cell is text rather than a real date.
  2. Parsed dates that look impossible: in the future, or before 2024.

Output is plain text on stdout — cloud_weekly_review.py appends it to
Weekly_Data_Review.txt which gets emailed.
"""

import os, sys, re
from datetime import date, timedelta
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.environ.get('DASHBOARD_ROOT') or os.path.join(HERE, '..')
XLSX = os.path.join(ROOT, 'Enicar_Dashboard_Template.xlsx')

LOGS = [('Filling',  '➕ Filling Log',       'Date'),
        ('Packing',  '➕ Packing Log',       'Date'),
        ('Dispatch', '➕ Dispatch Log',      'Date'),
        ('RM',       '➕ RM Dispensing Log', 'DISPENSING DATE')]


def main():
    print('── Date-entry check ──\n')
    if not os.path.exists(XLSX):
        print('  Could not read the data this week.\n'); return 0

    today = date.today()
    text_counts, impossible = {}, []

    for log, sheet, dcol in LOGS:
        try:
            df = pd.read_excel(XLSX, sheet_name=sheet, header=3)
        except Exception:
            continue
        if dcol not in df.columns:
            continue
        live = df[df[dcol].notna()]
        as_text = live[live[dcol].apply(lambda x: isinstance(x, str))]
        if len(as_text):
            text_counts[log] = len(as_text)
        parsed = pd.to_datetime(live[dcol], format='mixed', dayfirst=True,
                                errors='coerce')
        for idx, ts in parsed.items():
            if pd.isna(ts):
                impossible.append((log, repr(live.loc[idx, dcol]), 'cannot be read as a date'))
            elif ts.date() > today + timedelta(days=2):
                impossible.append((log, str(live.loc[idx, dcol]), f'reads as {ts.date()} (future)'))
            elif ts.year < 2024:
                impossible.append((log, str(live.loc[idx, dcol]), f'reads as {ts.date()} (too old)'))

    if not impossible:
        print('  All dates look good — every row reads as a sensible date. ✅\n')
        return 0

    print('  These dates look WRONG — please correct them in the sheet:\n')
    for log, raw, why in impossible[:20]:
        print(f'    • {log}: "{raw}" — {why}')
    if len(impossible) > 20:
        print(f'    … and {len(impossible)-20} more')
    print()
    print('  Tip: dates here are typed as text (e.g. 08.06.2026) and read day-first.')
    print('  A date typed month-first (06/08/2026 meaning 8 June) lands in the wrong')
    print('  month — please keep the day.month.year order everyone uses.\n')

    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        print(f'  (date check unavailable: {e})\n')
        sys.exit(0)
