import re
from pathlib import Path
from openpyxl import load_workbook
from excel_iterators import getValues
from smart_title import smart_title
from datetime import datetime

def extract_date(date_string):
    date_string = date_string.strip()

    # Try to match either:
    # - "July 28-August 3, 2025"
    # - "August 4-10, 2025"
    match = re.search(
        r'(?:\w+\s+\d{1,2}-)?(\w+)\s+(\d{1,2}),?\s*(\d{4})',
        date_string,
        flags=re.IGNORECASE
    )
    if not match:
        return None

    month, day, year = match.groups()

    for fmt in ("%B %d %Y", "%b %d %Y"):
        try:
            return datetime.strptime(f"{month} {day} {year}", fmt).date()
        except ValueError:
            continue

    return None

def is_report_date(date_obj):
    return date_obj < datetime.today().date()

WB_FOLDER = (Path(__file__).parent / "../SSR WORKBOOKS").resolve()

VALID_WB = [
    file for file in sorted(
        WB_FOLDER.glob('*.xlsx'),
        key=lambda f: f.stat().st_mtime,
        reverse=True
    )
    if (end_date := extract_date(file.stem)) is not None
    and is_report_date(end_date)
]

# === Auto-executed when imported ===

WB_PATH = (Path(__file__).parent / WB_FOLDER / VALID_WB[0]).resolve()

# WB_PATH = (Path(__file__).parent / "../PE-01-NSBP2-23 SSR.xlsx").resolve()

wb = load_workbook(WB_PATH, read_only=True, data_only=True)
ws = [s for s in wb.worksheets if s.sheet_state == "visible"][-1]

alias = lambda s: '' if not s else ''.join(c for c in str(s) if c.isupper())

REPORT, REF = (
    (
        [val.strip(), alias(val)]
        if not isinstance(val, list)
        else [s.strip() for s in val]
    )
    for i, v in enumerate(getValues(
        ws=ws,
        min_row=59, max_row=61,
        min_col=2, xrows=[60]
    ))
    if (o := smart_title(''.join(v), i==0))
    and (val := o if i==0 else o.split(':'))
)

HEADERS = [
    [REF[i], *row_data]
    for i, row in enumerate(getValues(
        ws=ws,
        min_row=63, max_row=66,
        min_col=2, max_col=4,
        xcols=[3], direction='col',
    ))
    if (row_data := [
        smart_title(r).replace(
            'Date Range', 'Report Period'
        )
        for r in row
    ])
]

SUMMARY = getValues(
    ws=ws, xcols=[17],
    min_row=58, max_row=67,
    min_col=16, max_col=20,
    fallback="Description"
)
