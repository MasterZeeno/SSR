import re
from pathlib import Path
from calendar import month_name, month_abbr
from datetime import datetime

MONTH_MAP = {full: abbr for full, abbr in zip(month_name[1:], month_abbr[1:])}
DATE_REGEX = r'\b(' + '|'.join(sorted(map(re.escape, MONTH_MAP), key=len, reverse=True)) + r')\b'
    
def fmt_date(date_string):
    return re.sub(DATE_REGEX, lambda m: MONTH_MAP[m.group(0)], date_string)

def extract_end_date(date_string):
    date_string = date_string.strip()
    
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

def is_report_date(date_string):
    end_date = extract_end_date(date_string)
    return (
        False if end_date is None else
        end_date < datetime.today().date()
    )

def resolve_dir(dirname, parentdir=None) -> Path:
    if parentdir is None:
        parentdir = Path.cwd()
    elif isinstance(parentdir, str):
        parentdir = Path(parentdir)
  
    directory = (parentdir / dirname).resolve()
    directory.mkdir(parents=True, exist_ok=True)
    return directory
    



