import re
from pathlib import Path
from datetime import datetime, date
from calendar import month_name, month_abbr
from typing import Optional, Union, Dict

MONTH_MAP: Dict[str, str] = {
    full: abbr for full, abbr in zip(month_name[1:], month_abbr[1:])
}

DATE_REGEX: str = r'\b(' + '|'.join(
    sorted(map(re.escape, MONTH_MAP), key=len, reverse=True)
) + r')\b'

# --- FUNCTIONS WITH TYPING ---

def fmt_date(date_string: str) -> str:
    return re.sub(
        DATE_REGEX,
        lambda m: MONTH_MAP[m.group(0)],
        date_string
    )

def extract_end_date(date_string: str) -> Optional[date]:
    date_string = date_string.strip()
    match = re.search(
        r'(?:\w+\s+\d{1,2}-)?(\w+)\s+(\d{1,2}),?\s*(\d{4})',
        date_string, flags=re.IGNORECASE
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

def is_report_date(date_string: str) -> bool:
    end_date = extract_end_date(date_string)
    return (
        False if end_date is None else
        end_date < datetime.today().date()
    )

def rslv_dir(dirname: Union[str, Path], parentdir: Optional[Union[str, Path]] = None) -> Path:
    base: Path = Path(parentdir) if parentdir else Path.cwd()
    directory: Path = (base / dirname).resolve()
    directory.mkdir(parents=True, exist_ok=True)
    return directory

def rel_to(filepath: Union[str, Path], basepath: Optional[Union[str, Path]] = None) -> str:
    filepath = Path(filepath).resolve()
    basepath = Path(basepath).resolve() if basepath else Path.cwd()
    try:
        return str(filepath.relative_to(basepath))
    except ValueError:
        return str(filepath)

