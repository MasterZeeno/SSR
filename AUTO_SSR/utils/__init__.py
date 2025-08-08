# utils/__init__.py

# from .env import set_env
# from .config import ASSETS_DIR, REPORT_DATA
from .smart_title import title
from .excel import extract_data, getws, openwb
from .resolver import resolve_dir, is_report_date, fmt_date

WB_DIR = resolve_dir("wb", "assets")
WB_PATH = None

for excel_file in sorted(
    WB_DIR.glob('*.xlsx'),
    key=lambda f: f.stat().st_mtime,
    reverse=True
):
    if is_report_date(excel_file.stem):
        WB_PATH = excel_file
        break

data = {}

if WB_PATH:
    data = extract_data(WB_PATH)
    

__all__ = [
    "resolve_dir",
    "is_report_date",
    "fmt_date",
    "title",
    "extract_data",
    "getws",
    "openwb",
    "data"
]
    