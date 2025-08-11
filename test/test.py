from pathlib import Path
from datetime import datetime
from typing import Optional, Any
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

from utils import rslv_dir, is_report_date  # Ensure these are properly defined


def sheet_exists(spreadsheet, name: str):
    for sheet in spreadsheet.worksheets():
        if sheet.title == name:
            return sheet
    return None


def delete_sheet(spreadsheet, name: str):
    sheet = sheet_exists(spreadsheet, name)
    if sheet:
        spreadsheet.del_worksheet(sheet)
        print(f"✅ Deleted sheet: {name}")
    else:
        print(f"❕ No sheet named '{name}' found.")


def resolve_shtname(name: str) -> str:
    return name.replace(",", "").replace(" ", "_")


def make_json_safe(value: Any) -> Any:
    """Convert any unsupported type (like datetime) to JSON-safe formats."""
    if pd.isna(value):  # Handles NaN, NaT, None
        return ""
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value

def df_to_json_safe(df: pd.DataFrame) -> list[list]:
    """Convert DataFrame to a JSON-serializable list of lists."""
    return [[make_json_safe(cell) for cell in row] for row in df.values.tolist()]
    
    
    


# === CONFIG ===
SCRIPT_DIR = Path(__file__).resolve().parent
WB_DIR = rslv_dir("assets/wb", SCRIPT_DIR)
CREDS_DIR = rslv_dir("credentials")
CREDS_JSON = CREDS_DIR / "credentials.json"
SHEET_ID = '1nWiV3K5RFogHipKo_Kxbj9qqh8si2l79WFlal0i4ZaU'

# === Locate latest valid XLSX ===
XLSX_FILE: Optional[Path] = None
for xlsx in sorted(WB_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True):
    if is_report_date(xlsx.stem):
        XLSX_FILE = xlsx
        break

if not XLSX_FILE:
    raise FileNotFoundError("❌ No valid report XLSX file found in assets/wb")

NEW_SHEET_NAME = resolve_shtname(XLSX_FILE.stem)

# # === AUTH ===if
# creds = Credentials.from_service_account_file(str(CREDS_JSON), scopes=SCOPES)
# gc = gspread.authorize(creds)

# === OPEN XLSX and convert ===
# # df = pd.read_excel(XLSX_FILE, sheet_name=0)


fmt = lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) else x
df = pd.read_excel(
    XLSX_FILE,
    skiprows=58,
    usecols="R:T",
    nrows=9,
    converters={
        col: (
            lambda x: f"{x:,.0f}"
            if isinstance(x, (int, float))
            else x
        )
        for col in range(3)
    }
)

# df = pd.read_excel(
    # XLSX_FILE,
    # usecols="R:T",
    # skiprows=58,
    # header=None,
    # nrows=9
# )

print(df)

exit(0)

debug_path = SCRIPT_DIR / "debug_output.xlsx"
df.to_excel(debug_path, index=False)
print(f"Debug file saved to: {debug_path}")

exit(0)

# === OPEN SHEET ===
spreadsheet = gc.open_by_key(SHEET_ID)

# Delete existing sheet if exists
delete_sheet(spreadsheet, NEW_SHEET_NAME)

# Create new sheet
worksheet = spreadsheet.add_worksheet(
    title=NEW_SHEET_NAME,
    rows=str(len(df) + 1),
    cols=str(len(df.columns))
)

# Headers + cleaned rows
values = [df.columns.tolist()] + df_to_json_safe(df)

worksheet.update(range_name="A1", values=values)

# Optional: Delete default "Sheet1" if still exists
delete_sheet(spreadsheet, "Sheet1")