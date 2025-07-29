import os
import sys
from PIL import Image as PILImage
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.exceptions import InvalidFileException

REPORT_TITLE = "SAFETY STATISTICS REPORT"
PROJECT_INFO = {
    "Name": [10, "Construction of the New Senate Building (Phase II)"],
    "Site": [11, "Navy village, Fort Bonifacio, Taguig City"],
    "Code": [12, "PE-01-NSBP2-23"]
}

START_ROW = 15
ADDNL_ROW = 53
BASE_WIDTH_PTS = 69
SHEET_DIMENS = {
 "A": 0.67, "B": 8.14, "C": 7.71, "D": 15, "E": 10.43,
 "F": 12.29, "G": 12.29, "H": 12.29, "I": 12.29, "J": 12.29,
 "K": 12.29, "L": 5.14, "M": 2, "N": 0.67
}

FG = "002445"
FG_LIGHT = "00386C"
BG = "93CBFF"
BG_LIGHT = "E2F1FF"

SOURCE_FILE = "NSB-P2 SSR"
TEMPLATE_FILE = f"{SOURCE_FILE} - TEMPLATE"

IMG_PATH = "hcclogo.png"
NEW_IMG_PATH = f"resized-{IMG_PATH}"

temp_wb = Workbook()
temp_ws = temp_wb.active


def points_to_pixels(points):
    return int(points * 96 / 72)

def to_str(value=None, fallback=''):
    return str(value) if value is not None else fallback

def to_int(value, fallback=0):
    try:
        return int(value)
    except (ValueError, TypeError):
        return fallback

def set_wb(path, read_only=True, data_only=True):
    if not path.lower().endswith('.xlsx'):
        path += '.xlsx'
        
    if not os.path.isfile(path):
        print(f"Error: File not found — {path}")
        sys.exit(1)

    try:
        wb = load_workbook(filename=path, read_only=read_only, data_only=data_only)
        return wb
    except (InvalidFileException, OSError) as e:
        print(f"Error: Unable to load workbook — {e}")
        sys.exit(1)

def borderArray(border=None):
    thin = Side(border_style="thin", color=FG_LIGHT)

    if border == 'all':
        return Border(top=thin, right=thin, bottom=thin, left=thin)
    elif border == 'top':
        return Border(top=thin)
    elif border == 'right':
        return Border(right=thin)
    elif border == 'bottom':
        return Border(bottom=thin)
    elif border == 'left':
        return Border(left=thin)
    elif border == 'top_bottom':
        return Border(top=thin, bottom=thin)
    elif border == 'left_right':
        return Border(left=thin, right=thin)
    else:
        return Border()

def resolve_column_letter(column=None, fallback=None):
    if not isinstance(fallback, str):
        fallback = None

    if not column:
        return fallback or column

    try:
        if isinstance(column, int):
            if 1 <= column <= 255:
                return get_column_letter(column)
        else:
            column = str(column).strip().upper()
            if column.isdigit():
                col_int = int(column)
                if 1 <= col_int <= 255:
                    return get_column_letter(col_int)
    except Exception:
        pass

    return fallback or column

def resolve_column_integer(column=None, fallback=None):
    if not isinstance(fallback, int):
        fallback = None

    if column is None:
        return fallback

    if isinstance(column, int):
        return column if 1 <= column <= 254 else fallback

    column_str = str(column).strip().upper()
    if not column_str:
        return fallback

    try:
        return column_index_from_string(column_str)
    except Exception:
        return fallback


def FORMAT_CELL(
  worksheet=temp_ws,
  start_row=0,
  start_column='B',
  end_row=0,
  end_column=None,
  value=None,
  font_size=12,
  bold=True,
  italic=False,
  horizontal_align="center",
  vertical_align="center",
  border='all',
  foreground=FG,
  background=None
  ):
    start_row = to_int(start_row)
    if start_row < 1:
        return
        
    end_row = to_int(end_row)
    if end_row < start_row:
        end_row = start_row
        
    start_column = resolve_column_letter(start_column)
    end_column = resolve_column_letter(end_column)
        
    if end_column:
        worksheet.merge_cells(
          start_row=start_row,
          start_column=start_column,
          end_row=end_row,
          end_column=end_column
        )
        # worksheet.merge_cells(range_string=f"{start_column}{start_row}:{end_column}{end_row}")
        
    cell = worksheet.cell(start_row, column_index_from_string(start_column))
    
    if value:
        if horizontal_align != "center":
            value = f"  {value}"
        cell.value = value
        cell.font = Font(color=FG, bold=bold, italic=italic, name='Arial', size=font_size)
    
    if background:
        if background == 'light':
            fgColor = BG_LIGHT
        else:
            fgColor = BG
            
        cell.fill = PatternFill(fill_type="solid", fgColor=fgColor)
    
    cell.border = borderArray(border)
    cell.alignment = Alignment(horizontal=horizontal_align, vertical=vertical_align)

wb = set_wb(SOURCE_FILE)
visible_sheets = [sheet for sheet in wb.worksheets if sheet.sheet_state == 'visible']

if not visible_sheets:
    print("No visible sheets found.")
    sys.exit(1)

ws = visible_sheets[-1]

report_date = ws[f"Q{ADDNL_ROW + 3}"].value
ref_code = ws[f"B{ADDNL_ROW + 8}"].value
date_range = ws[f"D{ADDNL_ROW + 10}"].value

DEST_FILE = f"{SOURCE_FILE} as of {report_date}.xlsx"

MANPOWER_LIST = {}

start = False
for r in range(69, ws.max_row + 1):
    value = to_str(ws.cell(r, 3).value).strip().upper()

    if start:
        if 'TOTAL' not in value:
            if value:
                data = []
                for c in range(5, 12):
                    data.append(to_int(ws.cell(r, c).value))
                    
                MANPOWER_LIST[value] = data
        else:
            break

    if 'DATE' in value:
        start = True

wb.close()
del wb

# temp_wb = set_wb(TEMPLATE_FILE, False, False)
# temp_ws = temp_wb.active
# temp_ws.title = f"As of {report_date}"

temp_ws.title = f"As of {report_date}"

# Apply column widths
for col, wid in SHEET_DIMENS.items():
    temp_ws.column_dimensions[col].width = wid
    
# Clear existing images
temp_ws._images.clear()

# Get image height in pixels
IMG_HEIGHT = points_to_pixels(int(BASE_WIDTH_PTS * 0.69))

with PILImage.open(IMG_PATH) as img:
    ASPECT_RATIO = img.height / img.width
    IMG_WIDTH = int(IMG_HEIGHT / ASPECT_RATIO)
    img = img.resize((IMG_WIDTH, IMG_HEIGHT))
    img.save(NEW_IMG_PATH)

# Insert image into cell F1
img = Image(NEW_IMG_PATH)
img.anchor = 'F1'
temp_ws.add_image(img)
    
FORMAT_CELL(
  start_row=5, end_row=6,
  start_column='b', end_column='m',
  value=REPORT_TITLE, font_size=20
)

FORMAT_CELL(
  start_row=7, start_column='b',
  end_column='m', value=f"{ref_code}  ",
  font_size=9, italic=True, background=True,
  horizontal_align='right'
)

PROJECT_INFO["Date Range"] = [9, date_range]
for key, value in PROJECT_INFO.items():
    if not "Date" in key:
        start_column = 'b'
        end_column = 'c'
        value = f"Project {key}"
        border = None
        bold = False
    else:
        start_column = 'd'
        end_column = 'k'
        value = value[1]
        border='bottom'
        bold = True

    FORMAT_CELL(
      start_row=value[0],
      start_column=start_column, end_column=end_column,
      value=value, bold=bold, border=border,
      horizontal_align='left', font_size=10
    )
    
KEYS_LEN = len(MANPOWER_LIST.keys())
for CLASS in ("POWER", "HOURS"):
    ROW_KEY = f"I. MAN{CLASS}"
    if CLASS == "HOURS":
        MULTIPLIER = 8
        ROW_KEY = f"II. MAN{CLASS}"
        START_ROW += 2
    else:
        MULTIPLIER = 1
    
    FORMAT_CELL(
      start_row=START_ROW,
      start_column='b', end_column='n',
      value=ROW_KEY, horizontal_align='left',
      border=None, background = "light"
    )

    START_ROW += 1
    for TYPE in ("REGULAR (8am-5pm)", "OVERTIME (6pm-10pm)"):
        START_ROW += 1
        for SUBTYPE in [TYPE, "DATE", *MANPOWER_LIST.keys(), "TOTAL"]:
            if SUBTYPE in [TYPE, "DATE", "TOTAL"]:
                is_bold = True
                horizontal_align = "center"
                if SUBTYPE == TYPE:
                    background = True
                else:
                    background = "light"
            else:
                is_bold = False
                horizontal_align = "left"
                background = None
                
            FORMAT_CELL(
              start_row=START_ROW,
              start_column='c', end_column='d',
              value=SUBTYPE, bold=is_bold, font_size=10,
              horizontal_align=horizontal_align, background=background
            )
            START_ROW += 1
                    
temp_wb.save(DEST_FILE)
