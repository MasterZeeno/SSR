import os, re
import xlwings as xw
from calendar import month_name, month_abbr

WB_NAME = 'PE-01-NSBP2-23 SSR'
MONTH_MAP = {full: abbr for full, abbr in zip(month_name[1:], month_abbr[1:])}
DATE_REGEX = r'\b(' + '|'.join(sorted(map(re.escape, MONTH_MAP), key=len, reverse=True)) + r')\b'

def fmt_date(text):
    text = str(text).strip()
    return re.sub(
            DATE_REGEX,
            lambda m: MONTH_MAP[m.group(0)],
            text
        ) if text else ''

# Start a hidden Excel instance
app = xw.App(visible=False)

try:
    wb = app.books.open(f"{WB_NAME}.xlsx")
    shts = [sht for sht in wb.sheets if sht.visible]

    if len(shts) > 2:
        ws, pws = shts[-1], shts[-2]
        report_date = fmt_date(ws.range('Q56').value)
        for row in range(59,68):
            col = 'T' if 'Manhours' in ws.range(f"P{row}").value else 'S'
            ws.range(f"R{row}").value = pws.range(f"{col}{row}").value
            if row < 67:
                ws.range(f"T{row}").value = max(ws.range(f"R{row}").value, pws.range(f"{col}{row}").value, pws.range(f"T{row}").value)
        wb.save()

        new_wb = xw.Book(visible=False)
        ws.api.Copy(Before=new_wb.sheets[0].api)
        # new_wb.sheets[0].delete()
        os.makedirs(WB_NAME.split()[1], exist_ok=True)
        new_wb.save(f"{WB_NAME.split()[1]}/{WB_NAME.split()[1]} - {report_date}.xlsx")
        new_wb.close()
    else:
        print("No visible sheets found.")

    wb.close()

finally:
    app.quit()