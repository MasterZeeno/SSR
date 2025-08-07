import os
import xlwings as xw

WB_NAME = 'PE-01-NSBP2-23 SSR'
WB_FOLDER = f"{WB_NAME.split()[1]} WORKBOOKS"

# Start a hidden Excel instance
app = xw.App(visible=False)

try:
    wb = app.books.open(f"{WB_NAME}.xlsx")
    shts = [sht for sht in wb.sheets if sht.visible]

    if len(shts) > 2:
        ws, pws = shts[-1], shts[-2]
        report_date = str(ws.range('Q56').value).strip()
        for row in range(59,68):
            col = 'T' if 'Manhours' in ws.range(f"P{row}").value else 'S'
            ws.range(f"R{row}").value = pws.range(f"{col}{row}").value
            if row < 67:
                ws.range(f"T{row}").value = max(ws.range(f"R{row}").value, pws.range(f"{col}{row}").value, pws.range(f"T{row}").value)
        wb.save()

        new_wb = xw.Book(visible=False)
        before = new_wb.sheets[0]
        ws.api.Copy(Before=before.api)
        before.delete()
        os.makedirs(WB_FOLDER, exist_ok=True)
        new_wb.save(f"{WB_FOLDER}/{report_date}.xlsx")
        new_wb.close()
    else:
        print("No visible sheets found.")

    wb.close()

finally:
    app.quit()