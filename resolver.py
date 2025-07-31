import xlwings as xw

# Start a hidden Excel instance
app = xw.App(visible=False)

try:
    wb = app.books.open('NSB-P2 SSR.xlsx')
    shts = [sht for sht in wb.sheets if sht.visible]

    if shts:
        pws = shts[-2]
        ws = shts[-1]
        report_date = ws.range('Q56').value
        for row in range(59,68):
            col = 'T' if 'Manhours' in ws.range(f"P{row}").value else 'S'
            ws.range(f"R{row}").value = pws.range(f"{col}{row}").value
            if row < 67:
                ws.range(f"T{row}").value = max(ws.range(f"R{row}").value, pws.range(f"{col}{row}").value)
    else:
        print("No visible sheets found.")

    wb.save()
    wb.close()

finally:
    app.quit()