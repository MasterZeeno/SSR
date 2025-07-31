import xlwings as xw

# Start a hidden Excel instance
app = xw.App(visible=False)

try:
    # Open workbook without displaying Excel
    wb = app.books.open('NSB-P2 SSR.xlsx')

    # Loop through sheets in reverse to find the latest visible one
    latest_visible_sheet = next(
        (sheet for sheet in reversed(wb.sheets) if sheet.visible),
        None
    )

    if latest_visible_sheet:
        print(f"Latest visible sheet: {latest_visible_sheet.name}")
    else:
        print("No visible sheets found.")

    # Optional: Close workbook (don't save)
    wb.close(save_changes=False)

finally:
    # Always quit the Excel app
    app.quit()