import xlwings as xw

# Open an existing workbook
wb = xw.Book('NSB-P2 SSR.xlsx')  # Replace with your filename

# Loop through sheets in reverse order
latest_visible_sheet = next(
    (sheet for sheet in reversed(wb.sheets) if sheet.visible),
    None
)

if latest_visible_sheet:
    print(f"Latest visible sheet: {latest_visible_sheet.name}")
else:
    print("No visible sheets found.")