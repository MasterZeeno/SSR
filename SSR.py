import xlwings as xw
import os
import sys
import re

def format_sheet_name(name: str) -> str:
    # Normalize repeated month range: "july 21-july 27, 2025" → "july 21-27, 2025"
    name = re.sub(r'\b(\w+)\s+(\d+)-\1\s+(\d+),', r'\1 \2-\3,', name, flags=re.IGNORECASE)

    # Capitalize each word (like \u$1)
    name = re.sub(r'\b\w', lambda m: m.group().upper(), name)

    return name

# Path to your Excel file
file_path = r'D:\Scripts\NEW\SSR\NSB-P2 SSR.xlsx'

# Exit if the file doesn't exist
if not os.path.exists(file_path):
    print(f"Error: File does not exist -> {file_path}")
    sys.exit(1)

try:
    with xw.App(visible=False) as app:
        wb = app.books.open(file_path)

        # Filter only visible sheets
        visible_sheets = [sheet for sheet in wb.sheets if sheet.visible]

        # Check for visible sheets
        if not visible_sheets:
            print("Error: No visible sheets found.")
            wb.close()
            sys.exit(1)

        last_visible_sheet = visible_sheets[-1]
        last_visible_sheet_name = format_sheet_name(last_visible_sheet.name)

        # Attempt to read the range
        data = last_visible_sheet.range('P58:T67').value
        wb.close()

except Exception as e:
    print(f"An error occurred: {e}")
    sys.exit(1)

# Remove the second column (index 1) from each row
cleaned_data = [[cell for i, cell in enumerate(row) if i != 1] for row in data]
cleaned_data[0][0] = last_visible_sheet_name

# Format cell content
def format_cell(cell):
    if isinstance(cell, (int, float)):
        return f"{cell:,.0f}"
    return cell if cell is not None else ''

# Generate HTML table rows with alignment classes
table_rows = ""
default_styles = "border:0.016em solid #0061ba !important;padding:0.5em !important"
for i, row in enumerate(cleaned_data):
    tag = "th" if i == 0 else "td"
    row_html = "<tr>"
    for j, cell in enumerate(row):
        align_class = "center" if (i == 0 and j == 0) else ("left" if j == 0 else "center")
        # Determine font-style
        italic_class = ""
        if i > 0 and j == 0:  # Not header row & first column
            italic_class = "font-style:italic"
        cell_html = f'<{tag} align="{align_class}" style="text-align:{align_class};{default_styles};{italic_class}">{format_cell(cell)}</{tag}>'
        row_html += cell_html
    row_html += "</tr>"
    table_rows += row_html

# Full HTML document
html_content = f"""<div style="background:0 0;margin:0;padding:0;border:0 none transparent;outline:0 none transparent;width:100%;height:auto;box-sizing:border-box">
  <table width="100%" cellspacing="0" cellpadding="0" style="font-size:16px !important;max-width:537px !important;min-width:280px !important;width:96.69% !important;box-sizing:border-box">
    <tbody>
      <tr>
        <td align="left">
          <table width="100%" cellspacing="0" cellpadding="0" border="0" style="font-family:Helvetica,system-ui,sans-serif;font-size:clamp(.5em,2.353vw + .059em,1em)!important;line-height:clamp(.4em,6.588vw + -.835em,1.8em)!important;padding:0 clamp(1.125em,3.882vw + .397em,1.95em) 0!important;color:#0061ba!important;text-align:left;word-spacing:normal;letter-spacing:normal;box-sizing:border-box">
            <tbody>
              <tr>
                <td>
                  <div style="margin:clamp(1.125em,3.882vw + .397em,1.95em) auto 0;font-size:0.96em;box-sizing:border-box">
                    <h3>Good day, everyone!</h3>
                    <p>
                      Please see the attached updated <b>Safety Statistics Report (SSR)</b>
                      <br>
                      for the <b>Construction of the New Senate Building Project&nbsp;&ndash;&nbsp;P2.</b>
                    </p>
                    <p>
                      Also, kindly review the brief summary in the table below:
                    </p>
                    <table style="border-collapse:collapse;width:100%;font-family:Helvetica,system-ui,sans-serif;font-size:0.96em;color:#0061ba!important;box-sizing:border-box">
                        {table_rows}
                    </table>
                    <p>
                      <br>
                      Thank you&nbsp;&mdash;&nbsp;<i>and as always,</i><b>Safety First!</b> 👊
                    </p>
                  </div>
                </td>
              </tr>
              <tr>
                <td>
                  <div style="border:0 none transparent;border-top:.032em dashed rgba(0,97,186,.69);width:90%;margin:clamp(1.125em,3.882vw + .397em,1.95em) auto;box-sizing:border-box"></div>
                  <p>Best regards,</p>
                </td>
              </tr>
              <tr>
                <td align="center">
                  <a href="https://www.hcc.com.ph/" style="display:block;text-decoration:none" target="_blank">
                    <img src="https://raw.githubusercontent.com/MasterZeeno/Repository/main/zee-signature.png" alt="Jay Ar Adlaon Cimacio, RN, OHN" width="100%" height="auto" style="display:block;max-width:100%;height:auto;border:none;box-sizing:border-box">
                  </a>
                </td>
              </tr>
            </tbody>
          </table>
        </td>
      </tr>
    </tbody>
  </table>
</div>"""

# Save to HTML file
output_path = "body.html"
with open(output_path, "w", encoding="utf-8") as f:
    f.write(html_content)
