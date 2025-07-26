from openpyxl import load_workbook
import sys
import os

if len(sys.argv) < 2:
    sys.exit(1)

path = sys.argv[1]

if not os.path.exists(path):
    sys.exit(1)

try:
    wb = load_workbook(path, read_only=True, data_only=True)
    visible_sheets = [sheet for sheet in wb.worksheets if sheet.sheet_state == 'visible']
    if not visible_sheets:
        sys.exit(1)
    ws = visible_sheets[-1]
    report_date = ws['Q56'].value
    data = [[cell for i, cell in enumerate(row) if i != 1] for row in ws['P58:T67']]
    data[0][0] = report_date
    if report_date is None:
        sys.exit(1)
except Exception:
    sys.exit(1)
    
# Format cell content
def format_cell(cell):
    if isinstance(cell, (int, float)):
        return f"{cell:,.0f}"
    return cell if cell is not None else ''

# Generate HTML table rows with alignment classes
table_rows = ""
table_attrs = """width="100%" cellspacing="0" cellpadding="0" border="0"""
font_defaults = "font-family:Helvetica,system-ui,sans-serif !important;color:#0061ba !important"
default_styles = "border:0.016em solid #0061ba !important;padding:0.5em !important"
for i, row in enumerate(data):
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
  <table {table_attrs}" style="font-size:16px !important;max-width:537px !important;min-width:280px !important;width:96.69% !important;box-sizing:border-box">
    <tbody>
      <tr>
        <td align="left">
          <table {table_attrs}" style="{font_defaults};font-size:clamp(.5em,2.353vw + .059em,1em) !important;line-height:clamp(.4em,6.588vw + -.835em,1.8em) !important;padding:0 clamp(1.125em,3.882vw + .397em,1.95em) 0 !important;text-align:left;word-spacing:normal;letter-spacing:normal;box-sizing:border-box">
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
                    <table {table_attrs}" style="border-collapse:collapse;width:100%;{font_defaults};font-size:0.96em;box-sizing:border-box">
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

print(report_date)
