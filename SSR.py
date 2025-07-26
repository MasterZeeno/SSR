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
    data = []
    for r in range(58, 68):
        row = []
        for c in range(16, 21):
            if c != 17:
                row.append(ws.cell(r,c).value)
        data.append(row)
    data[0][0] = f"As of {report_date}"
    wb.close()
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
font_defaults = "font-family:Helvetica,system-ui,sans-serif;color:#0061ba;text-align:left;word-spacing:normal;letter-spacing:normal"
default_styles = "border:0.0123em solid #0061ba;padding:0.5em"
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
  <table {table_attrs}" style="font-size:16px;max-width:532px;min-width:300px;width:96.69%;box-sizing:border-box">
    <tbody>
      <tr>
        <td align="left">
          <table {table_attrs}" style="{font_defaults};font-size:clamp(.5em,2.353vw + .059em,1em);line-height:clamp(.4em,6.588vw + -.835em,1.8em);padding:0 clamp(1.125em,3.882vw + .397em,1.95em) 0;box-sizing:border-box">
            <tbody>
              <tr>
                <td>
                  <div style="margin:clamp(1.125em,3.882vw + .397em,1.95em) auto 0;font-size:0.96em;box-sizing:border-box">
                    <h3>Good day, everyone!</h3>
                    <p>Please find the attached updated <b>Safety Statistics Report (SSR)</b>
                      for Project Code: <b>PE-01-NSBP2-23&nbsp;&ndash;&nbsp;Construction of the New Senate Building (Phase II).</b>
                    </p>
                    <p>
                      Alternatively, for your convenience, a brief summary is provided in the table below:
                    </p>
                    <table {table_attrs}" style="border-collapse:collapse;width:100%;{font_defaults};font-size:0.9em;box-sizing:border-box">
                        {table_rows}
                    </table>
                    <p>
                      <br>
                      Thank you, and as always&mdash;<b>Safety First!</b>&nbsp;👊
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
