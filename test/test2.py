import html
import re
from datetime import datetime
from calendar import month_name, month_abbr
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gl

def get_html_content():
    def date_abbrv(text):
        month_map = {full: abbr for full, abbr in zip(month_name[1:], month_abbr[1:])}
        sorted_months = sorted(month_map.keys(), key=len, reverse=True)
    
        def replace(match):
            word = match.group(0)
            return month_map.get(word, word)
        
        pattern = r'\b(?:' + '|'.join(re.escape(month) for month in sorted_months) + r')\b'
        return re.sub(pattern, replace, text)
    
    FONT_FAMILY = "-apple-system,system-ui,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue','Fira Sans',Ubuntu,Oxygen,'Oxygen Sans',Cantarell,'Droid Sans','Apple Color Emoji','Segoe UI Emoji','Segoe UI Emoji','Segoe UI Symbol','Lucida Grande',Helvetica,Arial,sans-serif"
    COLOR1 = '002445'
    COLOR2 = '0a66c2'
    COLOR3 = '56687a'
    BG_COLOR = 'f5faff'
    BG_COLOR2 = 'f3f2f0'
    
    def html_safe(text=''):
        replacements = {
            '{{br}}': '<br aria-hidden="true">',
            '{{br2}}': '<br aria-hidden="true">' * 2,
            '{{nbsp}}': '&nbsp;',
            '{{nbsp4}}': '&nbsp;' * 4,
            '{{b}}': '<b>',
            '{{/b}}': '</b>',
            '{{i}}': '<i>',
            '{{/i}}': '</i>'
        }
        
        pattern = re.compile('|'.join(map(re.escape, replacements)))
        return pattern.sub(lambda m: replacements[m.group(0)], html.escape(str(text)))
    
    def make_table(width='100%', bg='transparent', border_radius=0):
        return f'<table role="presentation" align="center" valign="middle" border="0" cellspacing="0" cellpadding="0" width="{re.sub(r'\D+', '', str(width))}" height="auto" style="border-collapse:collapse;box-sizing:border-box;margin:0 auto;padding:0;border-radius:{border_radius}px;width:{width};max-width:{width};min-width:{width};background-color:{bg};font-family:{FONT_FAMILY};">'
    
    def make_cell(
        value='',
        el='td',
        parse=True,
        closed=True,
        margin=0,
        padding=0,
        bg='transparent',
        align='left',
        valign='middle',
        outline='none',
        border='none',
        border_radius=0,
        font_size=16,
        font_weight=400,
        font_color=COLOR1,
        font_style='normal',
        line_height=1.5,
        width='auto'
    ):
        if not value:
            parse, closed = False, False
        value = html_safe(value) if parse else value
        ret_val = f'<{el} width="{str(width).replace('px', '')}" align="{"left" if align == "justify" else align}" valign="{valign}" style="width:{width}box-sizing:border-box;margin:{margin};padding:{padding};background-color:{bg};text-align:{align};vertical-align:{valign};outline:{outline};border:{border};border-radius:{border_radius};font-size:{font_size}px;font-weight:{font_weight};color:#{font_color};font-style:{font_style};line-height:{line_height}em;word-wrap:break-word;">{value}'
        ret_val += f"</{el}>" if closed else ''
        return ret_val
        
    def make_img(img='hcclogo', href='hcc.com.ph', alt='HCC LOGO', display='block', width=155, padding=0, border_radius=0):
        anchor = f'<a href="https://{href}" style="color:#{COLOR1};display:{display};text-decoration:none;width:{width}{"" if "%" in str(width) else "px"};margin:0;padding:{padding};outline:none;border:none;" target="_blank">'
        img = f"""<img alt="{alt}" src="https://raw.githubusercontent.com/MasterZeeno/SSR/refs/heads/main/test/{img}.png" style="display:{display};outline:none;text-decoration:none;height:auto;width:100%;max-width:100%;border-radius:{border_radius};" width="100%" height="auto">"""
        
        return anchor + img + '</a>'
    
    def fmt_cell(cell):
        if isinstance(cell, (int, float)):
            return f"{cell:,.0f}"
        return str(cell).strip() if cell is not None else ''
        
    
    wb = load_workbook('NSB-P2 SSR.xlsx', read_only=True, data_only=True)
    ws = [s for s in wb.worksheets if s.sheet_state == 'visible'][-1]
    report_date = ws['Q56'].value
    
    data = []
    for r in range(58, 68):
        row = []
        for c in range(16, 21):
            if c != 17:
                row.append(ws[f"{gl(c)}{r}"].value)
        data.append(row)
    wb.close()
    
    data[0][0] = date_abbrv(report_date)
    table_rows = make_table()
    for i, row in enumerate(data):
        tag, row_html = "td", "<tr>"
        for j, cell in enumerate(row):
            align = "center" if (i == 0 and j == 0) else ("left" if j == 0 else "center")
            bg = 'transparent'
            font_style = 'normal'
            font_weight = 400
            border = f"0.069px solid #{COLOR1}" 
            if i == 0 or (i in (8,9) and j == 3):
                bg = f"#{BG_COLOR}"
                font_weight = 600
            elif i > 0 and j == 0:
                cell = f"{'{{nbsp}}'*2}{cell}" 
                font_style = 'italic'
            
            cell_html = make_cell(value=fmt_cell(cell), el=tag, padding='2px 4px', align=align, bg=bg, font_style=font_style, font_size=15, font_weight=font_weight, border=border)
            row_html += cell_html
        row_html += "</tr>"
        table_rows += row_html
    
    zee_details = make_cell(parse=False, padding='0 8px', value=f"{make_table()}<tbody><tr>{make_cell(value='Jay Ar Adlaon Cimacio, RN', font_weight=600, font_color=COLOR2)}</tr><tr>{make_cell(value='Occupational Health Nurse', font_size=14)}</tr><tr>{make_cell(value='License no.: 0847170', font_color=COLOR3, font_size=12, padding='2px 0 0 0')}</tr></tbody></table>")
    signature = f'<tr>{make_cell(parse=False, padding='0 0 30px 0', value=f"{make_table()}<tbody><tr>{make_cell(parse=False, width='64px', value=f'{make_img(img='cimacio', href='linkedin.com/in/masterzeeno', alt='Jay Ar Cimacio, RN', display='inline-block', width='100%', border_radius='100%')}')}{zee_details}</tr></tbody></table>")}</tr>'
    
    notice = 'This email, including its attachments and thread, is confidential and intended solely for the designated recipient. Any unauthorized disclosure, reproduction, or use of this message or its content is strictly prohibited. If you have received this message in error or have unauthorized access to it, please notify the sender immediately and delete all copies permanently.'
    footer = make_cell(parse=False, padding='24px', border_radius=6, bg=f"#{BG_COLOR2}", value=f"{make_table()}<tbody><tr>{make_cell(parse=False, padding='0 0 24px', align='justify', line_height=1.169, font_size=12, value=notice)}</tr><tr>{make_cell(parse=False, value=make_img())}</tr><tr>{make_cell(padding='8px 0', line_height=1.169, font_size=12, value="{{b}}Hilmarc's Construction Corporation (HCC){{/b}}{{br}}1835 E. Rodriguez Sr. Ave., Immaculate Conception, Quezon City, Philippines")}</tr><tr>{make_cell(padding='8px 0', font_style='italic', line_height=1, font_size=10, value="ISO 9001:2015 Certified | PCAB License No. 3886 AAA{{br}}© 1977–" + str(datetime.now().year) + " Hilmarc's Construction Corporation. All rights reserved.")}</tr></tbody></table>")
    
    return f"""<div dir="ltr" style="margin:0;padding:0;outline:none;border:none;box-sizing:border-box;width:100%;height:auto;background-color:#{BG_COLOR2};font-family:{FONT_FAMILY};font-size:16px;color:#{COLOR1};font-weight:400;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;padding-top:8px;">
    {'\n'.join([f'<div aria-hidden="true" style="height:0;max-height:0;width:0;max-width:0;overflow:hidden;opacity:0">{v}</div>' for v in ('NSB-P2 SSR', '&nbsp;' * 169) ])}{make_table(width='512px', bg='#fff', border_radius=6)}<tbody><tr>{make_cell()}{make_table()}<tbody><tr>{make_cell(padding='24px')}{make_table()}<tbody><tr>
    {make_cell(padding='24px')}<div>{make_table()}<tbody><tr>{make_cell()}{make_img()}</td></tr><tr>{make_cell(padding='16px 0 0', font_size=22, font_weight=600, value='Safety Statistics Report (SSR)')}</tr><tr>
    {make_cell(padding='0 0 30px', font_color=COLOR3, value='Construction of the New Senate Building (Phase II){{br}}Navy Village, Fort Bonifacio, Taguig City, Philippines')}</tr><tr>
    {make_cell(padding='0 0 16px', value='{{b}}Greetings! ✨{{/b}}{{br2}}Please see attached file for the above-mentioned subject.{{br2}}For your convenience, a brief summary is also provided in the table below.')}
    </tr><tr>{make_cell(value=f"{table_rows + '</table>'}", parse=False)}</tr><tr>{make_cell(value='{{br}}Thank you—and as always, {{b}}Safety First! 👊{{/b}}{{br2}}Best regards,{{br2}}')}</tr>{signature}
    </tbody></table></div></td></tr><tr>{footer}</tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></div>""".replace(f'\n{" "*4}', '').replace('\n', '')

