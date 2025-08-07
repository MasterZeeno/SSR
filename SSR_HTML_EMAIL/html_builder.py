import html, os, re, requests, subprocess
from datetime import datetime
from urllib.parse import urlparse, urlunparse, quote, parse_qsl, urlencode
from excel_data_extractor import REPORT, HEADERS, SUMMARY

def minify(html_text):
    html_text = re.sub(r'>\s+<', '><', html_text)         # Remove whitespace between tags
    html_text = re.sub(r'\s{2,}', ' ', html_text)         # Collapse multiple spaces
    html_text = re.sub(r'\n+', '', html_text)             # Remove newlines
    return html_text.strip().replace('; ', ';')

def urlify(raw_url):
    parsed = urlparse(html.unescape(raw_url))
    path = quote(parsed.path, safe="/-',")
    query = urlencode(parse_qsl(parsed.query), doseq=True)

    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        path,
        parsed.params,
        query,
        parsed.fragment
    ))

REPO_OWNER, REPO_NAME = (
    (m := re.search(r'[:/]([^/]+)/([^/]+?)(?:\.git)?$', 
     (subprocess.run(['git', 'config', '--get', 'remote.origin.url'],
        capture_output=True, text=True)).stdout.strip()))
    and (m.group(1), m.group(2)) or ("MasterZeeno", REPORT[1] or 'SSR')
)

ZEE = (lambda r: dict(zip(['name', 'position', 'licenses', 'link'],
    [*r[:1], *r[1].split('|'), r[2]])))(
        [res.json().get(k) for k in ['name', 'bio', 'blog']]
        if (res := requests.get(urlify(f'https://api.github.com/users/{REPO_OWNER}'))).ok
        else [
            "Jay Ar Adlaon Cimacio, RN",
            "Occupational Health Nurse | License No.: 0847170",
            f"https://facebook.com/{REPO_OWNER}"
        ]
    )

HCC = {
    "link": "https://hcc.com.ph",
    "name": "Hilmarc&apos;s Construction Corporation",
    "address": "1835 E. Rodriguez Sr. Ave., Immaculate Conception, Quezon City",
    "licenses": "ISO 9001:2015 Certified | PCAB License No. 3886 AAA",
    "copyleft": f"© 1977-{datetime.now().year}. All rights reserved."
}

MSGS = [
    "Greetings! ✨",
    "Please find attached the updated <i>Excel file</i> regarding the subject mentioned above.",
    "For your quick reference, a brief summary is provided in the table below:",
    "Thank you&mdash;and as always, <b>Safety First! 👊</b>"
]

SUBJECT = f"Submission of {REPORT[0]} ({REPORT[1]}) — {HEADERS[1][1]} for {HEADERS[0][-1]}: {HEADERS[1][-1]}"

LANG_DIR = 'lang="en" dir="ltr"'
FONT_FAMILY = "Helvetica, Arial, sans-serif"
FONT_SIZE = 16
FONT_WEIGHT_NORMAL = 300
FONT_WEIGHT_BOLD = 600
LINE_HEIGHT = 1.5
MAX_WIDTH = 512
BORDER_RADIUS = MAX_WIDTH * 0.01171875

FG, FG_LITE, FG_VAR = "#002445", "#0a66c2", "#60607b"
BG, BG_DARK = "#f5faff", "#f3f2f0"

VALIGN = 'valign="middle" style="vertical-align:middle'
WRAPPER = f"""{LANG_DIR} style="background-color:{BG_DARK};
    margin:0;padding:0;border:none;outline:none;width:100%;
    height:auto;font-family:{FONT_FAMILY};">"""
CONTENT_DIV = f"""<div style="dir:ltr;text-align:left;
    word-wrap:break-word;color:{FG};font-size:{FONT_SIZE}px;
    line-height:{LINE_HEIGHT}"""
CONTAINER_STYLE = f"""border-radius:{BORDER_RADIUS}px;font-family:{FONT_FAMILY};
    width:{MAX_WIDTH}px;max-width:{MAX_WIDTH}px;height:auto;background-color:#fff;
    margin:8px auto;padding:{FONT_SIZE * 1.5}px;"""
IMG_DEFAULT_STYLES = f"""{VALIGN};color:{FG_VAR};height:auto;font-size:{FONT_SIZE/3}px;
    text-decoration:none;outline:none;border:none;"""
SRC_IMG_LINK = f"""https://raw.githubusercontent.com/{REPO_OWNER}/{REPO_NAME}/refs/heads/main/{os.path.basename(os.getcwd())}/assets"""

def make_img(src):
    src_obj = HCC if src == 'hcc' else ZEE
    href, filename = src_obj["link"], src_obj["name"]
    size = 64 if src == 'zee' else 32
    default_styles = IMG_DEFAULT_STYLES + ("border-radius:100%;" if src == 'zee' else '')
    src_link = urlify(f"{SRC_IMG_LINK}/{filename}.png")
        
    return f"""<a href="{urlify(href)}" width="{size}"
        {default_styles};display:inline-block;
        width:{size}px;max-width:{size}px;">
        <img src="{src_link}" alt="{filename}" 
            {default_styles};width:100%;max-width:100%;"></a>"""
        
def hr(size=FONT_SIZE*0.75, count=3, align="center", opacity=0.69, color=FG_VAR):
    return f"""<p role="separator" {VALIGN};color:{color};font-size:{size}px;
        opacity:{opacity};text-align:{align};">{'&mdash;' * count}</p>"""
        
def br(count=1):
    return '<br aria-hidden="true">' * count

def bold(text=None):
    prop = f"font-weight:{FONT_WEIGHT_BOLD}"
    return f'<span style="{prop};">{text}</span>' if text else prop

HEADERS_HTML = ''.join(
    f'<div {VALIGN};display:inline-block;' +
    f'font-size:{FONT_SIZE * 0.875}px;color:' +
    (
        FG_VAR if x==0 else
        f'{FG};font-weight:{FONT_WEIGHT_BOLD}'
    ) + ';">' + ''.join(
        f"""<p style="line-height:0.5;">
        {
            f"{item}{'&nbsp;'*2}" if x==0
            else f":{'&nbsp;'*4}{item}"
        }
        </p>{br() if y==1 else ""}"""
        for y, item in enumerate(row)
    ) + '</div>'
    for x, row in enumerate(HEADERS)
)

SUMMARY_HTML = (
    f'<div style="padding:{FONT_SIZE*0.75}px 0;"><table width="100%" border="0" cellspacing="0" cellpadding="0" ' +
    f'{VALIGN};border-collapse:collapse;font-size:{FONT_SIZE*0.8125}px;color:{FG};"><tbody>' +
    ''.join(
        '<tr>' + ''.join(
            f'<{"th" if i<1 else "td"} {";".join([
                f'{VALIGN};padding:{FONT_SIZE*0.1875}px {FONT_SIZE*0.375}px',
                f'border:{FONT_SIZE*0.0105625}px solid {FG}',
                f'text-align:{"center" if i<1 or j>0 else "left"}',
                f"background-color:{BG}" if i<1 or (i>7 and j==3) else '',
                "font-style:italic" if i>0 and j<1 else '', bold() if i>7 and j==3 else ''
            ])}">{("&nbsp;"*3 + cell if i>0 and j<1 else cell).replace('Teatment', 'Treatment')}</{"th" if i<1 else "td"}>'
            for j, cell in enumerate(row)
        ) + '</tr>'
        for i, row in enumerate(SUMMARY)
    ) + '</tbody></table></div>'
)

HTML_BODY = minify(f"""
<!DOCTYPE html>
<html {LANG_DIR}>
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>{SUBJECT}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
</head>
<body {WRAPPER}
    <div {WRAPPER}
        <div style="{CONTAINER_STYLE}">
            <div style="padding:{FONT_SIZE*0.75}px {FONT_SIZE*1.5}px {FONT_SIZE*1.5}px;">
                {CONTENT_DIV};line-height:0;">
                    <p style="font-size:{FONT_SIZE*1.375}px;{bold()};color:{FG_LITE}">{REPORT[0]} ({REPORT[1]})</p>
                    {HEADERS_HTML}
                </div>{br(2)}{hr(FONT_SIZE*0.375, 69)}
                {CONTENT_DIV};padding-top:{FONT_SIZE*1}px;">
                    <p style="{bold()};font-size:{FONT_SIZE*1.125}px;">{MSGS[0]}</p>
                    <p>{MSGS[1]}</p><p>{MSGS[2]}</p>
                    {SUMMARY_HTML}
                    <p>{MSGS[3]}</p>
                </div>
                {CONTENT_DIV};">
                    {hr(align='left', color=FG)}
                    <p>Best Regards,</p>
                    <div style="line-height:0;">
                        {make_img("zee")}
                        <div {VALIGN};display:inline-block;line-height:0.23;padding-left:{FONT_SIZE*0.5}px;">
                            <p style="color:{FG_LITE};{bold()};font-size:{FONT_SIZE}px;">{ZEE["name"]}</p>
                            <p style="font-size:{FONT_SIZE*0.875}px;">{ZEE["position"]}</p>
                            <p style="font-size:{FONT_SIZE*0.75}px;color:{FG_VAR};">{ZEE["licenses"]}</p>
                        </div>
                    </div>
                </div>
            </div>
            {br(2)}
            {CONTENT_DIV};border-radius:{BORDER_RADIUS}px;background-color:{BG_DARK};padding:{FONT_SIZE*0.75}px {FONT_SIZE*1.5}px {FONT_SIZE*1.5}px;">
                <p style="color:{FG_VAR};text-align:justify;font-size:{FONT_SIZE*0.75}px;">
                    {bold("Disclaimer:")}{br(2)}
                    This is an {bold("automated message.")}
                    Please do not reply directly to this email.{br(2)}
                    This email, including any attachments and previous correspondence in the thread, is {bold("confidential")} and intended solely for the designated recipient(s).
                    If you are not the intended recipient, you are hereby notified that any review, dissemination, distribution, printing, or copying of this message and its contents is strictly prohibited.
                    If you have received this email in error or have unauthorized access to it, please notify the sender immediately and {bold("permanently delete all copies")} from your system.{br(2)}
                    The sender and the organization shall not be held liable for any unintended transmission of confidential or privileged information.{br(3)}
                </p>
                <div style="color:{FG_VAR};text-align:center;">
                    {hr(FONT_SIZE*0.375, 69)}{br()}{make_img("hcc")}
                    <div {VALIGN};display:inline-block;padding:{FONT_SIZE*0.25}px 0;line-height:0;">
                        <p style="font-size:{FONT_SIZE*1}px;{bold()};color:{FG};">{HCC["name"]}</p>
                        <p style="font-size:{FONT_SIZE*0.5625}px;padding-bottom:{FONT_SIZE*0.375}px;">{HCC["address"]}</p>
                    </div>
                    <p style="font-size:{FONT_SIZE*0.625}px;">{HCC["licenses"]}{br()}{HCC["copyleft"]}</p>
                </div>
            </div>
        </div>
    </div>
</body>
</html>""")

with open("index.html", "w", encoding="utf-8") as f:
    f.write(HTML_BODY)








