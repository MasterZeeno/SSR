import re

def minify(html_text):
    return re.sub(r'>\s+<', '><',
           re.sub(r'\s+', ' ',
           re.sub(r'\n+', '',
           html_text.strip())))

HTML_BODY = minify(f"""
<!DOCTYPE html>
<html lang="en" dir="ltr">
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>{SUBJECT}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
</head>
<body lang="en" dir="ltr" style="background-color:#f3f2f0;width:100%;height:auto;font-family:Helvetica,sans-serif;margin:0;padding:0;border:none;outline:none">'
    <div lang="en" dir="ltr" style="background-color:#f3f2f0;width:100%;height:auto;font-family:Helvetica,sans-serif;margin:0;padding:0;border:none;outline:none">'
        <div style="{CONTAINER_STYLE}">
            <div style="height:0;max-height:0;width:0;overflow:hidden;opacity:0">{"&nbsp;"*169}</div>
            <div style="padding:{FONT_SIZE*0.75}px {FONT_SIZE*1.5}px {FONT_SIZE*1.5}px;">
                {CONTENT_DIV};line-height:0;">
                    <p style="font-size:{FONT_SIZE*1.375}px;{bold()};color:{FG_LITE}">{REPORT[0]} ({REPORT[1]})</p>
                    {HEADERS_HTML}
                </div>
                {br(2)}
                <div aria-label="Download file" aria-role="button" {VALIGN};padding:{FONT_SIZE*0.5}px auto;">
                    <a href="{WB_DOWNLOAD_LINK}"
                        {VALIGN};width:max-content;display:block;text-decoration:none;border-radius:{FONT_SIZE*1.5}px;padding:{FONT_SIZE*0.5}px {FONT_SIZE}px;{bold()};font-size:{FONT_SIZE * 0.875};text-align:center;color:#fff;background-color:{FG_LITE};">
                        Download file
                    </a>
                </div>
                {br(2)}
                {hr(FONT_SIZE*0.375, 69)}
                {CONTENT_DIV};padding-top:{FONT_SIZE*1}px;">
                    <p style="{bold()};font-size:{FONT_SIZE*1.125}px;">{MSGS[0]}</p>
                    <p>{MSGS[1]}</p
                    ><p>{MSGS[2]}</p>
                    {SUMMARY_HTML}
                    <p>{MSGS[3]}</p>
                </div>
                {CONTENT_DIV};">
                    {hr(align='left', color=FG)}
                    <p>Best Regards,</p>
                    <div style="line-height:0;">
                        {make_img("zee")}
                        <div {VALIGN};display:inline-block;line-height:0.23;padding-left:{FONT_SIZE*0.5}px;">
                            <p style="color:{FG_LITE};{bold()};font-size:{FONT_SIZE}px;">
                                {ZEE["name"]}
                            </p>
                            <p style="font-size:{FONT_SIZE*0.875}px;">
                                {ZEE["position"]}
                            </p>
                            <p style="font-size:{FONT_SIZE*0.75}px;color:{FG_VAR};">
                                {ZEE["licenses"]}
                            </p>
                        </div>
                    </div>
                </div>
            </div>
            {br(2)}
            {CONTENT_DIV};border-radius:{BORDER_RADIUS}px;background-color:{BG_DARK};padding:{FONT_SIZE*0.75}px {FONT_SIZE*1.5}px {FONT_SIZE*1.5}px;">
                <p style="color:{FG_VAR};text-align:justify;font-size:{FONT_SIZE*0.75}px;">
                    {bold("Disclaimer:")}
                    {br(2)}
                    This is an
                    {bold("automated message.")}
                    Please do not reply directly to this email.
                    {br(2)}
                    This email, including any attachments and previous correspondence in the thread, is
                    {bold("confidential")} and intended solely for the designated recipient(s).
                    If you are not the intended recipient, you are hereby notified that any review, dissemination, distribution, printing, or copying of this message and its contents is strictly prohibited.
                    If you have received this email in error or have unauthorized access to it, please notify the sender immediately and 
                    {bold("permanently delete all copies")} from your system.
                    {br(2)}
                    The sender and the company shall not be held liable for any unintended transmission of confidential or privileged information.
                    {br(3)}
                </p>
                <div style="color:{FG_VAR};text-align:center;">
                    {hr(FONT_SIZE*0.375, 69)}
                    {br()}
                    {make_img("hcc")}
                    <div {VALIGN};display:inline-block;padding:{FONT_SIZE*0.25}px 0;line-height:0;">
                        <p style="font-size:{FONT_SIZE*1}px;{bold()};color:{FG};">
                            {HCC["name"]}
                        </p>
                        <p style="font-size:{FONT_SIZE*0.5625}px;padding-bottom:{FONT_SIZE*0.375}px;">
                            {HCC["address"]}
                        </p>
                    </div>
                    <p style="font-size:{FONT_SIZE*0.625}px;">
                        {HCC["licenses"]}
                        {br()}
                        {HCC["copyleft"]}
                    </p>
                </div>
            </div>
        </div>
    </div>
</body>
</html>""")