import inspect, os, re
from pathlib import Path
from datetime import datetime
from calendar import month_name, month_abbr

# def xdelim(s, d='()', r=False):
    # if not s or len(d) != 2:
        # return ''
    
    # od, cd = map(re.escape, d)
    # p = f'{od}(.*?){cd}'
    
    # return " ".join(
        # re.findall(p, s) if not r else
        # re.sub(p, '', s).split()
    # ).strip()
    
xdelim = lambda s, d='()', r=False: '' if not s or len(d) != 2 else " ".join((re.sub(p := f'{re.escape(d[0])}(.*?){re.escape(d[1])}', '', s).split() if r else re.findall(p, s))).strip()

MONTH_MAP = {full: abbr for full, abbr in zip(month_name[1:], month_abbr[1:])}
DATE_REGEX = r'\b(' + '|'.join(sorted(map(re.escape, MONTH_MAP), key=len, reverse=True)) + r')\b'
    
PROJECT_DATA = {
    "company": {
        "alias": "HCC",
        "website": "hcc.com.ph",
        "name": "Hilmarc's Construction Corporation",
        "address": "1835 E. Rodriguez Sr. Ave., Immaculate Conception, Quezon City",
        "licenses": "ISO 9001:2015 Certified | PCAB License No. 3886 AAA",
        "copyleft": f"© 1977-{str(datetime.now().year)}. All rights reserved."
    },
    "zee": {
        "name": "Jay Ar Adlaon Cimacio, RN",
        "position": "Occupational Health Nurse",
        "licenses": "License No.: 0847170",
        "website": "facebook.com/MasterZeeno",
        "assets": f"{os.path.basename(os.getcwd())}/assets"
    },
    "email": {
        "from": {"zeenoliev": [1]},
        "to": [
            {"jojofundales": [0, 2]}
        ],
        "cc": [
            {"arch_rbporral": [2]},
            {"glachel.arao": [2]},
            {"rbzden": [2]},
            {"aljonporcalla": [1]},
            {"eduardo111680": [1]}
        ],
        "msgs": [
            "Greetings! ✨",
            "Please see the attached file regarding the subject mentioned above.",
            "For your convenience, a brief summary is also provided in the table below.",
            "Thank you&mdash;and as always, ",
            "Safety First! 👊"
        ],
        "password": "frmoyroohmevbgvb"
    },
    "colors": {
        "fg": "#002445",
        "fg_lite": "#0a66c2",
        "fg_var": "#60607b",
        "bg": "#f5faff",
        "bg_dark": "#f3f2f0"
    },
    "excel_file": os.path.abspath('../PE-01-NSBP2-23 SSR.xlsx')
    "excel_file_sample": Path(__file__).parent / '../PE-01-NSBP2-23 SSR.xlsx'
}
