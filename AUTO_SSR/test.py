# import json
# img_url = "https://raw.githubusercontent.com/MasterZeeno/SSR/refs/heads/main/AUTO_SSR/assets/imgs/"
# data = {
    # "Subject": "Updated: Safety Statistics Report (SSR)",
    # "Headers": {
        # "Reference No.": "",
        # "Report Period": "",
        # "Project Name": "Construction of the New Senate Building (Phase II)",
        # "Project Site": "Navy Village, Fort Bonifacio, Taguig City",
        # "Project Code": "PE-01-NSBP2-23"
    # },
    # "Message Body": [
        # "Greetings! ✨",
        # "You may download the updated Excel file using the button above or from the attached copy.",
        # "For your quick reference, a summary of the safety statistics is provided below:",
        # "Thank you—and as always, Safety First! 👊"
    # ],
    # "Summary": {
        # "Description": [
            # "Loss Time Injury",
            # "Restricted Work Case",
            # "First Aid Treatment Case",
            # "Medical Treatment Case",
            # "Fire Incident Case",
            # "Near Miss Incident",
            # "Property Damage Case",
            # "Highest Manpower",
            # "Cumulative Manhours"
        # ],
        # "Previous": [],
        # "This Period": [],
        # "Present": []
    # },
    # "Disclaimer:": [
        # "This is an automated message. Please do not reply directly to this email.",
        # "This email, including any attachments and previous correspondence in the thread, is confidential and intended solely for the designated recipient(s). If you are not the intended recipient, you are hereby notified that any review, dissemination, distribution, printing, or copying of this message and its contents is strictly prohibited. If you have received this email in error or have unauthorized access to it, please notify the sender immediately and permanently delete all copies from your system.",
        # "The sender and the organization shall not be held liable for any unintended transmission of confidential or privileged information."
    # ],
    # "Bold Texts": [
        # "Greetings!",
        # "Safety First!",
        # "Disclaimer:",
        # "automated message",
        # "confidential",
        # "permanently delete all copies"
    # ],
    # "Colors": {
        # "fg": [
            # "#002445",
            # "#0a66c2",
            # "#60607b"
        # ],
        # "bg": [
            # "#f5faff",
            # "#f3f2f0"
        # ]
    # },
    # "zee": {
        # "img": {
            # "src": img_url + "zee.png",
            # "href": "https://linkedin/in/masterzeeno/",
            # "size": 64,
            # "border-radius": "100%"
        # },
        # "content": [
            # {
                # "value": "Jay Ar Adlaon Cimacio",
                # "color": 1,
                # "font-size": 16
            # },
            # {
                # "value": "Occupational Health Nurse",
                # "color": 0,
                # "font-size": 14
            # },
            # {
                # "value": "License No.: 0847170",
                # "color": 2,
                # "font-size": 12
            # }
        # ]
    # },
    # "hcc": {
        # "img": {
            # "src": img_url + "hcc.png",
            # "href": "https://hcc.com.ph/",
            # "size": 32,
            # "border-radius": 0
        # },
        # "content": [
            # {
                # "value": "Hilmarc's Construction Corporation",
                # "color": 0,
                # "font-size": 16
            # },
            # {
                # "value": "1835 E. Rodriguez Sr. Ave., Immaculate Conception, Quezon City",
                # "color": 2,
                # "font-size": 9
            # }
        # ]
    # },
    # "Footer": [
        # {
            # "value": "ISO 9001:2015 Certified | PCAB License No. 3886 AAA",
            # "color": 2,
            # "font-size": 10
        # },
        # {
            # "value": "© 1977-%year%. All rights reserved.",
            # "color": 2,
            # "font-size": 10
        # }
    # ],
    # "Reported": [
        # "Jul 28-Aug 3,2025"
    # ]
# }

# with open("assets/data.json", "w", encoding="utf-8") as f:
    # json.dump(data, f, indent=4)

from pathlib import Path
from urllib.parse import urlparse, urlunparse, quote, parse_qsl, urlencode
from typing import Optional, Union


def rslv_dir(dirname: Union[str, Path], parentdir: Optional[Union[str, Path]] = None) -> Path:
    base = Path(parentdir) if parentdir else Path.cwd()
    directory = (base / dirname).resolve()
    directory.mkdir(parents=True, exist_ok=True)
    return directory

def rel_to(filepath: Union[str, Path], basepath: Optional[Union[str, Path]] = None) -> str:
    filepath = Path(filepath).resolve()
    basepath = Path(basepath).resolve() if basepath else Path.cwd()

    try:
        return str(filepath.relative_to(basepath))
    except ValueError:
        return str(filepath)

def urlify(url: Union[str, bytes]) -> str:
    parsed = urlparse(url)
    
    # Encode path and query
    encoded_path = quote(parsed.path, safe="")  # encode everything including slashes
    encoded_query = urlencode(parse_qsl(parsed.query), doseq=True)
    
    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        encoded_path,
        parsed.params,
        encoded_query,
        parsed.fragment
    ))
    
# base_url = "https://raw.githubusercontent.com/MasterZeeno/SSR/refs/heads/main/"
# url_yawa = f"{Path(__file__).parent.name}/assets/imgs/zee.png"
# "https://raw.githubusercontent.com/MasterZeeno/SSR/refs/heads/main/AUTO_SSR/assets/wb/August 4-10, 2025.xlsx"
# "https://raw.githubusercontent.com/MasterZeeno/SSR/refs/heads/main/AUTO_SSR%2Fassets%2Fwb%2FAugust 4-10%2C 2025.xlsx"
# print(urlify("AUTO_SSR/assets/imgs"))
# Example usage
# 

SCRIPT_DIR = Path(__file__).resolve().parent
IMGS_DIR, WB_DIR = [
    rslv_dir(f"assets/{v}", SCRIPT_DIR)
    for v in ["imgs", "wb"]
]

data = {
    f"{p.name}_dir": rel_to(p, SCRIPT_DIR.parent)
    for p in [IMGS_DIR, WB_DIR]
}


print(data)



