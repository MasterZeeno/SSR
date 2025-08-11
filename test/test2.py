import re
from types import SimpleNamespace
from datetime import datetime
from urllib.parse import quote

DATE_RE = re.compile(
        (r"(\w+)\s+(\d{1,2})\s*-\s*(?:" +
        r"(\w+)\s+)?(\d{1,2}),?\s*(\d{4})"),
        flags=re.IGNORECASE
    )

def extract_dates(date_string):
    if not date_string or not date_string.strip():
        return [None, ""]*2
    
    def parse(m, d, y):
        for fmt in (fmts := ("%B %d %Y", "%b %d %Y")):
            try:
                return SimpleNamespace(
                    date=(dt := datetime.strptime(f"{m} {d} {y}", fmt).date()),
                    string=SimpleNamespace(**{
                        k: quote(d) if "q" in k else d
                        for k, v in zip(
                            ["short", "full", "quoted"],
                            [f"{dt:%b}", [f"{dt:%B}"]*2]
                        )
                        if (d := f"{v} {dt.day}, {dt:%Y}")
                    })
                )
            except ValueError:
                continue
        return None
    
    match = DATE_RE.search(date_string).groups()
    
    if not match:
        return [None, ""]*2
    
    return SimpleNamespace(**{
        key: [
            parse((m if m else match[0]), d, match[-1])
            for m, d in zip(match[::2], match[1::2])
        ]
        for key in ["start", "end"]
    })
        
        
dates = [
    "Aug 4-10, 2025",
    "Jul 28-Aug 3, 2025"
]

new_date = extract_dates(dates[0])

print(new_date)




