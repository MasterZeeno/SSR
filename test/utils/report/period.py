import re
from types import SimpleNamespace
from datetime import datetime, date
from calendar import month_name, month_abbr

class DatePeriod(SimpleNamespace):
    def __init__(self, date_string):
        **{
            k: {
                
            }
            for k, d in zip(
                ["start", "end"],
                self._extract(date_string)
            )
        }
        
        super().__init__(
            period=date_string,
            **{
                k: d
                for k, d in zip(
                    [(v, f"{v}_str")for v in ["start", "end"]],
                    [start_date, end_date]
                ) if d
            },
            **{
                f"{k}_str": d.strftime("%b %d, %Y")
                for k, d in zip(
                    ["start", "end"],
                    [start_date, end_date]
                ) if d
            }
        )
    
 
    def _extract(self, date_string):
        date_string = str(date_string).replace("None", "").strip()
        
        if not date_string:
            return [None]*4
            
        month_map = {
            full: abbr for full, abbr
            in zip(month_name[1:], month_abbr[1:])
        }
        
        short_date_regex = r'\b(' + '|'.join(
            sorted(map(re.escape, month_map), key=len, reverse=True)
        ) + r')\b'
        
        date_string = re.sub(
            short_date_regex,
            lambda m: month_map[m.group(0)],
            date_string
        )
        
        match = re.compile(
            r"(\w+)\s+(\d{1,2})\s*-\s*(?:(\w+)\s+)?(\d{1,2}),?\s*(\d{4})",
            flags=re.IGNORECASE
        ).search(date_string)
        
        if not match:
            return [None]*4
    
        start_month, start_day, end_month, end_day, year = match.groups()
    
        if not end_month:
            end_month = start_month
    
        def parse(m, d, y):
            for fmt in ("%B %d %Y", "%b %d %Y"):
                try:
                    return datetime.strptime(f"{m} {d} {y}", fmt).date()
                except ValueError:
                    continue
            return None
            
        
        
        start_date = parse(start_month, start_day, year)
        end_date = parse(end_month, end_day, year)
    
        if start_date and end_date:
            return [start_date, end_date]
            
        return return [None]*2
    
    def to_dict(self):
        return vars(self)

    def __repr__(self):
        return f"""\nReport = {
            json.dumps({
                key.title(): v.strip()
                for k, v in vars(self).items()
                if isinstance(v, str) and
                (key := k.replace("_str", ""))
            }, indent=2)}\n"""