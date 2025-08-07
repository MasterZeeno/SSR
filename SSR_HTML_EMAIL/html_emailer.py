import difflib
import os
import re
import smtplib
import sys
import html
from html.parser import HTMLParser
from email.message import EmailMessage
from email.utils import formataddr
from mimetypes import guess_type
from types import MappingProxyType

from html_builder import HTML_BODY, SUBJECT, MSGS, EXCEL_FILE_PATH, ZEE

class HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.fed = []

    def handle_data(self, data):
        self.fed.append(data)

    def get_data(self):
        return html.unescape(''.join(self.fed))

def clean(html_text):
    stripper = HTMLStripper()
    stripper.feed(html_text)
    return stripper.get_data()

EXCEL_FILE = EXCEL_FILE_PATH

if not os.path.isfile(EXCEL_FILE):
    raise FileNotFoundError(f"File not found — {EXCEL_FILE}")

class CONST:
    SENSITIVE_KEYWORDS = [
        'pass', 'password', 'passwd', 'pwd',
        'token', 'secret', 'apikey', 'api_key',
        'access', 'private', 'secure', 'auth',
        'authentication', 'credentials', 'cred'
    ]

    def __init__(self, data):
        if isinstance(data, str):
            data = [data]
        if isinstance(data, (list, tuple, set)):
            data = {k: None for k in data}

        if not isinstance(data, dict):
            raise TypeError(f"{self.__class__.__name__} - invalid input values:\n\n{data}")
        
        _data = {}
        for key, val in data.items():
            if 'alias' in key and isinstance(val, (list, tuple, set)):
                for k, v in zip(val[::2], val[1::2]):
                    self.SENSITIVE_KEYWORDS.append(k)
                    _data[k] = data[v]
            else:
                _data[key] = val
        
        self._data = MappingProxyType(_data)
        
        keys = set(self._data.keys())
        matches = set()
        for keyword in set(self.SENSITIVE_KEYWORDS):
            close = difflib.get_close_matches(keyword, keys, n=5, cutoff=0.69)
            matches.update(close)

        self._excluded = matches
    
    def __get_value(self, key):
        try:
            if isinstance(key, int):
                values = [v for v in self._data.values()]
                return values[min(max(key,0), len(values)-1)]
            else:
                return self._data[key]
        except KeyError:
            return None
        
    def __getattr__(self, key):
        return self.__get_value(key)

    def __getitem__(self, key):
        return self.__get_value(key)
    
    def __repr__(self):
        if all(v is None for v in self._data.values()):
            return ', '.join(filter(None, self._data.values()))
        
        lines = []
        pad = lambda k: ' ' * (len(k) + 4)
        exclude = set(self.__exclude())
        
        for k, v in self._data.items():
            if k in exclude:
                continue
            val = '' if v is None else v.replace(', ', f"',\n{pad(k)}'")
            lines.append(f"  {k}: '{val}',")
        
        return '{\n' + '\n'.join(lines).rstrip(',') + '\n}'

    def __contains__(self, key):
        return key in self._data

    def __iter__(self):
        return iter(self._data)
        
    def __exclude(self, exclude=None):
        excluded = set(self._excluded)  # make a copy to avoid mutation

        if isinstance(exclude, str):
            excluded.add(exclude)
        elif isinstance(exclude, (list, tuple, set)):
            excluded.update(exclude)
        elif exclude is not None:
            raise TypeError("'exclude' must be a string, list, tuple, set, or None")

        return excluded
    
    def __retrieve(self, pair=True, keys=None, exclude=None):
        keys = [keys] if isinstance(keys, str) else list(keys) if isinstance(keys, (list, tuple, set)) else [k for k in self._data if k not in self.__exclude(exclude)]
        return [(k, v) if pair else v for k, v in self._data.items() if k in keys]
    
    def as_dict(self):
        return dict(self._data)

    def keys(self, exclude=None):
        return [k for k in self._data if k not in self.__exclude(exclude)]

    def values(self, keys=None, exclude=None):
        return self.__retrieve(False, keys, exclude)

    def items(self, keys=None, exclude=None):
        return self.__retrieve(True, keys, exclude)

    def get(self, key, default=None):
        search = str(default).strip() if default else None

        if search in self._data.values():
            default_val = search
        else:
            default_val = None
            for k in (search, 'default', next(iter(self._data), None)):
                if k in self._data:
                    default_val = self._data[k]
                    break

        return self._data.get(str(key).strip(), default_val)

SUFFXS = CONST({
    'hcc': 'com.ph',
    'default': 'com'
})

PROVIDERS = tuple(
    f"{S}.{SUFFXS.get(S)}"
    for S in ('hcc', 'gmail', 'yahoo', 'outlook',
        'icloud', 'protonmail', 'mail', 'yandex')
)

PORT, SERVER = 587, f"smtp.{PROVIDERS[1]}"

def eaddrs(S=None, T=0, PS=PROVIDERS):
    """
    Builds a properly formatted email address from a string name and provider index.
    Returns:
        The cleaned email string if valid, None otherwise.
    """
    if not isinstance(S, str):
        return None
    T = 0 if not isinstance(T, int) or not (0 <= T < len(PS)) else T
    S, P = f"{S}@{PS[T]}", r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return S if re.match(P, S) else None

def estr(S=None, T=0):
    """
    Returns a properly formatted email string:
    - If list: join with commas
    - If email-like string: return as-is
    - If name string: convert to email via eaddrs()
    - Else: return None
    """
    if isinstance(S, list):
        return ', '.join(estr(item, T) for item in S if estr(item, T))
    elif isinstance(S, str):
        return S if '@' in S else eaddrs(S, T)
    return None
    
# --- CONFIGURATION ---
OPERATION = sys.argv[1] if len(sys.argv) > 1 else ""

CFG = {
    "subject": SUBJECT,
    "sender": estr('zeenoliev',1),
    "from": formataddr((
        ZEE["name"],
        estr('zeenoliev',1)
    )),
    "password": "frmoyroohmevbgvb",
}

if "force" in OPERATION.lower():
    CFG.update({
        "to": ', '.join([
            estr("jojofundales"),
            estr("jojofundales",2)
        ]),
        "cc": ', '.join([
            estr(["arch_rbporral","glachel.arao","rbzden"],2),
            estr(["maravilladarwin87.dm@gmail.com", "aljonporcalla","eduardo111680"],1)
        ])
    })
else:
    CFG.update({
        "to": estr(["yawapisting7", "cimaciojay0"],1),
        "cc": estr(["rayajcimacio", "rayajcimacio2"],1)
    })

CFG = CONST(CFG)

# --- BUILD EMAIL ---
msg = EmailMessage()
for k, v in CFG.items():
    msg[k.title()] = v
msg.set_content(clean('\n'.join([*MSGS[:2], MSGS[-1]])))
msg.add_alternative(HTML_BODY, subtype='html')

# --- ADD EXCEL_FILE ---
mime_type, _ = guess_type(EXCEL_FILE)
maintype, subtype = mime_type.split('/') if mime_type else ('application', 'octet-stream')

with open(EXCEL_FILE, 'rb') as f:
    file_data = f.read()
    file_name = os.path.basename(EXCEL_FILE)
    msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)

# --- SEND EMAIL ---
try:
    with smtplib.SMTP(SERVER, PORT) as smtp:
        smtp.starttls()
        smtp.login(CFG.sender, CFG.password)
        smtp.send_message(msg)
        print("Email sent successfully.")
    
except Exception as e:
    print(f"Error sending email: {e}")