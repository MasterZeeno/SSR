import re

EXCEPTION_WORDS = {
    # Tech brands
    "iPhone", "iPad", "iMac", "eBay", "YouTube", "GitHub", "ChatGPT", "WiFi", "macOS", "Android", "iOS",
    "JavaScript", "TypeScript", "Node.js", "React", "Vue", "Nuxt", "Next.js",
    "OpenAI", "VSCode", "PyTorch", "TensorFlow", "Google", "Facebook", "Twitter", "X", "Instagram", "TikTok",

    # File formats & extensions
    "PDF", "JPEG", "PNG", "GIF", "MP4", "MP3", "ZIP", "CSV", "HTML", "CSS", "JSON", "XML",

    # Acronyms / Programs
    "NASA", "FBI", "CIA", "UNESCO", "UNICEF", "WHO", "USAID", "DOH", "IT", "AI", "ML", "UI", "UX", "API", "SQL",

    # Software / Platforms
    "Windows", "Linux", "Ubuntu", "Debian", "Kali", "Termux", "Git", "GitHub", "npm", "pip", "Docker", "WSL",

    # Others
    "Jay Ar Adlaon Cimacio, RN", "Zeeno", "MasterZeeno", "COVID-19", "CCTV", "PDFViewer", "IDMaker", "HCC", "SSR", "EHS"
}

# Optional: Roman numeral regex (up to 3999)
ROMAN_NUMERAL_RE = re.compile(
    r'^(M{0,3})(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$', re.IGNORECASE
)
ROMAN_NUMERAL_RE = re.compile(
    r"""^
    M{0,4}               
    (CM|CD|D?C{0,3})
    (XC|XL|L?X{0,3})
    (IX|IV|V?I{0,3})$
    """, re.VERBOSE | re.IGNORECASE
)

def title(text: str, strict: bool = False) -> str:
    def is_exception(word):
        return word in EXCEPTION_WORDS

    def is_acronym_or_roman(word):
        return word.isupper() or bool(ROMAN_NUMERAL_RE.fullmatch(word))

    def should_skip(word):
        return (
            '@' in word or                 # emails
            word.startswith('http') or    # URLs
            word.startswith('#') or       # hashtags
            word.isnumeric()
        )

    def transform(word):
        # Keep exceptions intact
        if is_exception(word):
            return word
        if should_skip(word):
            return word
        if strict:
            return word.capitalize()
        if is_acronym_or_roman(word):
            return word
        return word[0].upper() + word[1:].lower() if len(word) > 1 else word.upper()

    # Handle hyphenated and apostrophe-joined words properly
    def fix_complex_word(match):
        base = match.group()
        return re.sub(r"[\w]+", lambda m: transform(m.group()), base)

    return re.sub(r"\b[\w'-]+\b", fix_complex_word, text)