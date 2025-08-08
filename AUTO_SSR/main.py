# import re
# from utils import data
# from pathlib import Path

# def minify(html_text):
    # return re.sub(r'>\s+<', '><',
            # re.sub(r'\s+', ' ',
            # re.sub(r'\n+', '',
            # html_text.strip())))

# source_path = Path("template.html")
# destination_path = Path("index.html")

# html_content = source_path.read_text(encoding="utf-8")

# for k, v in data.items():
    # if k != "summary":
        # html_content = html_content.replace(f"{{{k}}}", v)

# for x, row in enumerate(data["summary"]):
    # for y, cell in enumerate(row):
        # html_content = html_content.replace(f"{{{x}|{y}}}", cell)

# destination_path.write_text(minify(html_content), encoding="utf-8")


import re
from pathlib import Path
from utils import data

def minify(text: str) -> str:
    text = re.sub(r'\n+', '', text.strip())
    text = re.sub(r'\s+', ' ', text)
    return re.sub(r'>\s+<', '><', text)

template_path = Path("template.html")
output_path = Path("index.html")

if data:
    html_content = template_path.read_text(encoding="utf-8")
    for k, v in data.items():
        html_content = html_content.replace(k, v)
    
    output_path.write_text(minify(html_content), encoding="utf-8")
    

# def apply_placeholders(template: str, replacements: dict, summary: list) -> str:
    # l, r = '{{', '}}'
    # # Replace {key} placeholders
    # for k, v in replacements.items():
        # if k != "summary":
            # template = template.replace(f'{l} {k} {r}', str(v))

    # # Replace {x|y} placeholders in summary
    # for x, row in enumerate(summary):
        # for y, c in enumerate(row):
            # template = template.replace(f'{l} {x} {y} {r}', str(c))

    # return template

# def main():
    # template_path = Path("template.html")
    # output_path = Path("index.html")
    
    # if data:
        # html_content = template_path.read_text(encoding="utf-8")
        # html_filled = apply_placeholders(html_content, data, data.get("summary", []))
        # output_path.write_text(minify_html(html_filled), encoding="utf-8")
    
# if __name__ == "__main__":
    # main()