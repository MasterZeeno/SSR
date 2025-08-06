import pandas as pd

filename = "PE-01-NSBP2-23 SSR"
df = pd.read_excel(f"../{filename}.xlsx")
df.to_html(f"{filename}.html", index=False)  # Saves an HTML table