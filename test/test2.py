# import os
# import sys
# import cairosvg
# from openpyxl import load_workbook
# from PIL import Image as PILImage
# from openpyxl.drawing.image import Image

# IMG_PATH = "hcclogo.png"
# IMG_WIDTH = 1550
# IMG_HEIGHT = 360

# NEW_IMG_PATH = "resized_image.png"
# WB_BASENAME = "NSB-P2 SSR - TEMPLATE"
# SOURCE_FILE = f"{WB_BASENAME}.xlsx"

# if not os.path.exists(SOURCE_FILE):
    # print(f"Error: File not found - {SOURCE_FILE}")
    # sys.exit(1)

# print(f"Opening workbook: {SOURCE_FILE}")
# wb = load_workbook(SOURCE_FILE)
# ws = wb.active

# if ws.title != "TEMPLATE":
    # ws.title = "TEMPLATE"

# def points_to_pixels(points):
    # return int(points * 96 / 72)  # 1 point = 1/72 inch, 1 inch = 96 pixels

# ws._images.clear()

# FG = "002445"

# svg_data = f"""
# <svg xmlns="http://www.w3.org/2000/svg"
    # preserveAspectRatio="xMidYMid meet"
    # stroke-linecap="round"
    # stroke-linejoin="round"
    # shape-rendering="geometricPrecision"
    # width="{IMG_WIDTH}px"
    # height="{IMG_HEIGHT}px"
    # viewBox="0 0 {int(IMG_WIDTH/10)} {int(IMG_HEIGHT/10)}">
    # <defs>
        # <linearGradient
            # id="blue"
            # gradientTransform="rotate(69)">
            # <stop
                # offset="0%"
                # stop-color="#007ff5" />
            # <stop
                # offset="23%"
                # stop-color="#0061ba" />
            # <stop
                # offset="96%"
                # stop-color="#00387c" />
        # </linearGradient>
        # <linearGradient
            # id="red"
            # gradientTransform="rotate(69)">
            # <stop
                # offset="0%"
                # stop-color="#ff575c" />
            # <stop
                # offset="23%"
                # stop-color="#ff3036" />
            # <stop
                # offset="96%"
                # stop-color="#cd0006" />
        # </linearGradient>
    # </defs>
    # <rect width="100%" height="100%" fill="none" />
    # <g transform="translate(1,3)">
        # <path
            # fill="url(#blue)"
            # d="M13.95 17.883h-2.512c-1.887.094-3.004.418-3.137 3.687v8.325c-8.207 0-8.207 0-8.207-3.72V3.868C.094.148.094.148 8.3.148v8.309c0 3.063 1.25 3.59 3.136 3.7h7.126c1.91-.145 3.136-.637 3.136-3.684V.148c8.242 0 8.207 0 8.207 3.72v22.308c0 3.719.035 3.719-8.207 3.719V21.57c.024-3.27-1.226-3.574-3.137-3.687Zm0 0" />
        # <path
            # fill="url(#red)"
            # d="M15 6.2c-1.05 0-1.05-1.59.047-1.59h5.164V2.456c0-.746 0-2.363-1.492-2.363L11.66.148c-1.867 0-1.867 1.485-1.867 2.23v5.95c0 .742 0 2.36 1.867 2.36l7.059.027c1.492 0 1.492-1.617 1.492-2.36V6.2Zm.016 19.19c-1.051 0-1.051-1.59.046-1.59h5.165v-2.155c0-.743 0-2.36-1.493-2.36l-7.058.05c-1.867 0-1.867 1.49-1.867 2.235v5.95c0 .742 0 2.359 1.867 2.359l7.058.027c1.493 0 1.493-1.617 1.493-2.363v-2.152Zm0 0" />
        # <g
            # fill="#{FG}"
            # stroke="#{FG}"
            # text-rendering="geometricPrecision"
            # font-family="system-ui, sans-serif"
            # font-spacing="1.069"
            # font-weight="800">
            # <text
                # transform="scale(1.026)"
                # stroke-width="1.32"
                # font-size="22"
                # x="35"
                # y="17.69">
        # HILMARC'S
    # </text>
            # <text
                # stroke-width=".32"
                # font-size="8"
                # x="36"
                # y="27.69">
        # CONSTRUCTION CORPORATION
    # </text>
        # </g>
    # </g>
# </svg>
# """

# cairosvg.svg2png(
    # bytestring=svg_data.encode('utf-8'),
    # write_to=IMG_PATH,
    # output_width=IMG_WIDTH,
    # output_height=IMG_HEIGHT,
    # dpi=600
# )

# row_height = ws.row_dimensions[1].height

# if row_height is not None:
    # new_height = points_to_pixels(row_height)
# else:
    # new_height = points_to_pixels(15)  # assume default

# with PILImage.open(IMG_PATH) as img:
    # aspect_ratio = img.height / img.width
    # new_width = int(new_height / aspect_ratio)
    # img = img.resize((new_width, new_height))
    # img.save(NEW_IMG_PATH)

# img = Image(NEW_IMG_PATH)
# img.anchor = 'F1'
# ws.add_image(img)
# wb.save(SOURCE_FILE)

# if os.path.exists(IMG_PATH):
    # os.remove(IMG_PATH)


import os
import sys
from openpyxl import load_workbook
from PIL import Image as PILImage
from openpyxl.drawing.image import Image

# Constants
IMG_PATH = "hcclogo.png"
NEW_IMG_PATH = "resized_image.png"
WB_BASENAME = "NSB-P2 SSR - TEMPLATE"
SOURCE_FILE = f"{WB_BASENAME}.xlsx"

# Check if Excel file exists
if not os.path.exists(SOURCE_FILE):
    print(f"Error: File not found - {SOURCE_FILE}")
    sys.exit(1)

# Open workbook and sheet
print(f"Opening workbook: {SOURCE_FILE}")
wb = load_workbook(SOURCE_FILE)
ws = wb.active

# Rename sheet if needed
if ws.title != "TEMPLATE":
    ws.title = "TEMPLATE"

# Function to convert points to pixels
def points_to_pixels(points):
    return int(points * 96 / 72)

# Clear existing images
ws._images.clear()

# Get row height in pixels
row_height = int((ws.row_dimensions[1].height) * 0.85)
new_height = points_to_pixels(row_height) if row_height else points_to_pixels(15)

# Resize the image to match row height
with PILImage.open(IMG_PATH) as img:
    aspect_ratio = img.height / img.width
    new_width = int(new_height / aspect_ratio)
    img = img.resize((new_width, new_height))
    img.save(NEW_IMG_PATH)

# Insert image into cell F1
img = Image(NEW_IMG_PATH)
img.anchor = 'F1'
ws.add_image(img)

# Save workbook
wb.save(SOURCE_FILE)
print("Image inserted and workbook saved.")

if os.path.exists(NEW_IMG_PATH):
    os.remove(NEW_IMG_PATH)