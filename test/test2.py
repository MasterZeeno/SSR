import os
import sys
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image

# Function to convert points to pixels
def points_to_pixels(points):
    return int(points * 96 / 72)

# Constants
BASE_PTS = 69
DIMENS = {
 "A": 0.67, "B": 8.14, "C": 7.71, "D": 15, "E": 10.43,
 "F": 12.29, "G": 12.29, "H": 12.29, "I": 12.29, "J": 12.29,
 "K": 12.29, "L": 5.14, "M": 2, "N": 0.67
}

IMG_PATH = "hcclogo.png"
NEW_IMG_PATH = "resized_image.png"
WB_BASENAME = "NSB-P2 SSR - TEMPLATE"
DEST_FILE = f"{WB_BASENAME}.xlsx"

print("Creating new workbook...")

wb = Workbook()
ws = wb.active
ws.title = "TEMPLATE"

# ws.row_dimensions[1].height = 60
# Apply column widths
for col, wid in DIMENS.items():
    ws.column_dimensions[col].width = wid
    
# Clear existing images
ws._images.clear()

# Get image height in pixels
IMG_HEIGHT = points_to_pixels(int(BASE_PTS * 0.69))
# Resize the image to match row height
with PILImage.open(IMG_PATH) as img:
    ASPECT_RATIO = img.height / img.width
    IMG_WIDTH = int(IMG_HEIGHT / ASPECT_RATIO)
    img = img.resize((IMG_WIDTH, IMG_HEIGHT))
    img.save(NEW_IMG_PATH)

# Insert image into cell F1
img = Image(NEW_IMG_PATH)
img.anchor = 'F1'
ws.add_image(img)

# Save workbook
wb.save(DEST_FILE)
print("Image inserted and workbook saved.")

if os.path.exists(NEW_IMG_PATH):
    os.remove(NEW_IMG_PATH)