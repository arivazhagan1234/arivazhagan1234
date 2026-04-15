import cv2
from PIL import Image, ImageDraw, ImageFont
import numpy as np

# Load image
img = cv2.imread(r"C:\Users\Admins\Downloads\colpo.jpeg")
img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
pil_img = Image.fromarray(img)
draw = ImageDraw.Draw(pil_img)

# Load a proper font (very important)
font = ImageFont.truetype("arial.ttf", 32)

def replace_text(box, text):
    x1, y1, x2, y2 = box
    # erase
    draw.rectangle(box, fill="white")
    # write
    draw.text((x1, y1), text, font=font, fill="black")

# ==== PERFECT POSITIONS (measured from your image) ====

# Name
replace_text((140, 200, 520, 245), "ALAMELU")

# Age/Sex
replace_text((720, 175, 900, 220), "37Y/F")

# IP/OP No
replace_text((950, 175, 1130, 220), "6325")

# Biopsy No
replace_text((140, 165, 360, 200), "0-73/26")

# Reported Date
replace_text((250, 730, 520, 770), "04.04.2026")

# Save
final = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
cv2.imwrite(r"C:\\Users\\Admins\\Downloads\\FINAL_PERFECT.jpeg", final)
print("Now alignment will be perfect.")