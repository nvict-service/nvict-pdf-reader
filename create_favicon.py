#!/usr/bin/env python3
"""
Create favicon.ico for NVict Reader application
Generates a professional icon with PDF document styling
"""

from PIL import Image, ImageDraw, ImageFont
import os

# Create assets folder if it doesn't exist
os.makedirs("assets", exist_ok=True)

# Define icon sizes and colors
size = 256  # Create 256x256, will be converted to ico
bg_color = "#10a2dd"  # NVict Reader accent color (blue)
text_color = "#ffffff"  # White text

# Create image
img = Image.new("RGB", (size, size), color=bg_color)
draw = ImageDraw.Draw(img)

# Draw a PDF document shape (simplified)
# Left margin
margin = 20
doc_left = margin
doc_top = margin
doc_right = size - margin
doc_bottom = size - margin

# Draw rounded rectangle background (document shape)
draw.rectangle(
    [doc_left, doc_top, doc_right, doc_bottom],
    fill=bg_color,
    outline="#0880b8",
    width=3
)

# Draw document fold/corner
fold_size = 30
draw.polygon(
    [(doc_right - fold_size, doc_top),
     (doc_right, doc_top),
     (doc_right, doc_top + fold_size)],
    fill="#0880b8"
)

# Add "PDF" text - try to use a font, fallback to default
text = "PDF"
try:
    # Try to load a bold font
    font_size = 80
    font = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", font_size)
except:
    # Fallback to default font
    font = ImageFont.load_default()
    text = "PDF"

# Get text bounding box to center it
bbox = draw.textbbox((0, 0), text, font=font)
text_width = bbox[2] - bbox[0]
text_height = bbox[3] - bbox[1]

# Center text in document
text_x = (size - text_width) // 2
text_y = (size - text_height) // 2 + 10

# Draw text
draw.text((text_x, text_y), text, fill=text_color, font=font)

# Save as icon file (PIL will handle ico conversion)
ico_path = os.path.join("assets", "favicon.ico")
img.save(ico_path, format="ICO")

print(f"✓ favicon.ico created successfully!")
print(f"  Location: {os.path.abspath(ico_path)}")
print(f"  Size: {os.path.getsize(ico_path)} bytes")
