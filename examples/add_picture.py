"""Insert and format images in a presentation.

Generates a small image using Pillow (no external files needed), inserts it
onto a slide with explicit sizing, and applies a line border. Saves to
add_picture.pptx.
"""

from io import BytesIO

from PIL import Image, ImageDraw

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

# Generate a simple gradient image in memory
img = Image.new("RGB", (400, 300))
draw = ImageDraw.Draw(img)
for y in range(300):
    r = int(255 * y / 300)
    draw.line([(0, y), (400, y)], fill=(r, 100, 255 - r))
draw.text((120, 130), "pypptx", fill="white")
buf = BytesIO()
img.save(buf, format="PNG")
buf.seek(0)

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank layout

# Insert the image with explicit position and size
pic = slide.shapes.add_picture(buf, Inches(1.5), Inches(1.5), Inches(5), Inches(3.75))

# Add a border around the picture
pic.line.color.rgb = RGBColor(0x40, 0x40, 0x40)
pic.line.width = Pt(2)

prs.save("add_picture.pptx")
print("Saved add_picture.pptx")
