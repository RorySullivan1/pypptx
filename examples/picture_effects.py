"""Picture adjustments: brightness, contrast, grayscale, transparent colour and cropping.

Places the same generated image five times and applies one effect to each copy, so the
differences are visible side by side. The image is built in-process with Pillow (already a
hard dependency of pypptx) so the script needs no asset file on disk.

Every adjustment is read back out of the saved file and asserted.

Demonstrates: feature-manifest.md "Pictures — Adjustments" and "Pictures — Cropping".
Saves to picture_effects.pptx.
"""

import io

from PIL import Image

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt

FILENAME = "picture_effects.pptx"
TRANSPARENT = RGBColor(0xFF, 0x00, 0x00)


def swatch_png() -> io.BytesIO:
    """A 240x160 PNG: red left half, blue right half, so effects are easy to see."""
    image = Image.new("RGB", (240, 160), (0xFF, 0x00, 0x00))
    for x in range(120, 240):
        for y in range(160):
            image.putpixel((x, y), (0x1F, 0x5C, 0xA8))
    stream = io.BytesIO()
    image.save(stream, format="PNG")
    stream.seek(0)
    return stream


prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

labels = ["original", "brightness +40%", "contrast +50%", "grayscale", "red transparent"]
pictures = []
for column, label in enumerate(labels):
    left = Inches(0.35 + column * 1.9)
    picture = slide.shapes.add_picture(swatch_png(), left, Inches(2.2), width=Inches(1.7))
    pictures.append(picture)

    caption = slide.shapes.add_textbox(left, Inches(3.5), Inches(1.7), Inches(0.4))
    caption.text_frame.text = label
    caption.text_frame.paragraphs[0].font.size = Pt(10)

original, brighter, contrasted, gray, keyed = pictures

# --- One effect each --------------------------------------------------------------------
# Both adjustments are fractions in -1.0..1.0, not percentages.
brighter.brightness = 0.4
contrasted.contrast = 0.5
gray.is_grayscale = True

# A transparent colour keys out every pixel matching it exactly -- the red half disappears
# and whatever sits behind the picture shows through.
keyed.transparency_color = TRANSPARENT

# Cropping is expressed as the fraction removed from each edge, applied to the original copy.
original.crop_left = 0.25
original.crop_right = 0.1

prs.save(FILENAME)
print(f"Saved {FILENAME}")

# --- Round trip -------------------------------------------------------------------------
reopened = Presentation(FILENAME)
r_pictures = [
    shape
    for shape in reopened.slides[0].shapes
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
]
assert len(r_pictures) == 5, len(r_pictures)
r_original, r_brighter, r_contrasted, r_gray, r_keyed = r_pictures

assert r_brighter.brightness == 0.4, r_brighter.brightness
assert r_contrasted.contrast == 0.5, r_contrasted.contrast
assert r_gray.is_grayscale is True
assert r_keyed.transparency_color == TRANSPARENT, r_keyed.transparency_color

# Each effect stayed on its own picture.
assert r_original.brightness == 0.0, r_original.brightness
assert r_original.contrast == 0.0
assert r_brighter.is_grayscale is False
assert r_original.transparency_color is None

assert round(r_original.crop_left, 4) == 0.25, r_original.crop_left
assert round(r_original.crop_right, 4) == 0.1, r_original.crop_right
assert r_brighter.crop_left == 0.0

print("Round trip verified: 4 adjustments and a crop, each isolated to its own picture")
