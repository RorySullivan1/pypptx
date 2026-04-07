"""Duplicate, reorder, and delete slides.

Creates a presentation with several slides, then demonstrates slide
duplication, reordering, and deletion. Saves to slide_operations.pptx.
"""

from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()

# Create three slides with labels
for i in range(1, 4):
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank layout
    txBox = slide.shapes.add_textbox(Inches(3), Inches(3), Inches(4), Inches(1))
    txBox.text_frame.text = f"Slide {i}"
    txBox.text_frame.paragraphs[0].font.size = Pt(32)

print(f"Starting with {len(prs.slides)} slides")

# Duplicate slide 1
dup = prs.slides.duplicate(prs.slides[0])
# Update the duplicated slide's label
for shape in dup.shapes:
    if shape.has_text_frame and shape.text_frame.text == "Slide 1":
        shape.text_frame.text = "Slide 1 (copy)"
print(f"After duplicate: {len(prs.slides)} slides")

# Reorder: move the last slide (index 3) to position 1
prs.slides.move(3, 1)
print("Moved duplicated slide to position 1")

# Delete the slide at index 2
prs.slides.delete(prs.slides[2])
print(f"After delete: {len(prs.slides)} slides")

prs.save("slide_operations.pptx")
print("Saved slide_operations.pptx")
