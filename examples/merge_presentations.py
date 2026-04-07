"""Import and merge slides across presentations.

Creates two source presentations, imports a slide from one into the other,
then merges all slides from a third. Saves to merge_presentations.pptx.
"""

from pptx import Presentation
from pptx.util import Inches, Pt

# Create source presentation A
prs_a = Presentation()
slide_a = prs_a.slides.add_slide(prs_a.slide_layouts[5])
txBox = slide_a.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(1))
txBox.text_frame.text = "From Presentation A"
txBox.text_frame.paragraphs[0].font.size = Pt(28)

# Create source presentation B with two slides
prs_b = Presentation()
for label in ["From Presentation B - Slide 1", "From Presentation B - Slide 2"]:
    slide_b = prs_b.slides.add_slide(prs_b.slide_layouts[5])
    txBox = slide_b.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(1))
    txBox.text_frame.text = label
    txBox.text_frame.paragraphs[0].font.size = Pt(28)

# Create target and import one slide from A
target = Presentation()
target.slides.import_slide(prs_a.slides[0])
print(f"After importing from A: {len(target.slides)} slide(s)")

# Merge all slides from B
target.slides.merge(prs_b)
print(f"After merging B: {len(target.slides)} slide(s)")

target.save("merge_presentations.pptx")
print("Saved merge_presentations.pptx")
