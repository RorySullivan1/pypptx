"""Bullet formatting, numbered lists, and indentation.

Creates a slide with character bullets at multiple indent levels.
Saves to bullet_lists.pptx.
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank layout

txBox = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(6))
tf = txBox.text_frame
tf.word_wrap = True

# Title paragraph (no bullet)
tf.text = "Project Status Update"
tf.paragraphs[0].font.size = Pt(28)
tf.paragraphs[0].font.bold = True

# Level 0 bullets
items = [
    ("Backend API", [
        "Database migration complete",
        "Authentication module in review",
    ]),
    ("Frontend", [
        "Dashboard redesign shipped",
        "Mobile responsive fixes pending",
    ]),
    ("DevOps", [
        "CI/CD pipeline optimized",
        "Monitoring alerts configured",
    ]),
]

for heading, sub_items in items:
    # Top-level bullet
    p = tf.add_paragraph()
    p.text = heading
    p.font.size = Pt(18)
    p.font.bold = True
    p.bullet_char = "\u2022"  # bullet character
    p.level = 0
    p.space_before = Pt(12)

    # Sub-bullets
    for sub in sub_items:
        p = tf.add_paragraph()
        p.text = sub
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
        p.bullet_char = "\u2013"  # en dash
        p.level = 1

prs.save("bullet_lists.pptx")
print("Saved bullet_lists.pptx")
