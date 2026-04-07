"""Text boxes with rich font formatting and paragraph alignment.

Adds a text box with multiple paragraphs demonstrating bold, italic,
color, font size, underline, and text alignment. Saves to add_textbox.pptx.
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank layout

txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(4))
tf = txBox.text_frame
tf.word_wrap = True

# First paragraph — large bold title
tf.text = "Text Formatting Demo"
p0 = tf.paragraphs[0]
p0.font.size = Pt(28)
p0.font.bold = True
p0.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)

# Second paragraph — italic with color
p1 = tf.add_paragraph()
p1.text = "This paragraph is italic and blue."
p1.font.italic = True
p1.font.color.rgb = RGBColor(0x00, 0x70, 0xC0)
p1.font.size = Pt(18)

# Third paragraph — underlined and right-aligned
p2 = tf.add_paragraph()
p2.text = "Right-aligned and underlined."
p2.font.underline = True
p2.font.size = Pt(16)
p2.alignment = PP_ALIGN.RIGHT

# Fourth paragraph — centered with custom color
p3 = tf.add_paragraph()
p3.text = "Centered text in dark red."
p3.font.color.rgb = RGBColor(0x8B, 0x00, 0x00)
p3.font.size = Pt(20)
p3.alignment = PP_ALIGN.CENTER

prs.save("add_textbox.pptx")
print("Saved add_textbox.pptx")
