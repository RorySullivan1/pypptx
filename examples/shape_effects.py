"""Shadow, glow, reflection, and 3D formatting on shapes.

Demonstrates visual effects applied to autoshapes. Saves to shape_effects.pptx.
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_BEVEL_PRESET
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank layout

# 1. Shape with outer shadow
box1 = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1), Inches(2.5), Inches(1.5)
)
box1.fill.solid()
box1.fill.fore_color.rgb = RGBColor(0x44, 0x72, 0xC4)
box1.text_frame.text = "Shadow"
shadow = box1.shadow
shadow.blur_radius = Pt(8)
shadow.distance = Pt(5)
shadow.direction = 315
shadow.color.rgb = RGBColor(0x00, 0x00, 0x00)

# 2. Shape with glow effect
box2 = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.OVAL, Inches(3.7), Inches(1), Inches(2.5), Inches(1.5)
)
box2.fill.solid()
box2.fill.fore_color.rgb = RGBColor(0xED, 0x7D, 0x31)
box2.text_frame.text = "Glow"
box2.glow.radius = Pt(10)

# 3. Shape with soft edge
box3 = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.5), Inches(3.5), Inches(2.5), Inches(1.5)
)
box3.fill.solid()
box3.fill.fore_color.rgb = RGBColor(0x00, 0xB0, 0x50)
box3.text_frame.text = "Soft Edge"
box3.soft_edge.radius = Pt(12)

# 4. Shape with 3D bevel
box4 = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(3.7), Inches(3.5), Inches(2.5), Inches(1.5)
)
box4.fill.solid()
box4.fill.fore_color.rgb = RGBColor(0x7B, 0x2D, 0x8E)
box4.text_frame.text = "3D Bevel"
bevel = box4.three_d.get_or_add_bevel_top()
bevel.preset = MSO_BEVEL_PRESET.CIRCLE
bevel.width = Pt(8)
bevel.height = Pt(4)
box4.three_d.extrusion_height = Pt(6)

prs.save("shape_effects.pptx")
print("Saved shape_effects.pptx")
