"""Connectors with arrowheads and group shapes.

Creates shapes, connects them with an arrowed connector, and groups
shapes together. Saves to connectors_and_groups.pptx.
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_END_TYPE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR_TYPE
from pptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank layout

# Create two boxes to connect
box1 = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(1), Inches(2), Inches(2), Inches(1.2)
)
box1.fill.solid()
box1.fill.fore_color.rgb = RGBColor(0x44, 0x72, 0xC4)
box1.text_frame.text = "Start"
for p in box1.text_frame.paragraphs:
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

box2 = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(6), Inches(2), Inches(2), Inches(1.2)
)
box2.fill.solid()
box2.fill.fore_color.rgb = RGBColor(0x00, 0xB0, 0x50)
box2.text_frame.text = "End"
for p in box2.text_frame.paragraphs:
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

# Add a straight connector between them with an arrowhead
connector = slide.shapes.add_connector(
    MSO_CONNECTOR_TYPE.STRAIGHT,
    Inches(3), Inches(2.6),  # begin point (right edge of box1)
    Inches(6), Inches(2.6),  # end point (left edge of box2)
)
connector.line.color.rgb = RGBColor(0x40, 0x40, 0x40)
connector.line.width = Pt(2)
connector.line.end_arrowhead_type = MSO_LINE_END_TYPE.TRIANGLE

# Create shapes to group (bottom of slide)
circle1 = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.OVAL, Inches(2), Inches(4.5), Inches(1), Inches(1)
)
circle1.fill.solid()
circle1.fill.fore_color.rgb = RGBColor(0xFF, 0x69, 0x00)

circle2 = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.OVAL, Inches(3.5), Inches(4.5), Inches(1), Inches(1)
)
circle2.fill.solid()
circle2.fill.fore_color.rgb = RGBColor(0xFF, 0xD7, 0x00)

# Group the two circles
group = slide.shapes.add_group_shape([circle1, circle2])
print(f"Group contains shapes at group level")

prs.save("connectors_and_groups.pptx")
print("Saved connectors_and_groups.pptx")
