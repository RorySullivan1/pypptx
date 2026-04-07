"""AutoShapes with solid, gradient, and pattern fills plus line styles.

Adds four shapes with different fill types and line formatting.
Saves to shapes_and_fills.pptx.
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE, MSO_PATTERN_TYPE, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank layout

# 1. Rectangle with solid fill
rect = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.5), Inches(1), Inches(2), Inches(1.5)
)
rect.fill.solid()
rect.fill.fore_color.rgb = RGBColor(0x44, 0x72, 0xC4)
rect.text_frame.text = "Solid Fill"

# 2. Oval with gradient fill
oval = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.OVAL, Inches(3.2), Inches(1), Inches(2), Inches(1.5)
)
oval.fill.gradient()
oval.fill.gradient_angle = 45.0
stops = oval.fill.gradient_stops
stops[0].color.rgb = RGBColor(0x00, 0xB0, 0x50)
stops[1].color.rgb = RGBColor(0xFF, 0xFF, 0x00)
oval.text_frame.text = "Gradient"

# 3. Rounded rectangle with pattern fill
rrect = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(5.9), Inches(1), Inches(2), Inches(1.5)
)
rrect.fill.patterned()
rrect.fill.pattern = MSO_PATTERN_TYPE.DARK_DOWNWARD_DIAGONAL
rrect.fill.fore_color.rgb = RGBColor(0xED, 0x7D, 0x31)
rrect.text_frame.text = "Pattern"

# 4. Chevron with no fill (background) and dashed line
chevron = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.CHEVRON, Inches(1.5), Inches(3.5), Inches(3), Inches(2)
)
chevron.fill.background()
chevron.line.color.rgb = RGBColor(0xC0, 0x00, 0x00)
chevron.line.width = Pt(3)
chevron.line.dash_style = MSO_LINE_DASH_STYLE.DASH_DOT
chevron.text_frame.text = "No Fill + Dashed Line"

prs.save("shapes_and_fills.pptx")
print("Saved shapes_and_fills.pptx")
