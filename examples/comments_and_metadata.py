"""Comments, tags, sections, and custom document properties.

Demonstrates adding comments to slides, setting key-value tags,
creating presentation sections, and custom properties.
Saves to comments_and_metadata.pptx.
"""

from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()

# Create two slides
slide1 = prs.slides.add_slide(prs.slide_layouts[5])
txBox1 = slide1.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(1))
txBox1.text_frame.text = "Slide with comments and tags"
txBox1.text_frame.paragraphs[0].font.size = Pt(24)

slide2 = prs.slides.add_slide(prs.slide_layouts[5])
txBox2 = slide2.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(1))
txBox2.text_frame.text = "Second slide"
txBox2.text_frame.paragraphs[0].font.size = Pt(24)

# Add comments to slide 1
slide1.comments.add("Alice", "AJ", "Please review the layout.", x=100, y=200)
slide1.comments.add("Bob", "BK", "Looks good to me!", x=100, y=400)
print(f"Slide 1 has {len(slide1.comments)} comments")

# Set tags on slide 1 (key-value metadata)
slide1.tags["status"] = "draft"
slide1.tags["reviewer"] = "alice"
print(f"Tags: status={slide1.tags['status']}, reviewer={slide1.tags['reviewer']}")

# Create sections to organize slides
intro = prs.sections.add("Introduction")
intro.add_slide_id(prs.slides[0].slide_id)

body = prs.sections.add("Content")
body.add_slide_id(prs.slides[1].slide_id)
print(f"Sections: {len(prs.sections)}")

# Set custom document properties
prs.custom_properties["Department"] = "Engineering"
prs.custom_properties["Version"] = 2
prs.custom_properties["Approved"] = True
print(f"Department: {prs.custom_properties['Department']}")

prs.save("comments_and_metadata.pptx")
print("Saved comments_and_metadata.pptx")
