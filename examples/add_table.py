"""Tables with merged cells, borders, and cell fills.

Creates a data table with a merged header row, alternating row colors,
and border formatting. Saves to add_table.pptx.
"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank layout

# Create a 4-row, 3-column table
rows, cols = 4, 3
table_shape = slide.shapes.add_table(rows, cols, Inches(1.5), Inches(1.5), Inches(6), Inches(3))
table = table_shape.table

# Merge the first row across all columns for a header
table.cell(0, 0).merge(table.cell(0, 2))
header_cell = table.cell(0, 0)
header_cell.text = "Quarterly Sales Report"
header_cell.fill.solid()
header_cell.fill.fore_color.rgb = RGBColor(0x1F, 0x4E, 0x79)
for p in header_cell.text_frame.paragraphs:
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.font.size = Pt(16)
    p.font.bold = True

# Column headers
headers = ["Region", "Q1 Revenue", "Q2 Revenue"]
for i, text in enumerate(headers):
    cell = table.cell(1, i)
    cell.text = text
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(0xD6, 0xE4, 0xF0)
    for p in cell.text_frame.paragraphs:
        p.font.bold = True
        p.font.size = Pt(12)

# Data rows
data = [
    ["North", "$1,250,000", "$1,380,000"],
    ["South", "$980,000", "$1,050,000"],
]
for row_idx, row_data in enumerate(data, start=2):
    for col_idx, text in enumerate(row_data):
        cell = table.cell(row_idx, col_idx)
        cell.text = text
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(11)

# Enable banding
table.first_row = True
table.horz_banding = True

prs.save("add_table.pptx")
print("Saved add_table.pptx")
