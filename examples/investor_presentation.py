"""Multi-slide investor presentation with charts, tables, and styled text.

Creates a five-slide deck: cover page, investment rationale, performance
overview (bulleted highlights, line chart, and data table), drawdown
analysis, and a contact page. Saves to investor_presentation.pptx.
"""

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

# -- Palette --
NAVY = RGBColor(0x0B, 0x1D, 0x3A)
DARK_BLUE = RGBColor(0x14, 0x2D, 0x54)
ACCENT_BLUE = RGBColor(0x1B, 0x6B, 0xB5)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
MID_GRAY = RGBColor(0x58, 0x58, 0x58)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x1A, 0x9E, 0x5C)
RED = RGBColor(0xC0, 0x39, 0x2B)
GOLD = RGBColor(0xD4, 0xA0, 0x1E)

prs = Presentation()

# ======================================================================
# Helper: add a dark-background slide with an optional title bar
# ======================================================================

def _dark_slide(prs, title_text=None):
    """Return a blank slide with a navy background and optional title."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = NAVY
    if title_text:
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.6))
        tf = tb.text_frame
        tf.text = title_text
        p = tf.paragraphs[0]
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = WHITE
        # accent underline
        line = slide.shapes.add_shape(
            1,  # MSO_AUTO_SHAPE_TYPE.RECTANGLE
            Inches(0.6), Inches(0.95), Inches(1.6), Pt(3),
        )
        line.fill.solid()
        line.fill.fore_color.rgb = ACCENT_BLUE
        line.line.fill.background()
    return slide


# ======================================================================
# Slide 1 — Cover Page
# ======================================================================
cover = _dark_slide(prs)

# Fund name
tb = cover.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(7.6), Inches(1.0))
tf = tb.text_frame
tf.word_wrap = True
tf.text = "Apex Capital Partners"
p = tf.paragraphs[0]
p.font.size = Pt(40)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

# Subtitle
tb2 = cover.shapes.add_textbox(Inches(1.2), Inches(3.0), Inches(7.6), Inches(0.8))
tf2 = tb2.text_frame
tf2.word_wrap = True
tf2.text = "Global Equity Strategy \u2014 Investor Presentation"
p2 = tf2.paragraphs[0]
p2.font.size = Pt(20)
p2.font.color.rgb = ACCENT_BLUE
p2.alignment = PP_ALIGN.CENTER

# Date
tb3 = cover.shapes.add_textbox(Inches(1.2), Inches(3.8), Inches(7.6), Inches(0.5))
tf3 = tb3.text_frame
tf3.text = "Q1 2026"
p3 = tf3.paragraphs[0]
p3.font.size = Pt(16)
p3.font.color.rgb = RGBColor(0x88, 0x99, 0xAA)
p3.alignment = PP_ALIGN.CENTER

# Decorative line
accent = cover.shapes.add_shape(
    1, Inches(3.5), Inches(4.5), Inches(3.0), Pt(2),
)
accent.fill.solid()
accent.fill.fore_color.rgb = ACCENT_BLUE
accent.line.fill.background()

# Confidential footer
tb4 = cover.shapes.add_textbox(Inches(0), Inches(6.8), Inches(10), Inches(0.4))
tf4 = tb4.text_frame
tf4.text = "CONFIDENTIAL \u2014 For Qualified Investors Only"
p4 = tf4.paragraphs[0]
p4.font.size = Pt(9)
p4.font.color.rgb = RGBColor(0x66, 0x77, 0x88)
p4.alignment = PP_ALIGN.CENTER

# ======================================================================
# Slide 2 — Investment Rationale
# ======================================================================
rationale = _dark_slide(prs, "Investment Rationale")

bullets = [
    ("Differentiated Strategy", [
        "Systematic long/short equity approach across developed markets",
        "Proprietary factor model blending value, momentum, and quality signals",
    ]),
    ("Robust Risk Management", [
        "Dynamic portfolio hedging with real-time exposure monitoring",
        "Maximum single-name concentration of 3%; sector cap of 20%",
    ]),
    ("Alignment of Interests", [
        "Partners have over $50 M co-invested alongside clients",
        "High-water mark with 18-month crystallization period",
    ]),
]

tx = rationale.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(8.4), Inches(5.2))
tf = tx.text_frame
tf.word_wrap = True

for heading, subs in bullets:
    p = tf.add_paragraph()
    p.text = heading
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = ACCENT_BLUE
    p.bullet_char = "\u25A0"  # filled square
    p.level = 0
    p.space_before = Pt(14)
    for sub in subs:
        sp = tf.add_paragraph()
        sp.text = sub
        sp.font.size = Pt(13)
        sp.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
        sp.bullet_char = "\u2013"
        sp.level = 1
        sp.space_before = Pt(4)

# ======================================================================
# Slide 3 — Performance
# ======================================================================
perf = _dark_slide(prs, "Performance Overview")

# --- Bulleted highlights (left column) ---
hl = perf.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(3.6), Inches(2.4))
hlf = hl.text_frame
hlf.word_wrap = True
hlf.text = "Key Highlights"
hlf.paragraphs[0].font.size = Pt(14)
hlf.paragraphs[0].font.bold = True
hlf.paragraphs[0].font.color.rgb = GOLD

highlights = [
    "Annualized net return of 14.2% since inception",
    "Sharpe ratio of 1.35 (vs. 0.72 for benchmark)",
    "Positive returns in 78% of rolling 12-month periods",
    "Max drawdown limited to \u221212.4%",
]
for h in highlights:
    p = hlf.add_paragraph()
    p.text = h
    p.font.size = Pt(11)
    p.font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
    p.bullet_char = "\u2022"
    p.level = 0
    p.space_before = Pt(6)

# --- Line chart (right column) ---
chart_data = CategoryChartData()
chart_data.categories = [
    "2020", "2021", "2022", "2023", "2024", "2025",
]
chart_data.add_series("Apex Global Equity", (100, 118.5, 109.3, 132.6, 151.0, 168.4))
chart_data.add_series("MSCI World (TR)", (100, 114.2, 98.7, 117.5, 128.3, 139.6))

chart_frame = perf.shapes.add_chart(
    XL_CHART_TYPE.LINE_MARKERS,
    Inches(4.5), Inches(1.3), Inches(5.2), Inches(2.7),
    chart_data,
)
chart = chart_frame.chart
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
chart.chart_title.has_text_frame = True
chart.chart_title.text_frame.text = "Growth of $100"
chart.chart_title.text_frame.paragraphs[0].font.size = Pt(10)

# --- Performance table (full width, bottom) ---
rows, cols = 4, 7
tbl_shape = perf.shapes.add_table(
    rows, cols, Inches(0.6), Inches(4.3), Inches(8.8), Inches(2.2),
)
tbl = tbl_shape.table

col_headers = ["", "2020", "2021", "2022", "2023", "2024", "2025"]
fund_returns = ["Apex Global Equity", "12.8%", "18.5%", "\u22127.8%", "21.3%", "13.9%", "11.5%"]
bench_returns = ["MSCI World (TR)", "10.1%", "14.2%", "\u221213.3%", "19.1%", "9.2%", "8.8%"]

for ci, text in enumerate(col_headers):
    cell = tbl.cell(0, ci)
    cell.text = text

for ci, text in enumerate(fund_returns):
    tbl.cell(1, ci).text = text
for ci, text in enumerate(bench_returns):
    tbl.cell(2, ci).text = text

# Excess return row
excess = ["Excess Return", "+2.7%", "+4.3%", "+5.5%", "+2.2%", "+4.7%", "+2.7%"]
for ci, text in enumerate(excess):
    tbl.cell(3, ci).text = text

# Style header row
for ci in range(cols):
    cell = tbl.cell(0, ci)
    cell.fill.solid()
    cell.fill.fore_color.rgb = DARK_BLUE
    for p in cell.text_frame.paragraphs:
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = WHITE

# Style data rows
for ri in range(1, rows):
    for ci in range(cols):
        cell = tbl.cell(ri, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x0F, 0x24, 0x40) if ri % 2 == 1 else RGBColor(0x14, 0x2D, 0x54)
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(9)
            p.font.color.rgb = WHITE if ci == 0 else (GREEN if not text.startswith("\u2212") else RED)

# Color excess row green
for ci in range(1, cols):
    for p in tbl.cell(3, ci).text_frame.paragraphs:
        p.font.color.rgb = GREEN
        p.font.bold = True

# Label column bold
for ri in range(1, rows):
    for p in tbl.cell(ri, 0).text_frame.paragraphs:
        p.font.bold = True

# ======================================================================
# Slide 4 — Drawdown Analysis
# ======================================================================
dd = _dark_slide(prs, "Drawdown Analysis")

# Context text
ctx = dd.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(8.8), Inches(0.7))
ctf = ctx.text_frame
ctf.word_wrap = True
ctf.text = (
    "The chart below shows peak-to-trough drawdowns since inception. "
    "Maximum drawdown was \u221212.4% (Mar\u2013Jun 2022) vs. \u221218.9% for MSCI World."
)
ctf.paragraphs[0].font.size = Pt(12)
ctf.paragraphs[0].font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)

# Drawdown line chart
dd_data = CategoryChartData()
dd_data.categories = [
    "Jan 20", "Jul 20", "Jan 21", "Jul 21",
    "Jan 22", "Jul 22", "Jan 23", "Jul 23",
    "Jan 24", "Jul 24", "Jan 25", "Dec 25",
]
dd_data.add_series("Apex Global Equity", (
    0, -2.1, 0, -1.5,
    -3.8, -12.4, -4.2, 0,
    -1.9, -5.3, -2.0, 0,
))
dd_data.add_series("MSCI World (TR)", (
    0, -4.5, 0, -3.2,
    -8.1, -18.9, -9.7, -1.3,
    -3.6, -8.0, -4.1, 0,
))

dd_frame = dd.shapes.add_chart(
    XL_CHART_TYPE.LINE,
    Inches(0.6), Inches(2.2), Inches(8.8), Inches(4.2),
    dd_data,
)
dd_chart = dd_frame.chart
dd_chart.has_legend = True
dd_chart.legend.position = XL_LEGEND_POSITION.BOTTOM
dd_chart.legend.include_in_layout = False
dd_chart.chart_title.has_text_frame = True
dd_chart.chart_title.text_frame.text = "Drawdown (%)"
dd_chart.chart_title.text_frame.paragraphs[0].font.size = Pt(11)

# ======================================================================
# Slide 5 — Contact
# ======================================================================
contact = _dark_slide(prs, "Contact Us")

contacts = [
    {
        "name": "Sarah Chen, CFA",
        "title": "Managing Partner & Chief Investment Officer",
        "email": "schen@apexcapital.com",
        "phone": "+1 (212) 555-0142",
    },
    {
        "name": "James Whitfield",
        "title": "Head of Investor Relations",
        "email": "jwhitfield@apexcapital.com",
        "phone": "+1 (212) 555-0198",
    },
    {
        "name": "Maria Vasquez",
        "title": "Director of Client Services",
        "email": "mvasquez@apexcapital.com",
        "phone": "+44 20 7946 0321",
    },
]

y_offset = 1.6
for person in contacts:
    tb = contact.shapes.add_textbox(
        Inches(1.2), Inches(y_offset), Inches(7.6), Inches(1.2),
    )
    tf = tb.text_frame
    tf.word_wrap = True

    # Name
    tf.text = person["name"]
    tf.paragraphs[0].font.size = Pt(18)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = WHITE

    # Title
    p1 = tf.add_paragraph()
    p1.text = person["title"]
    p1.font.size = Pt(12)
    p1.font.color.rgb = ACCENT_BLUE
    p1.space_before = Pt(2)

    # Email & phone
    p2 = tf.add_paragraph()
    p2.text = f"{person['email']}  |  {person['phone']}"
    p2.font.size = Pt(11)
    p2.font.color.rgb = RGBColor(0x99, 0xAA, 0xBB)
    p2.space_before = Pt(2)

    y_offset += 1.6

# Office address
addr = contact.shapes.add_textbox(Inches(1.2), Inches(6.2), Inches(7.6), Inches(0.5))
atf = addr.text_frame
atf.text = "Apex Capital Partners  \u2022  450 Park Avenue, 28th Floor, New York, NY 10022"
atf.paragraphs[0].font.size = Pt(10)
atf.paragraphs[0].font.color.rgb = RGBColor(0x66, 0x77, 0x88)
atf.paragraphs[0].alignment = PP_ALIGN.CENTER

# ======================================================================
# Save
# ======================================================================
prs.save("investor_presentation.pptx")
print("Saved investor_presentation.pptx")
