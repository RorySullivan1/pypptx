"""Bar and line charts with data labels and legends.

Creates a clustered bar chart and a line chart on separate slides.
Saves to add_chart.pptx.
"""

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

prs = Presentation()

# --- Slide 1: Clustered Bar Chart ---
slide1 = prs.slides.add_slide(prs.slide_layouts[5])

chart_data = CategoryChartData()
chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
chart_data.add_series("East", (19.2, 21.4, 16.7, 23.1))
chart_data.add_series("West", (14.5, 18.3, 22.9, 17.6))

graphic_frame = slide1.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1), Inches(8), Inches(5),
    chart_data,
)
chart = graphic_frame.chart
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
chart.chart_title.has_text_frame = True
chart.chart_title.text_frame.text = "Revenue by Region"

# --- Slide 2: Line Chart ---
slide2 = prs.slides.add_slide(prs.slide_layouts[5])

line_data = CategoryChartData()
line_data.categories = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
line_data.add_series("Users", (120, 180, 240, 350, 410, 520))
line_data.add_series("Sessions", (300, 420, 510, 680, 790, 950))

line_frame = slide2.shapes.add_chart(
    XL_CHART_TYPE.LINE_MARKERS,
    Inches(1), Inches(1), Inches(8), Inches(5),
    line_data,
)
line_chart = line_frame.chart
line_chart.has_legend = True
line_chart.chart_title.has_text_frame = True
line_chart.chart_title.text_frame.text = "Monthly Growth"

prs.save("add_chart.pptx")
print("Saved add_chart.pptx")
