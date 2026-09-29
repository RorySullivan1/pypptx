"""3D charts: the camera (`view_3d`), the floor and walls, and three 3D plot types.

Slide 1 is a "true 3D" column chart -- every series in its own row along a depth axis --
with a perspective camera and a coloured floor and back wall. Slide 2 is a 3D clustered bar
chart keeping right-angle axes, the oblique look PowerPoint gives its clustered and stacked
3D types. Slide 3 is an exploded 3D pie tilted towards the viewer.

A new 3D chart already carries PowerPoint's default camera; the settings below change it.
Values the schema allows: `rot_x` -90..90, `rot_y` 0..360, `perspective` 0..240,
`depth_percent` 20..2000, `height_percent` 5..500.

The deck is reopened after saving and every value is asserted against what was written, so
running this script is a round-trip check and not just a drawing exercise.

Demonstrates: feature-manifest.md "Charts" (3D variants, 3D view, 3D surfaces).
Saves to chart_3d_view.pptx.
"""

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.chart.plot import Bar3DPlot, Pie3DPlot
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

FILENAME = "chart_3d_view.pptx"
FLOOR = RGBColor(0xD9, 0xD9, 0xD9)
BACK_WALL = RGBColor(0xEA, 0xF1, 0xFB)

prs = Presentation()
blank = prs.slide_layouts[6]

regional = CategoryChartData()
regional.categories = ["North", "South", "East", "West"]
regional.add_series("2024", (14.2, 9.8, 11.5, 7.9))
regional.add_series("2025", (16.1, 10.4, 12.9, 9.3))
regional.add_series("2026", (18.4, 11.7, 13.6, 10.8))

# --- Slide 1: true 3D column, perspective camera, formatted floor and back wall -----------
chart = prs.slides.add_slide(blank).shapes.add_chart(
    XL_CHART_TYPE.THREE_D_COLUMN, Inches(0.5), Inches(0.5), Inches(9), Inches(6.5), regional
).chart
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM

view = chart.view_3d
view.rot_x = 20  # tilt: how far the camera looks down on the plot
view.rot_y = 30  # turn: how far the plot is spun about its vertical axis
view.perspective = 45  # field of view; only applies while right_angle_axes is False
view.depth_percent = 150  # depth of the plot box relative to its width
view.right_angle_axes = False

chart.floor.format.fill.solid()
chart.floor.format.fill.fore_color.rgb = FLOOR
chart.back_wall.format.fill.solid()
chart.back_wall.format.fill.fore_color.rgb = BACK_WALL
chart.side_wall.format.fill.background()  # no fill: leave the side open

# --- Slide 2: 3D clustered bar, right-angle axes, taller plot box --------------------------
chart = prs.slides.add_slide(blank).shapes.add_chart(
    XL_CHART_TYPE.THREE_D_BAR_CLUSTERED, Inches(0.5), Inches(0.5), Inches(9), Inches(6.5), regional
).chart
chart.view_3d.rot_x = 10
chart.view_3d.height_percent = 150  # plot-box height relative to its width
chart.plots[0].gap_width = 80

# --- Slide 3: exploded 3D pie, tilted ------------------------------------------------------
share = CategoryChartData()
share.categories = ["Hardware", "Software", "Services"]
share.add_series("Revenue share", (0.45, 0.35, 0.20))

chart = prs.slides.add_slide(blank).shapes.add_chart(
    XL_CHART_TYPE.THREE_D_PIE_EXPLODED, Inches(1.5), Inches(0.5), Inches(7), Inches(6.5), share
).chart
chart.view_3d.rot_x = 50  # a steeper tilt shows more of the pie's edge
chart.view_3d.rot_y = 90  # turn the first slice round to the side
chart.font.size = Pt(14)

prs.save(FILENAME)
print(f"Saved {FILENAME}")

# --- Round trip ----------------------------------------------------------------------------
# Read every value back out of the saved file rather than off the objects above, which would
# only prove the setters worked in memory.
reopened = Presentation(FILENAME)
column, bar, pie = (slide.shapes[0].chart for slide in reopened.slides)

assert column.chart_type == XL_CHART_TYPE.THREE_D_COLUMN, column.chart_type
assert isinstance(column.plots[0], Bar3DPlot)
assert [s.name for s in column.plots[0].series] == ["2024", "2025", "2026"]
assert tuple(column.plots[0].series[2].values) == (18.4, 11.7, 13.6, 10.8)
view = column.view_3d
assert (view.rot_x, view.rot_y, view.perspective) == (20, 30, 45)
assert view.depth_percent == 150, view.depth_percent
assert view.right_angle_axes is False
assert column.floor.format.fill.fore_color.rgb == FLOOR
assert column.back_wall.format.fill.fore_color.rgb == BACK_WALL

assert bar.chart_type == XL_CHART_TYPE.THREE_D_BAR_CLUSTERED, bar.chart_type
assert bar.view_3d.rot_x == 10
assert bar.view_3d.height_percent == 150, bar.view_3d.height_percent
assert bar.view_3d.right_angle_axes is True
assert bar.plots[0].gap_width == 80

assert pie.chart_type == XL_CHART_TYPE.THREE_D_PIE_EXPLODED, pie.chart_type
assert isinstance(pie.plots[0], Pie3DPlot)
assert tuple(pie.plots[0].categories) == ("Hardware", "Software", "Services")
assert (pie.view_3d.rot_x, pie.view_3d.rot_y) == (50, 90)

print("Round trip verified: 3 charts, camera, floor and walls")
