"""Combo chart with a secondary value axis.

Builds a clustered column chart of Revenue and Profit ($) and overlays a
line plot of Margin (%) on a secondary value axis — so two series with
very different scales can be read on the same chart.

Saves to add_chart_secondary_axis.pptx.
"""

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (
    XL_CHART_TYPE,
    XL_LABEL_POSITION,
    XL_LEGEND_POSITION,
    XL_MARKER_STYLE,
)
from pptx.util import Inches, Pt

# -- Palette --
COLUMN_BLUE = RGBColor(0x1B, 0x6B, 0xB5)
COLUMN_NAVY = RGBColor(0x0B, 0x1D, 0x3A)
MARGIN_GOLD = RGBColor(0xD4, 0xA0, 0x1E)


def move_series_to_overlay_plot(chart, src_plot_idx, ser_idx, dst_plot_idx):
    """Re-parent a `<c:ser>` element from one xChart into another.

    `chart.add_plot(...)` creates an *empty* overlay plot; it has no built-in
    way to populate itself. The supported workflow is to build the primary
    chart with every series, then move the ones that belong on the
    secondary axis into the overlay plot via this OXML surgery.

    After the move, the series is re-typed automatically: a `<c:ser>` moved
    into a `<c:lineChart>` reports as a `LineSeries` on `chart.series`,
    because `PlotFactory` dispatches off the parent xChart tag.
    """
    plotArea = chart._chartSpace.chart.plotArea
    xCharts = list(plotArea.iter_xCharts())
    src = xCharts[src_plot_idx]
    dst = xCharts[dst_plot_idx]
    ser = src.sers[ser_idx]
    src.remove(ser)
    # Insert before the destination's first <c:axId> child to keep schema order.
    dst.xpath("c:axId")[0].addprevious(ser)


prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

# --- Build the primary chart with ALL series (columns) ---
data = CategoryChartData()
data.categories = ["Q1", "Q2", "Q3", "Q4"]
data.add_series("Revenue ($M)", (100, 120, 140, 180))
data.add_series("Profit ($M)", (20, 28, 34, 50))
data.add_series("Margin (%)", (20, 23, 24, 28))

graphic_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1), Inches(8), Inches(5),
    data,
)
chart = graphic_frame.chart

# --- Add an empty line overlay plot bound to a new secondary value axis ---
chart.add_plot("line", use_secondary_axis=True)

# --- Move "Margin (%)" out of the column plot and into the line plot ---
# Column xChart is plot 0 (built by add_chart). The line xChart is plot 1.
# "Margin (%)" is the 3rd column series, i.e. index 2 in the column plot.
move_series_to_overlay_plot(chart, src_plot_idx=0, ser_idx=2, dst_plot_idx=1)

# --- Baseline chart chrome ---
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
chart.chart_title.has_text_frame = True
chart.chart_title.text_frame.text = "Quarterly Performance"

# Style the column series for a clean baseline.
chart.series[0].format.fill.solid()
chart.series[0].format.fill.fore_color.rgb = COLUMN_BLUE
chart.series[1].format.fill.solid()
chart.series[1].format.fill.fore_color.rgb = COLUMN_NAVY


def style_secondary_axis_series(line_series, secondary_axis):
    """TODO: style the Margin (%) line so its scale is unambiguous.

    Context
    -------
    Two of the three series (Revenue, Profit) are in $M, in the 0–200 range.
    The third (Margin) is a percentage, 20–28. Without visual cues, a viewer
    looking at the line cannot tell which axis it reads against — and the
    secondary axis on the right will just look like a duplicate.

    The choices you make here decide the chart's readability. Trade-offs:

      - Bold contrasting color + matching axis title color makes the
        line ↔ right-axis pairing obvious, at the cost of visual noise.
      - Subtle gold line + small italic axis label keeps the columns
        dominant and treats margin as supporting context.
      - Markers + data labels at each point emphasize the *trend* of
        margin movement, useful if the story is "margin is expanding."

    Parameters
    ----------
    line_series : LineSeries
        The Margin (%) series, already re-parented into the overlay plot.
        Useful surface: `line_series.format.line.color.rgb`,
        `line_series.format.line.width`, `line_series.marker.style`,
        `line_series.marker.size`, `line_series.smooth`.
    secondary_axis : ValueAxis
        The right-side value axis. Useful surface: `secondary_axis.has_title`,
        `secondary_axis.axis_title.text_frame.text`, and per-run font
        properties via `.text_frame.paragraphs[0].runs[0].font`.

    Write 5–10 lines that express *one* of the design choices above.
    """
    # Option 3: emphasize the margin *trend* with markers and per-point labels.
    line_series.format.line.color.rgb = MARGIN_GOLD
    line_series.format.line.width = Pt(2.75)
    line_series.marker.style = XL_MARKER_STYLE.CIRCLE
    line_series.marker.size = 7
    labels = line_series.data_labels
    labels.show_value = True
    labels.position = XL_LABEL_POSITION.ABOVE
    labels.number_format = '0"%"'
    secondary_axis.has_title = True
    secondary_axis.axis_title.text_frame.text = "Margin (%)"


# Wire up the contribution point.
margin_line_series = chart.series[-1]
secondary_axis = chart.secondary_value_axis
style_secondary_axis_series(margin_line_series, secondary_axis)

prs.save("add_chart_secondary_axis.pptx")
print("Saved add_chart_secondary_axis.pptx")
