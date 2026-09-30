"""Plot-related objects.

A plot is known as a chart group in the MS API. A chart can have more than one plot overlayed on
each other, such as a line plot layered over a bar plot.
"""

from __future__ import annotations

from pptx.chart.category import Categories
from pptx.chart.chartlines import DropLines, HiLowLines, SeriesLines, UpDownBars
from pptx.chart.datalabel import DataLabels
from pptx.exc import ChartError
from pptx.chart.series import SeriesCollection
from pptx.enum.chart import XL_BAR_DIRECTION, XL_CHART_TYPE as XL, XL_GROUPING
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import ST_BarDir, ST_Grouping
from pptx.util import lazyproperty


class _BasePlot:
    """
    A distinct plot that appears in the plot area of a chart. A chart may
    have more than one plot, in which case they appear as superimposed
    layers, such as a line plot appearing on top of a bar chart.
    """

    def __init__(self, xChart, chart):
        super(_BasePlot, self).__init__()
        self._element = xChart
        self._chart = chart

    @property
    def axis_ids(self):
        """Tuple of integer axis IDs this plot references.

        Each xChart element contains ``c:axId`` children that identify which
        axes it uses. Returns an empty tuple for chart types that have no axes
        (e.g. pie charts).
        """
        return tuple(self._element.axId_vals)

    @lazyproperty
    def categories(self):
        """
        Returns a |category.Categories| sequence object containing
        a |category.Category| object for each of the category labels
        associated with this plot. The |category.Category| class derives from
        ``str``, so the returned value can be treated as a simple sequence of
        strings for the common case where all you need is the labels in the
        order they appear on the chart. |category.Categories| provides
        additional properties for dealing with hierarchical categories when
        required.
        """
        return Categories(self._element)

    @property
    def chart(self):
        """
        The |Chart| object containing this plot.
        """
        return self._chart

    @property
    def grouping(self) -> XL_GROUPING | None:
        """|XL_GROUPING| member describing how this plot's series are grouped.

        Returns |None| when grouping is not applicable to the chart type
        (e.g. pie, scatter, doughnut).
        """
        grouping_val = getattr(self._element, "grouping_val", None)
        if grouping_val is None:
            return None
        return XL_GROUPING.from_xml(grouping_val)

    @property
    def data_labels(self):
        """
        |DataLabels| instance providing properties and methods on the
        collection of data labels associated with this plot.
        """
        dLbls = self._element.dLbls
        if dLbls is None:
            raise ChartError("plot has no data labels, set has_data_labels = True first")
        return DataLabels(dLbls)

    @property
    def has_data_labels(self):
        """
        Read/write boolean, |True| if the series has data labels. Assigning
        |True| causes data labels to be added to the plot. Assigning False
        removes any existing data labels.
        """
        return self._element.dLbls is not None

    @has_data_labels.setter
    def has_data_labels(self, value):
        """
        Add, remove, or leave alone the ``<c:dLbls>`` child element depending
        on current state and assigned *value*. If *value* is |True| and no
        ``<c:dLbls>`` element is present, a new default element is added with
        default child elements and settings. When |False|, any existing dLbls
        element is removed.
        """
        if bool(value) is False:
            self._element._remove_dLbls()
        else:
            if self._element.dLbls is None:
                dLbls = self._element._add_dLbls()
                dLbls.showVal.val = True

    @lazyproperty
    def series(self):
        """
        A sequence of |Series| objects representing the series in this plot,
        in the order they appear in the plot.
        """
        return SeriesCollection(self._element)

    @property
    def vary_by_categories(self):
        """
        Read/write boolean value specifying whether to use a different color
        for each of the points in this plot. Only effective when there is
        a single series; PowerPoint automatically varies color by series when
        more than one series is present.
        """
        varyColors = self._element.varyColors
        if varyColors is None:
            return True
        return varyColors.val

    @vary_by_categories.setter
    def vary_by_categories(self, value):
        self._element.get_or_add_varyColors().val = bool(value)


class AreaPlot(_BasePlot):
    """An area plot."""

    @property
    def has_drop_lines(self):
        """Read/write boolean. |True| if drop lines are displayed."""
        return self._element.dropLines is not None

    @has_drop_lines.setter
    def has_drop_lines(self, value):
        if bool(value):
            self._element.get_or_add_dropLines()
        else:
            self._element._remove_dropLines()

    @property
    def drop_lines(self):
        """A |DropLines| object for formatting drop lines.

        Raises if no drop lines present; use :attr:`has_drop_lines` first.
        """
        dropLines = self._element.dropLines
        if dropLines is None:
            raise ChartError("plot has no drop lines")
        return DropLines(dropLines)


class Area3DPlot(_BasePlot):
    """A 3-dimensional area plot."""


class Bar3DPlot(_BasePlot):
    """A 3-dimensional bar chart-style plot."""

    @property
    def gap_width(self):
        """Width of the gap between the bars of neighbouring categories, as a percentage.

        An integer percentage of the bar width, from `c:gapWidth`; 150 (1.5 bar widths) when
        the element is absent, PowerPoint's default for a new chart. Read/write.
        """
        gapWidth = self._element.gapWidth
        if gapWidth is None:
            return 150
        return gapWidth.val

    @gap_width.setter
    def gap_width(self, value):
        gapWidth = self._element.get_or_add_gapWidth()
        gapWidth.val = value


class BarPlot(_BasePlot):
    """
    A bar chart-style plot.
    """

    @property
    def bar_direction(self) -> XL_BAR_DIRECTION:
        """|XL_BAR_DIRECTION| member describing this plot's bar orientation.

        ``XL_BAR_DIRECTION.BAR`` for horizontal bars,
        ``XL_BAR_DIRECTION.COLUMN`` for vertical columns.
        """
        return XL_BAR_DIRECTION.from_xml(self._element.barDir.val)

    @property
    def gap_width(self):
        """
        Width of gap between bar(s) of each category, as an integer
        percentage of the bar width. The default value for a new bar chart is
        150, representing 150% or 1.5 times the width of a single bar.
        """
        gapWidth = self._element.gapWidth
        if gapWidth is None:
            return 150
        return gapWidth.val

    @gap_width.setter
    def gap_width(self, value):
        gapWidth = self._element.get_or_add_gapWidth()
        gapWidth.val = value

    @property
    def overlap(self):
        """
        Read/write int value in range -100..100 specifying a percentage of
        the bar width by which to overlap adjacent bars in a multi-series bar
        chart. Default is 0. A setting of -100 creates a gap of a full bar
        width and a setting of 100 causes all the bars in a category to be
        superimposed. A stacked bar plot has overlap of 100 by default.
        """
        overlap = self._element.overlap
        if overlap is None:
            return 0
        return overlap.val

    @overlap.setter
    def overlap(self, value):
        """
        Set the value of the ``<c:overlap>`` child element to *int_value*,
        or remove the overlap element if *int_value* is 0.
        """
        if value == 0:
            self._element._remove_overlap()
            return
        self._element.get_or_add_overlap().val = value

    @property
    def has_series_lines(self):
        """Read/write boolean. |True| if series connector lines are displayed."""
        return self._element.serLines is not None

    @has_series_lines.setter
    def has_series_lines(self, value):
        if bool(value):
            self._element.get_or_add_serLines()
        else:
            self._element._remove_serLines()

    @property
    def series_lines(self):
        """A |SeriesLines| object. Raises if not present."""
        serLines = self._element.serLines
        if serLines is None:
            raise ChartError("plot has no series lines")
        return SeriesLines(serLines)


class BubblePlot(_BasePlot):
    """
    A bubble chart plot.
    """

    @property
    def bubble_scale(self):
        """
        An integer between 0 and 300 inclusive indicating the percentage of
        the default size at which bubbles should be displayed. Assigning
        |None| produces the same behavior as assigning `100`.
        """
        bubbleScale = self._element.bubbleScale
        if bubbleScale is None:
            return 100
        return bubbleScale.val

    @bubble_scale.setter
    def bubble_scale(self, value):
        bubbleChart = self._element
        bubbleChart._remove_bubbleScale()
        if value is None:
            return
        bubbleScale = bubbleChart._add_bubbleScale()
        bubbleScale.val = value


class DoughnutPlot(_BasePlot):
    """A doughnut plot."""

    @property
    def first_slice_angle(self):
        """Read/write integer. Angle in degrees of the first slice.

        Default is 0 when not explicitly set.
        """
        firstSliceAng = self._element.firstSliceAng
        if firstSliceAng is None:
            return 0
        return firstSliceAng.val

    @first_slice_angle.setter
    def first_slice_angle(self, value):
        self._element.get_or_add_firstSliceAng().val = value

    @property
    def hole_size(self):
        """Read/write integer. Size of the hole as a percentage (1-90).

        Default is 10 when not explicitly set.
        """
        holeSize = self._element.holeSize
        if holeSize is None:
            return 10
        return holeSize.val

    @hole_size.setter
    def hole_size(self, value):
        self._element.get_or_add_holeSize().val = value


class Line3DPlot(_BasePlot):
    """A 3-dimensional line plot."""


class LinePlot(_BasePlot):
    """A line chart-style plot."""

    @property
    def has_drop_lines(self):
        """Read/write boolean. |True| if drop lines are displayed."""
        return self._element.dropLines is not None

    @has_drop_lines.setter
    def has_drop_lines(self, value):
        if bool(value):
            self._element.get_or_add_dropLines()
        else:
            self._element._remove_dropLines()

    @property
    def drop_lines(self):
        """A |DropLines| object. Raises if not present."""
        dropLines = self._element.dropLines
        if dropLines is None:
            raise ChartError("plot has no drop lines")
        return DropLines(dropLines)

    @property
    def has_hi_low_lines(self):
        """Read/write boolean. |True| if high-low lines are displayed."""
        return self._element.hiLowLines is not None

    @has_hi_low_lines.setter
    def has_hi_low_lines(self, value):
        if bool(value):
            self._element.get_or_add_hiLowLines()
        else:
            self._element._remove_hiLowLines()

    @property
    def hi_low_lines(self):
        """A |HiLowLines| object. Raises if not present."""
        hiLowLines = self._element.hiLowLines
        if hiLowLines is None:
            raise ChartError("plot has no high-low lines")
        return HiLowLines(hiLowLines)

    @property
    def has_up_down_bars(self):
        """Read/write boolean. |True| if up/down bars are displayed."""
        return self._element.upDownBars is not None

    @has_up_down_bars.setter
    def has_up_down_bars(self, value):
        if bool(value):
            self._element.get_or_add_upDownBars()
        else:
            self._element._remove_upDownBars()

    @property
    def up_down_bars(self):
        """An |UpDownBars| object. Raises if not present."""
        upDownBars = self._element.upDownBars
        if upDownBars is None:
            raise ChartError("plot has no up/down bars")
        return UpDownBars(upDownBars)


class OfPiePlot(_BasePlot):
    """A pie-of-pie or bar-of-pie plot."""

    @property
    def gap_width(self):
        """Read/write integer. Gap width between primary and secondary pie/bar.

        Default is 150 when not explicitly set.
        """
        gapWidth = self._element.gapWidth
        if gapWidth is None:
            return 150
        return gapWidth.val

    @gap_width.setter
    def gap_width(self, value):
        self._element.get_or_add_gapWidth().val = value

    @property
    def has_series_lines(self):
        """Read/write boolean. |True| if series connector lines are displayed."""
        return self._element.serLines is not None

    @has_series_lines.setter
    def has_series_lines(self, value):
        if bool(value):
            self._element.get_or_add_serLines()
        else:
            self._element._remove_serLines()

    @property
    def second_pie_size(self):
        """Read/write integer. Size of the secondary pie as percentage (5-200).

        Default is 75 when not explicitly set.
        """
        secondPieSize = self._element.secondPieSize
        if secondPieSize is None:
            return 75
        return secondPieSize.val

    @second_pie_size.setter
    def second_pie_size(self, value):
        self._element.get_or_add_secondPieSize().val = value

    @property
    def series_lines(self):
        """A |SeriesLines| object. Raises if not present."""
        serLines = self._element.serLines
        if serLines is None:
            raise ChartError("plot has no series lines")
        return SeriesLines(serLines)

    @property
    def split_pos(self):
        """Read/write float. The threshold value for splitting data points.

        Returns None if not explicitly set.
        """
        splitPos = self._element.splitPos
        if splitPos is None:
            return None
        return splitPos.val

    @split_pos.setter
    def split_pos(self, value):
        if value is None:
            self._element._remove_splitPos()
            return
        self._element.get_or_add_splitPos().val = value

    @property
    def split_type(self):
        """Read/write string. How data is split between the two pies.

        One of 'auto', 'cust', 'percent', 'pos', 'val'. Default is 'auto'.
        """
        splitType = self._element.splitType
        if splitType is None:
            return "auto"
        return splitType.val

    @split_type.setter
    def split_type(self, value):
        self._element.get_or_add_splitType().val = value


class Pie3DPlot(_BasePlot):
    """A 3-dimensional pie plot."""


class PiePlot(_BasePlot):
    """A pie chart-style plot."""

    @property
    def first_slice_angle(self):
        """Read/write integer. Angle in degrees of the first slice.

        Default is 0 when not explicitly set.
        """
        firstSliceAng = self._element.firstSliceAng
        if firstSliceAng is None:
            return 0
        return firstSliceAng.val

    @first_slice_angle.setter
    def first_slice_angle(self, value):
        self._element.get_or_add_firstSliceAng().val = value


class RadarPlot(_BasePlot):
    """
    A radar-style plot.
    """


class StockPlot(_BasePlot):
    """A stock (OHLC/HLC) chart plot."""

    @property
    def has_drop_lines(self):
        """Read/write boolean. |True| if drop lines are displayed."""
        return self._element.dropLines is not None

    @has_drop_lines.setter
    def has_drop_lines(self, value):
        if bool(value):
            self._element.get_or_add_dropLines()
        else:
            self._element._remove_dropLines()

    @property
    def drop_lines(self):
        """A |DropLines| object. Raises if not present."""
        dropLines = self._element.dropLines
        if dropLines is None:
            raise ChartError("plot has no drop lines")
        return DropLines(dropLines)

    @property
    def has_hi_low_lines(self):
        """Read/write boolean. |True| if high-low lines are displayed."""
        return self._element.hiLowLines is not None

    @has_hi_low_lines.setter
    def has_hi_low_lines(self, value):
        if bool(value):
            self._element.get_or_add_hiLowLines()
        else:
            self._element._remove_hiLowLines()

    @property
    def hi_low_lines(self):
        """A |HiLowLines| object. Raises if not present."""
        hiLowLines = self._element.hiLowLines
        if hiLowLines is None:
            raise ChartError("plot has no high-low lines")
        return HiLowLines(hiLowLines)

    @property
    def has_up_down_bars(self):
        """Read/write boolean. |True| if up/down bars are displayed."""
        return self._element.upDownBars is not None

    @has_up_down_bars.setter
    def has_up_down_bars(self, value):
        if bool(value):
            self._element.get_or_add_upDownBars()
        else:
            self._element._remove_upDownBars()

    @property
    def up_down_bars(self):
        """An |UpDownBars| object. Raises if not present."""
        upDownBars = self._element.upDownBars
        if upDownBars is None:
            raise ChartError("plot has no up/down bars")
        return UpDownBars(upDownBars)


class Surface3DPlot(_BasePlot):
    """A 3-dimensional surface plot."""


class SurfacePlot(_BasePlot):
    """A surface (top-view) plot."""


class XyPlot(_BasePlot):
    """
    An XY (scatter) plot.
    """


def PlotFactory(xChart, chart):
    """
    Return an instance of the appropriate subclass of _BasePlot based on the
    tagname of *xChart*. *chart* is the |Chart| the plot belongs to.
    """
    try:
        PlotCls = {
            qn("c:areaChart"): AreaPlot,
            qn("c:area3DChart"): Area3DPlot,
            qn("c:bar3DChart"): Bar3DPlot,
            qn("c:barChart"): BarPlot,
            qn("c:bubbleChart"): BubblePlot,
            qn("c:doughnutChart"): DoughnutPlot,
            qn("c:line3DChart"): Line3DPlot,
            qn("c:lineChart"): LinePlot,
            qn("c:ofPieChart"): OfPiePlot,
            qn("c:pie3DChart"): Pie3DPlot,
            qn("c:pieChart"): PiePlot,
            qn("c:radarChart"): RadarPlot,
            qn("c:scatterChart"): XyPlot,
            qn("c:stockChart"): StockPlot,
            qn("c:surface3DChart"): Surface3DPlot,
            qn("c:surfaceChart"): SurfacePlot,
        }[xChart.tag]
    except KeyError:
        raise ChartError("unsupported plot type %s" % xChart.tag)

    return PlotCls(xChart, chart)


class PlotTypeInspector:
    """
    "One-shot" service object that knows how to identify the type of a plot
    as a member of the XL_CHART_TYPE enumeration.
    """

    @classmethod
    def chart_type(cls, plot):
        """
        Return the member of :ref:`XlChartType` that corresponds to the chart
        type of *plot*.
        """
        try:
            chart_type_method = {
                "AreaPlot": cls._differentiate_area_chart_type,
                "Area3DPlot": cls._differentiate_area_3d_chart_type,
                "Bar3DPlot": cls._differentiate_bar_3d_chart_type,
                "BarPlot": cls._differentiate_bar_chart_type,
                "BubblePlot": cls._differentiate_bubble_chart_type,
                "DoughnutPlot": cls._differentiate_doughnut_chart_type,
                "Line3DPlot": cls._differentiate_line_3d_chart_type,
                "LinePlot": cls._differentiate_line_chart_type,
                "OfPiePlot": cls._differentiate_of_pie_chart_type,
                "Pie3DPlot": cls._differentiate_pie_3d_chart_type,
                "PiePlot": cls._differentiate_pie_chart_type,
                "RadarPlot": cls._differentiate_radar_chart_type,
                "StockPlot": cls._differentiate_stock_chart_type,
                "Surface3DPlot": cls._differentiate_surface_3d_chart_type,
                "SurfacePlot": cls._differentiate_surface_chart_type,
                "XyPlot": cls._differentiate_xy_chart_type,
            }[plot.__class__.__name__]
        except KeyError:
            raise NotImplementedError(
                "chart_type() not implemented for %s" % plot.__class__.__name__
            )
        return chart_type_method(plot)

    @classmethod
    def _differentiate_area_3d_chart_type(cls, plot):
        return {
            ST_Grouping.STANDARD: XL.THREE_D_AREA,
            ST_Grouping.STACKED: XL.THREE_D_AREA_STACKED,
            ST_Grouping.PERCENT_STACKED: XL.THREE_D_AREA_STACKED_100,
        }[plot._element.grouping_val]

    @classmethod
    def _differentiate_bar_3d_chart_type(cls, plot):
        barChart = plot._element
        if barChart.barDir.val == ST_BarDir.BAR:
            return {
                ST_Grouping.CLUSTERED: XL.THREE_D_BAR_CLUSTERED,
                ST_Grouping.STACKED: XL.THREE_D_BAR_STACKED,
                ST_Grouping.PERCENT_STACKED: XL.THREE_D_BAR_STACKED_100,
            }[barChart.grouping_val]
        return {
            ST_Grouping.STANDARD: XL.THREE_D_COLUMN,
            ST_Grouping.CLUSTERED: XL.THREE_D_COLUMN_CLUSTERED,
            ST_Grouping.STACKED: XL.THREE_D_COLUMN_STACKED,
            ST_Grouping.PERCENT_STACKED: XL.THREE_D_COLUMN_STACKED_100,
        }[barChart.grouping_val]

    @classmethod
    def _differentiate_area_chart_type(cls, plot):
        return {
            ST_Grouping.STANDARD: XL.AREA,
            ST_Grouping.STACKED: XL.AREA_STACKED,
            ST_Grouping.PERCENT_STACKED: XL.AREA_STACKED_100,
        }[plot._element.grouping_val]

    @classmethod
    def _differentiate_bar_chart_type(cls, plot):
        barChart = plot._element
        if barChart.barDir.val == ST_BarDir.BAR:
            return {
                ST_Grouping.CLUSTERED: XL.BAR_CLUSTERED,
                ST_Grouping.STACKED: XL.BAR_STACKED,
                ST_Grouping.PERCENT_STACKED: XL.BAR_STACKED_100,
            }[barChart.grouping_val]
        if barChart.barDir.val == ST_BarDir.COL:
            return {
                ST_Grouping.CLUSTERED: XL.COLUMN_CLUSTERED,
                ST_Grouping.STACKED: XL.COLUMN_STACKED,
                ST_Grouping.PERCENT_STACKED: XL.COLUMN_STACKED_100,
            }[barChart.grouping_val]
        raise ChartError("invalid barChart.barDir value '%s'" % barChart.barDir.val)

    @classmethod
    def _differentiate_bubble_chart_type(cls, plot):
        def first_bubble3D(bubbleChart):
            results = bubbleChart.xpath("c:ser/c:bubble3D")
            return results[0] if results else None

        bubbleChart = plot._element
        bubble3D = first_bubble3D(bubbleChart)

        if bubble3D is None:
            return XL.BUBBLE
        if bubble3D.val:
            return XL.BUBBLE_THREE_D_EFFECT
        return XL.BUBBLE

    @classmethod
    def _differentiate_doughnut_chart_type(cls, plot):
        doughnutChart = plot._element
        explosion = doughnutChart.xpath("./c:ser/c:explosion")
        return XL.DOUGHNUT_EXPLODED if explosion else XL.DOUGHNUT

    @classmethod
    def _differentiate_line_chart_type(cls, plot):
        lineChart = plot._element

        def has_line_markers():
            matches = lineChart.xpath('c:ser/c:marker/c:symbol[@val="none"]')
            if matches:
                return False
            return True

        if has_line_markers():
            return {
                ST_Grouping.STANDARD: XL.LINE_MARKERS,
                ST_Grouping.STACKED: XL.LINE_MARKERS_STACKED,
                ST_Grouping.PERCENT_STACKED: XL.LINE_MARKERS_STACKED_100,
            }[plot._element.grouping_val]
        else:
            return {
                ST_Grouping.STANDARD: XL.LINE,
                ST_Grouping.STACKED: XL.LINE_STACKED,
                ST_Grouping.PERCENT_STACKED: XL.LINE_STACKED_100,
            }[plot._element.grouping_val]

    @classmethod
    def _differentiate_pie_chart_type(cls, plot):
        pieChart = plot._element
        explosion = pieChart.xpath("./c:ser/c:explosion")
        return XL.PIE_EXPLODED if explosion else XL.PIE

    @classmethod
    def _differentiate_radar_chart_type(cls, plot):
        radarChart = plot._element
        radar_style = radarChart.xpath("c:radarStyle")[0].get("val")

        def noMarkers():
            matches = radarChart.xpath("c:ser/c:marker/c:symbol")
            if matches and matches[0].get("val") == "none":
                return True
            return False

        if radar_style is None:
            return XL.RADAR
        if radar_style == "filled":
            return XL.RADAR_FILLED
        if noMarkers():
            return XL.RADAR
        return XL.RADAR_MARKERS

    @classmethod
    def _differentiate_xy_chart_type(cls, plot):
        scatterChart = plot._element

        def noLine():
            return bool(scatterChart.xpath("c:ser/c:spPr/a:ln/a:noFill"))

        def noMarkers():
            symbols = scatterChart.xpath("c:ser/c:marker/c:symbol")
            if symbols and symbols[0].get("val") == "none":
                return True
            return False

        scatter_style = scatterChart.xpath("c:scatterStyle")[0].get("val")

        if scatter_style == "lineMarker":
            if noLine():
                return XL.XY_SCATTER
            if noMarkers():
                return XL.XY_SCATTER_LINES_NO_MARKERS
            return XL.XY_SCATTER_LINES

        if scatter_style == "smoothMarker":
            if noMarkers():
                return XL.XY_SCATTER_SMOOTH_NO_MARKERS
            return XL.XY_SCATTER_SMOOTH

        return XL.XY_SCATTER

    @classmethod
    def _differentiate_line_3d_chart_type(cls, plot):
        return XL.THREE_D_LINE

    @classmethod
    def _differentiate_of_pie_chart_type(cls, plot):
        ofPieType = plot._element.xpath("c:ofPieType")
        if ofPieType and ofPieType[0].get("val") == "bar":
            return XL.BAR_OF_PIE
        return XL.PIE_OF_PIE

    @classmethod
    def _differentiate_pie_3d_chart_type(cls, plot):
        pie3DChart = plot._element
        explosion = pie3DChart.xpath("./c:ser/c:explosion")
        return XL.THREE_D_PIE_EXPLODED if explosion else XL.THREE_D_PIE

    @classmethod
    def _differentiate_stock_chart_type(cls, plot):
        stockChart = plot._element
        ser_count = len(stockChart.xpath("c:ser"))
        has_bar_sibling = bool(stockChart.xpath("preceding-sibling::c:barChart | following-sibling::c:barChart"))
        if has_bar_sibling:
            return XL.STOCK_VOHLC if ser_count >= 4 else XL.STOCK_VHLC
        return XL.STOCK_OHLC if ser_count >= 4 else XL.STOCK_HLC

    @classmethod
    def _differentiate_surface_3d_chart_type(cls, plot):
        wireframe = plot._element.xpath("c:wireframe")
        if wireframe and wireframe[0].get("val") in ("1", "true"):
            return XL.SURFACE_WIREFRAME
        return XL.SURFACE

    @classmethod
    def _differentiate_surface_chart_type(cls, plot):
        wireframe = plot._element.xpath("c:wireframe")
        if wireframe and wireframe[0].get("val") in ("1", "true"):
            return XL.SURFACE_TOP_VIEW_WIREFRAME
        return XL.SURFACE_TOP_VIEW
