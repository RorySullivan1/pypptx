"""Chart-related objects such as Chart and ChartTitle."""

from __future__ import annotations

from collections.abc import Sequence

from pptx.chart.axis import CategoryAxis, DateAxis, ValueAxis
from pptx.chart.legend import Legend
from pptx.chart.plotarea import PlotArea
from pptx.chart.view3d import ChartSurface, View3D
from pptx.exc import ChartError
from pptx.chart.plot import PlotFactory, PlotTypeInspector
from pptx.chart.series import SeriesCollection
from pptx.chart.xmlwriter import SeriesXmlRewriterFactory
from pptx.dml.chtfmt import ChartFormat
from pptx.oxml.ns import qn
from pptx.shared import ElementProxy, PartElementProxy
from pptx.text.text import Font, TextFrame
from pptx.util import lazyproperty

#: Maps overlay plot type names to their qualified XML tag names.
_OVERLAY_PLOT_TAGS = {
    "line": qn("c:lineChart"),
    "bar": qn("c:barChart"),
    "area": qn("c:areaChart"),
    "scatter": qn("c:scatterChart"),
}


class Chart(PartElementProxy):
    """A chart object."""

    def __init__(self, chartSpace, chart_part):
        super(Chart, self).__init__(chartSpace, chart_part)
        self._chartSpace = chartSpace

    def add_plot(self, plot_type="line", use_secondary_axis=True, grouping="standard"):
        """Add an overlay plot to this chart, creating a combo chart.

        *plot_type* is one of ``"line"``, ``"bar"``, ``"area"``, or
        ``"scatter"``.

        When *use_secondary_axis* is True (the default), a secondary category
        axis and value axis are created for the new plot. Set to False to share
        the primary axes.

        *grouping* specifies the series grouping, e.g. ``"standard"``,
        ``"stacked"``, or ``"percentStacked"``.

        Returns the new |Plot| object. Series can be added to the plot
        via the chart's data management.
        """
        tag = _OVERLAY_PLOT_TAGS.get(plot_type)
        if tag is None:
            raise ChartError(
                f"unsupported overlay plot type '{plot_type}', "
                f"must be one of: {', '.join(sorted(_OVERLAY_PLOT_TAGS))}"
            )

        plotArea = self._chartSpace.chart.plotArea

        if use_secondary_axis:
            cat_ax_id, val_ax_id = plotArea.add_secondary_axes()
        else:
            # Reuse primary axes — get IDs from first plot
            xCharts = plotArea.xCharts
            if not xCharts:
                raise ChartError("chart has no existing plot to share axes with")
            ax_ids = xCharts[0].axId_vals
            if len(ax_ids) < 2:
                raise ChartError("primary plot has no axis references")
            cat_ax_id, val_ax_id = ax_ids[0], ax_ids[1]

        xChart = plotArea.add_xChart(tag, cat_ax_id, val_ax_id, grouping)
        # Clear lazyproperty cache for plots so new plot appears
        if "plots" in self.__dict__:
            del self.__dict__["plots"]
        return PlotFactory(xChart, self)

    @property
    def back_wall(self):
        """A |ChartSurface| object providing access to back wall formatting.

        Accessing this property is destructive in the sense it adds a
        ``c:backWall`` element if not already present.
        """
        chart = self._chartSpace.chart
        return ChartSurface(chart.get_or_add_backWall())

    @property
    def category_axis(self):
        """
        The category axis of this chart. In the case of an XY or Bubble
        chart, this is the X axis. Raises |ValueError| if no category
        axis is defined (as is the case for a pie chart, for example).
        """
        catAx_lst = self._chartSpace.catAx_lst
        if catAx_lst:
            return CategoryAxis(catAx_lst[0])

        dateAx_lst = self._chartSpace.dateAx_lst
        if dateAx_lst:
            return DateAxis(dateAx_lst[0])

        valAx_lst = self._chartSpace.valAx_lst
        if valAx_lst:
            return ValueAxis(valAx_lst[0])

        raise ChartError("chart has no category axis")

    @property
    def chart_style(self):
        """
        Read/write integer index of chart style used to format this chart.
        Range is from 1 to 48. Value is |None| if no explicit style has been
        assigned, in which case the default chart style is used. Assigning
        |None| causes any explicit setting to be removed. The integer index
        corresponds to the style's position in the chart style gallery in the
        PowerPoint UI.
        """
        style = self._chartSpace.style
        if style is None:
            return None
        return style.val

    @chart_style.setter
    def chart_style(self, value):
        self._chartSpace._remove_style()
        if value is None:
            return
        self._chartSpace._add_style(val=value)

    @lazyproperty
    def chart_format(self):
        """|ChartFormat| object providing access to chart area formatting.

        Controls the fill and line properties of the overall chart area
        (the ``c:chartSpace/c:spPr`` element).
        """
        return ChartFormat(self._chartSpace)

    @property
    def chart_title(self):
        """A |ChartTitle| object providing access to title properties.

        Calling this property is destructive in the sense it adds a chart
        title element (`c:title`) to the chart XML if one is not already
        present. Use :attr:`has_title` to test for presence of a chart title
        non-destructively.
        """
        return ChartTitle(self._element.get_or_add_title())

    @property
    def chart_type(self):
        """Member of :ref:`XlChartType` enumeration specifying type of this chart.

        If the chart has two plots, for example, a line plot overlayed on a bar plot,
        the type reported is for the first (back-most) plot. Read-only.
        """
        first_plot = self.plots[0]
        return PlotTypeInspector.chart_type(first_plot)

    @property
    def display_blanks_as(self):
        """Read/write string specifying how blank cells are plotted.

        One of ``"gap"``, ``"zero"``, or ``"span"``. Default is ``"gap"``
        when no ``c:dispBlanksAs`` element is present.
        """
        chart = self._chartSpace.chart
        dispBlanksAs = chart.dispBlanksAs
        if dispBlanksAs is None:
            return "gap"
        return dispBlanksAs.val

    @display_blanks_as.setter
    def display_blanks_as(self, value):
        if value not in ("gap", "zero", "span"):
            raise ChartError(
                f"display_blanks_as must be 'gap', 'zero', or 'span', got '{value}'"
            )
        chart = self._chartSpace.chart
        chart._remove_dispBlanksAs()
        if value != "gap":
            chart._add_dispBlanksAs(val=value)

    @property
    def floor(self):
        """A |ChartSurface| object providing access to floor formatting.

        Accessing this property is destructive in the sense it adds a
        ``c:floor`` element if not already present.
        """
        chart = self._chartSpace.chart
        return ChartSurface(chart.get_or_add_floor())

    @lazyproperty
    def font(self):
        """Font object controlling text format defaults for this chart."""
        defRPr = self._chartSpace.get_or_add_txPr().p_lst[0].get_or_add_pPr().get_or_add_defRPr()
        return Font(defRPr)

    @property
    def has_legend(self):
        """
        Read/write boolean, |True| if the chart has a legend. Assigning
        |True| causes a legend to be added to the chart if it doesn't already
        have one. Assigning False removes any existing legend definition
        along with any existing legend settings.
        """
        return self._chartSpace.chart.has_legend

    @has_legend.setter
    def has_legend(self, value):
        self._chartSpace.chart.has_legend = bool(value)

    @property
    def has_title(self):
        """Read/write boolean, specifying whether this chart has a title.

        Assigning |True| causes a title to be added if not already present.
        Assigning |False| removes any existing title along with its text and
        settings.
        """
        title = self._chartSpace.chart.title
        if title is None:
            return False
        return True

    @has_title.setter
    def has_title(self, value):
        chart = self._chartSpace.chart
        if bool(value) is False:
            chart._remove_title()
            autoTitleDeleted = chart.get_or_add_autoTitleDeleted()
            autoTitleDeleted.val = True
            return
        chart.get_or_add_title()

    @property
    def legend(self):
        """
        A |Legend| object providing access to the properties of the legend
        for this chart.
        """
        legend_elm = self._chartSpace.chart.legend
        if legend_elm is None:
            return None
        return Legend(legend_elm)

    @lazyproperty
    def plot_area(self):
        """A |PlotArea| object providing access to plot area properties.

        Includes manual layout (position/size) and shape formatting (fill/line).
        """
        return PlotArea(self._chartSpace.chart.plotArea)

    @lazyproperty
    def plots(self):
        """
        The sequence of plots in this chart. A plot, called a *chart group*
        in the Microsoft API, is a distinct sequence of one or more series
        depicted in a particular charting type. For example, a chart having
        a series plotted as a line overlaid on three series plotted as
        columns would have two plots; the first corresponding to the three
        column series and the second to the line series. Plots are sequenced
        in the order drawn, i.e. back-most to front-most. Supports *len()*,
        membership (e.g. ``p in plots``), iteration, slicing, and indexed
        access (e.g. ``plot = plots[i]``).
        """
        plotArea = self._chartSpace.chart.plotArea
        return _Plots(plotArea, self)

    def replace_data(self, chart_data):
        """
        Use the categories and series values in the |ChartData| object
        *chart_data* to replace those in the XML and Excel worksheet for this
        chart.
        """
        rewriter = SeriesXmlRewriterFactory(self.chart_type, chart_data)
        rewriter.replace_series_data(self._chartSpace)
        self._workbook.update_from_xlsx_blob(chart_data.xlsx_blob)

    @lazyproperty
    def series(self):
        """
        A |SeriesCollection| object containing all the series in this
        chart. When the chart has multiple plots, all the series for the
        first plot appear before all those for the second, and so on. Series
        within a plot have an explicit ordering and appear in that sequence.
        """
        return SeriesCollection(self._chartSpace.plotArea)

    @property
    def side_wall(self):
        """A |ChartSurface| object providing access to side wall formatting.

        Accessing this property is destructive in the sense it adds a
        ``c:sideWall`` element if not already present.
        """
        chart = self._chartSpace.chart
        return ChartSurface(chart.get_or_add_sideWall())

    @property
    def value_axis(self):
        """The primary |ValueAxis| of this chart.

        Raises |ChartError| if the chart has no value axis.
        """
        valAx_lst = self._chartSpace.valAx_lst
        if not valAx_lst:
            raise ChartError("chart has no value axis")
        return ValueAxis(valAx_lst[0])

    @property
    def secondary_value_axis(self):
        """The secondary |ValueAxis| of this chart.

        Present on combo charts and charts with a secondary axis.
        Raises |ChartError| if no secondary value axis exists.
        """
        valAx_lst = self._chartSpace.valAx_lst
        if len(valAx_lst) < 2:
            raise ChartError("chart has no secondary value axis")
        return ValueAxis(valAx_lst[1])

    @property
    def secondary_category_axis(self):
        """The secondary |CategoryAxis| or |DateAxis| of this chart.

        Present on combo charts and charts with a secondary axis.
        Raises |ChartError| if no secondary category axis exists.
        """
        catAx_lst = self._chartSpace.catAx_lst
        if len(catAx_lst) >= 2:
            return CategoryAxis(catAx_lst[1])

        dateAx_lst = self._chartSpace.dateAx_lst
        if len(dateAx_lst) >= 2:
            return DateAxis(dateAx_lst[1])

        raise ChartError("chart has no secondary category axis")

    @property
    def view_3d(self):
        """A |View3D| object providing access to 3D view properties.

        Accessing this property is destructive in the sense it adds a
        ``c:view3D`` element if not already present.
        """
        chart = self._chartSpace.chart
        return View3D(chart.get_or_add_view3D())

    @property
    def _workbook(self):
        """
        The |ChartWorkbook| object providing access to the Excel source data
        for this chart.
        """
        return self.part.chart_workbook


class ChartTitle(ElementProxy):
    """Provides properties for manipulating a chart title."""

    # This shares functionality with AxisTitle, which could be factored out
    # into a base class, perhaps pptx.chart.shared.BaseTitle. I suspect they
    # actually differ in certain fuller behaviors, but at present they're
    # essentially identical.

    def __init__(self, title):
        super(ChartTitle, self).__init__(title)
        self._title = title

    @lazyproperty
    def format(self):
        """|ChartFormat| object providing access to line and fill formatting.

        Return the |ChartFormat| object providing shape formatting properties
        for this chart title, such as its line color and fill.
        """
        return ChartFormat(self._title)

    @property
    def has_text_frame(self):
        """Read/write Boolean specifying whether this title has a text frame.

        Return |True| if this chart title has a text frame, and |False|
        otherwise. Assigning |True| causes a text frame to be added if not
        already present. Assigning |False| causes any existing text frame to
        be removed along with its text and formatting.
        """
        if self._title.tx_rich is None:
            return False
        return True

    @has_text_frame.setter
    def has_text_frame(self, value):
        if bool(value) is False:
            self._title._remove_tx()
            return
        self._title.get_or_add_tx_rich()

    @property
    def text_frame(self):
        """|TextFrame| instance for this chart title.

        Return a |TextFrame| instance allowing read/write access to the text
        of this chart title and its text formatting properties. Accessing this
        property is destructive in the sense it adds a text frame if one is
        not present. Use :attr:`has_text_frame` to test for the presence of
        a text frame non-destructively.
        """
        rich = self._title.get_or_add_tx_rich()
        return TextFrame(rich, self)


class _Plots(Sequence):
    """
    The sequence of plots in a chart, such as a bar plot or a line plot. Most
    charts have only a single plot. The concept is necessary when two chart
    types are displayed in a single set of axes, like a bar plot with
    a superimposed line plot.
    """

    def __init__(self, plotArea, chart):
        super(_Plots, self).__init__()
        self._plotArea = plotArea
        self._chart = chart

    def __getitem__(self, index):
        xCharts = self._plotArea.xCharts
        if isinstance(index, slice):
            plots = [PlotFactory(xChart, self._chart) for xChart in xCharts]
            return plots[index]
        else:
            xChart = xCharts[index]
            return PlotFactory(xChart, self._chart)

    def __len__(self):
        return len(self._plotArea.xCharts)
