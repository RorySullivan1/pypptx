"""Writers for new chartex charts: the `cx:chartSpace` XML and its embedded workbook.

A chartex chart is described by a |CategoryChartData| object, like a classic category chart:

* waterfall, funnel -- one series, one value per category
* histogram -- one series of raw values (categories optional); PowerPoint bins them
* box & whisker -- one or more series of raw values (categories optional, to group points)
* treemap, sunburst -- one series of sizes over (usually hierarchical) categories

The workbook uses the same layout as a classic category chart (categories in the leftmost
columns, then one column per series), except that a parent category label is repeated on every
row it spans, so a hierarchy reads unambiguously from any row.
"""

from __future__ import annotations

from typing import TYPE_CHECKING
from xml.sax.saxutils import escape

from pptx.chart.xlsx import CategoryWorkbookWriter
from pptx.enum.chart import XL_CHART_TYPE
from pptx.exc import ChartError

if TYPE_CHECKING:
    from pptx.chart.data import CategoryChartData, CategorySeriesData

CHARTEX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex"

# -- MC namespace a PowerPoint must understand to render each chart type: waterfall, histogram,
# -- Pareto, box & whisker, treemap and sunburst came in one release, funnel in a later one --
_CX1 = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
_CX2 = "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"

_CAT_AXIS = '<cx:axis id="0"><cx:catScaling gapWidth="%s"/><cx:tickLabels/></cx:axis>'
_VAL_AXIS = '<cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>'
_LEGEND = '<cx:legend pos="t" align="ctr" overlay="0"/>'


def _data_labels(pos: str, category_name: bool, value: bool) -> str:
    return (
        '<cx:dataLabels pos="%s"><cx:visibility seriesName="0" categoryName="%d" value="%d"/>'
        "</cx:dataLabels>" % (pos, category_name, value)
    )


class _ChartExType:
    """Per-type particulars of the XML PowerPoint writes for a new chartex chart."""

    def __init__(
        self,
        layout_id: str,
        requires_ns: str,
        value_dim_type: str = "val",
        single_series: bool = True,
        needs_categories: bool = False,
        data_labels: str = "",
        layout_pr: str = "",
        axes: str = "",
        legend: str = "",
    ):
        self.layout_id = layout_id
        self.requires_ns = requires_ns
        self.value_dim_type = value_dim_type
        self.single_series = single_series
        self.needs_categories = needs_categories
        self.data_labels = data_labels
        self.layout_pr = layout_pr
        self.axes = axes
        self.legend = legend


CHARTEX_TYPES: dict[XL_CHART_TYPE, _ChartExType] = {
    XL_CHART_TYPE.WATERFALL: _ChartExType(
        "waterfall",
        _CX1,
        data_labels=_data_labels("outEnd", False, True),
        axes=_CAT_AXIS % "0.5" + _VAL_AXIS,
        legend=_LEGEND,
    ),
    XL_CHART_TYPE.HISTOGRAM: _ChartExType(
        "clusteredColumn",
        _CX1,
        layout_pr='<cx:layoutPr><cx:binning intervalClosed="r"/></cx:layoutPr>',
        axes=_CAT_AXIS % "0" + _VAL_AXIS,
    ),
    XL_CHART_TYPE.BOX_WHISKER: _ChartExType(
        "boxWhisker",
        _CX1,
        single_series=False,
        layout_pr=(
            '<cx:layoutPr><cx:visibility meanLine="0" meanMarker="1" nonoutliers="0"'
            ' outliers="1"/><cx:statistics quartileMethod="exclusive"/></cx:layoutPr>'
        ),
        axes=_CAT_AXIS % "1" + _VAL_AXIS,
        legend=_LEGEND,
    ),
    XL_CHART_TYPE.TREEMAP: _ChartExType(
        "treemap",
        _CX1,
        value_dim_type="size",
        needs_categories=True,
        data_labels=_data_labels("inEnd", True, False),
        layout_pr='<cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr>',
        legend=_LEGEND,
    ),
    XL_CHART_TYPE.SUNBURST: _ChartExType(
        "sunburst",
        _CX1,
        value_dim_type="size",
        needs_categories=True,
        data_labels=_data_labels("ctr", True, False),
    ),
    XL_CHART_TYPE.FUNNEL: _ChartExType(
        "funnel",
        _CX2,
        data_labels=_data_labels("inEnd", False, True),
        axes=_CAT_AXIS % "0.06",
    ),
}


def chartex_type(chart_type: XL_CHART_TYPE) -> _ChartExType:
    """The particulars for `chart_type`; raises |ChartError| if it is not an addable chartex type."""
    try:
        return CHARTEX_TYPES[chart_type]
    except KeyError:
        raise ChartError("chart type %s cannot be added as a chartex chart" % chart_type)


def is_chartex_type(chart_type: object) -> bool:
    """|True| if `chart_type` is a chartex chart type that `add_chart()` can create."""
    return chart_type in CHARTEX_TYPES


class ChartExXmlWriter:
    """Generates the `cx:chartSpace` XML for a new chartex chart of `chart_type`.

    The `cx:externalData` workbook reference is not written; the chart part adds it once the
    workbook part is related.
    """

    def __init__(self, chart_type: XL_CHART_TYPE, chart_data: CategoryChartData):
        self._chart_type = chart_type
        self._type = chartex_type(chart_type)
        self._chart_data = chart_data

    @property
    def xml(self) -> str:
        """The chart-part XML, a complete document including the XML declaration."""
        self._validate()
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            ' xmlns:cx="%s">'
            "<cx:chartData>%s</cx:chartData>"
            "<cx:chart><cx:plotArea><cx:plotAreaRegion>%s</cx:plotAreaRegion>%s</cx:plotArea>"
            "%s</cx:chart>"
            "</cx:chartSpace>"
        ) % (
            CHARTEX_NS,
            "".join(self._data_xml(series) for series in self._chart_data),
            "".join(self._series_xml(series) for series in self._chart_data),
            self._type.axes,
            self._type.legend,
        )

    def _cat_dim_xml(self) -> str:
        """`cx:strDim` for the categories, leaf level first; "" when there are no categories."""
        categories = self._chart_data.categories
        depth = categories.depth
        if depth == 0:
            return ""
        leaf_count = categories.leaf_count
        lvls = "".join(
            '<cx:lvl ptCount="%d">%s</cx:lvl>'
            % (
                leaf_count,
                "".join(
                    '<cx:pt idx="%d">%s</cx:pt>' % (idx, escape(str(label)))
                    for idx, label in enumerate(_filled_level(level, leaf_count))
                ),
            )
            for level in categories.levels
        )
        dir_attr = ' dir="row"' if depth > 1 else ""
        return '<cx:strDim type="cat"><cx:f%s>%s</cx:f>%s</cx:strDim>' % (
            dir_attr,
            self._chart_data.categories_ref,
            lvls,
        )

    def _data_xml(self, series: CategorySeriesData) -> str:
        values = series.values
        pts = "".join(
            '<cx:pt idx="%d">%s</cx:pt>' % (idx, value)
            for idx, value in enumerate(values)
            if value is not None
        )
        return (
            '<cx:data id="%d">%s<cx:numDim type="%s"><cx:f>%s</cx:f>'
            '<cx:lvl ptCount="%d" formatCode="%s">%s</cx:lvl></cx:numDim></cx:data>'
        ) % (
            series.index,
            self._cat_dim_xml(),
            self._type.value_dim_type,
            series.values_ref,
            len(values),
            escape(series.number_format, {'"': "&quot;"}),
            pts,
        )

    def _series_xml(self, series: CategorySeriesData) -> str:
        return (
            '<cx:series layoutId="%s" uniqueId="{%08X-0000-4000-8000-000000000000}">'
            "<cx:tx><cx:txData><cx:f>%s</cx:f><cx:v>%s</cx:v></cx:txData></cx:tx>"
            '%s<cx:dataId val="%d"/>%s</cx:series>'
        ) % (
            self._type.layout_id,
            series.index + 1,
            series.name_ref,
            escape(str(series.name)),
            self._type.data_labels,
            series.index,
            self._type.layout_pr,
        )

    def _validate(self) -> None:
        series_count = len(self._chart_data)
        if series_count == 0:
            raise ChartError("a %s chart needs at least one series" % self._chart_type.name)
        if self._type.single_series and series_count > 1:
            raise ChartError("a %s chart takes exactly one series" % self._chart_type.name)
        if self._type.needs_categories and self._chart_data.categories.depth == 0:
            raise ChartError("a %s chart needs categories" % self._chart_type.name)


class ChartExWorkbookWriter(CategoryWorkbookWriter):
    """Writes the embedded workbook for a chartex chart.

    Same layout as |CategoryWorkbookWriter|, but each parent category label fills every row of
    the leaf categories under it rather than only the first.
    """

    def _write_categories(self, workbook, worksheet):
        categories = self._chart_data.categories
        num_format = workbook.add_format({"num_format": categories.number_format})
        depth = categories.depth
        leaf_count = categories.leaf_count
        for idx, level in enumerate(categories.levels):
            col = depth - idx - 1
            filled = list(enumerate(_filled_level(level, leaf_count)))
            self._write_cat_column(worksheet, col, filled, num_format)


def _filled_level(level: list[tuple[int, object]], leaf_count: int) -> list[object]:
    """Labels of `level` (`(first-leaf-idx, label)` pairs) repeated over every leaf they span."""
    labels: list[object] = [""] * leaf_count
    starts = sorted(level, key=lambda item: item[0])
    for n, (start, label) in enumerate(starts):
        end = starts[n + 1][0] if n + 1 < len(starts) else leaf_count
        for idx in range(start, end):
            labels[idx] = label
    return labels
