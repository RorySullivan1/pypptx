"""Unit-test suite for `pptx.chart.chartexwriter` module."""

from __future__ import annotations

import pytest

from pptx.chart.chartex import ChartEx
from pptx.chart.chartexwriter import (
    ChartExWorkbookWriter,
    ChartExXmlWriter,
    chartex_type,
    is_chartex_type,
)
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.exc import ChartError
from pptx.oxml import parse_xml

from ..unitutil.chartex import xlsx_column


def _flat_data(n_series: int = 1) -> CategoryChartData:
    chart_data = CategoryChartData()
    chart_data.categories = ["Start", "Up", "End"]
    for idx in range(n_series):
        chart_data.add_series("S%d" % (idx + 1), (10, 5, 15))
    return chart_data


def _tree_data() -> CategoryChartData:
    chart_data = CategoryChartData()
    fruit = chart_data.add_category("Fruit")
    fruit.add_sub_category("Apple")
    fruit.add_sub_category("Pear")
    chart_data.add_category("Veg").add_sub_category("Kale")
    chart_data.add_series("Sales", (30, 12, 7))
    return chart_data


def _chartex(chart_type: XL_CHART_TYPE, chart_data: CategoryChartData) -> ChartEx:
    xml = ChartExXmlWriter(chart_type, chart_data).xml
    return ChartEx(parse_xml(xml.encode("utf-8")), None)


class DescribeChartExXmlWriter:
    """Unit-test suite for `pptx.chart.chartexwriter.ChartExXmlWriter` objects."""

    @pytest.mark.parametrize(
        ("chart_type", "layout_id", "dim_type", "has_axes"),
        [
            (XL_CHART_TYPE.WATERFALL, "waterfall", "val", True),
            (XL_CHART_TYPE.HISTOGRAM, "clusteredColumn", "val", True),
            (XL_CHART_TYPE.BOX_WHISKER, "boxWhisker", "val", True),
            (XL_CHART_TYPE.FUNNEL, "funnel", "val", True),
            (XL_CHART_TYPE.TREEMAP, "treemap", "size", False),
            (XL_CHART_TYPE.SUNBURST, "sunburst", "size", False),
        ],
    )
    def it_writes_a_chart_that_reads_back_as_its_type(
        self, chart_type, layout_id, dim_type, has_axes
    ):
        xml = ChartExXmlWriter(chart_type, _flat_data()).xml
        chartSpace = parse_xml(xml.encode("utf-8"))

        chartex = ChartEx(chartSpace, None)
        assert chartex.chart_type is chart_type
        (series,) = chartex.series
        assert series.layout_id == layout_id
        assert (series.name, series.values, series.categories) == (
            "S1",
            (10.0, 5.0, 15.0),
            ("Start", "Up", "End"),
        )
        assert chartSpace.xpath("//cx:numDim/@type") == [dim_type]
        assert bool(chartSpace.xpath("//cx:axis")) is has_axes
        assert xml.startswith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')

    def it_writes_hierarchical_categories_leaf_level_first(self):
        chartSpace = parse_xml(
            ChartExXmlWriter(XL_CHART_TYPE.TREEMAP, _tree_data()).xml.encode("utf-8")
        )

        assert chartSpace.xpath("//cx:strDim/cx:f/@dir") == ["row"]
        assert chartSpace.xpath("//cx:strDim/cx:f/text()") == ["Sheet1!$A$2:$B$4"]
        assert [lvl.texts() for lvl in chartSpace.xpath("//cx:strDim/cx:lvl")] == [
            ["Apple", "Pear", "Kale"],
            ["Fruit", "Fruit", "Veg"],
        ]
        assert chartSpace.xpath("//cx:numDim/cx:f/text()") == ["Sheet1!$C$2:$C$4"]

    def it_writes_one_data_item_per_series_for_box_and_whisker(self):
        chartex = _chartex(XL_CHART_TYPE.BOX_WHISKER, _flat_data(n_series=2))

        assert [s.name for s in chartex.series] == ["S1", "S2"]
        assert chartex.part is None
        assert chartex.series[1].values == (10.0, 5.0, 15.0)

    def it_omits_the_category_dimension_when_there_are_no_categories(self):
        chart_data = CategoryChartData()
        chart_data.add_series("Scores", (1, 2, 2, 3))

        chartex = _chartex(XL_CHART_TYPE.HISTOGRAM, chart_data)

        (series,) = chartex.series
        assert series.categories == ()
        assert series.values == (1.0, 2.0, 2.0, 3.0)

    def it_escapes_names_and_labels(self):
        chart_data = CategoryChartData()
        chart_data.categories = ["R&D", "<Ops>"]
        chart_data.add_series('Cost "A"', (1, 2))

        (series,) = _chartex(XL_CHART_TYPE.FUNNEL, chart_data).series

        assert series.name == 'Cost "A"'
        assert series.categories == ("R&D", "<Ops>")

    @pytest.mark.parametrize(
        ("chart_type", "chart_data", "message"),
        [
            (XL_CHART_TYPE.WATERFALL, CategoryChartData(), "needs at least one series"),
            (XL_CHART_TYPE.WATERFALL, _flat_data(n_series=2), "takes exactly one series"),
            (XL_CHART_TYPE.FUNNEL, _flat_data(n_series=2), "takes exactly one series"),
        ],
    )
    def it_rejects_data_the_chart_type_cannot_show(self, chart_type, chart_data, message):
        with pytest.raises(ChartError, match=message):
            ChartExXmlWriter(chart_type, chart_data).xml

    def but_treemap_and_sunburst_need_categories(self):
        chart_data = CategoryChartData()
        chart_data.add_series("Sales", (1, 2))

        for chart_type in (XL_CHART_TYPE.TREEMAP, XL_CHART_TYPE.SUNBURST):
            with pytest.raises(ChartError, match="needs categories"):
                ChartExXmlWriter(chart_type, chart_data).xml

    def it_rejects_a_chart_type_that_is_not_an_addable_chartex_type(self):
        with pytest.raises(ChartError, match="cannot be added as a chartex chart"):
            ChartExXmlWriter(XL_CHART_TYPE.PARETO, _flat_data())


class DescribeChartExTypes:
    """Unit-test suite for the chartex type table in `pptx.chart.chartexwriter`."""

    @pytest.mark.parametrize(
        ("chart_type", "expected_value"),
        [
            (XL_CHART_TYPE.WATERFALL, True),
            (XL_CHART_TYPE.SUNBURST, True),
            (XL_CHART_TYPE.PARETO, False),
            (XL_CHART_TYPE.REGION_MAP, False),
            (XL_CHART_TYPE.COLUMN_CLUSTERED, False),
        ],
    )
    def it_knows_which_types_it_can_add(self, chart_type, expected_value):
        assert is_chartex_type(chart_type) is expected_value

    def it_requires_the_later_namespace_for_funnel_only(self):
        assert chartex_type(XL_CHART_TYPE.FUNNEL).requires_ns.endswith("2015/10/21/chartex")
        assert chartex_type(XL_CHART_TYPE.WATERFALL).requires_ns.endswith("2015/9/8/chartex")


class DescribeChartExWorkbookWriter:
    """Unit-test suite for `pptx.chart.chartexwriter.ChartExWorkbookWriter` objects."""

    def it_repeats_parent_labels_on_every_row_they_span(self):
        xlsx_blob = ChartExWorkbookWriter(_tree_data()).xlsx_blob

        assert xlsx_column(xlsx_blob, "A") == ["Fruit", "Fruit", "Veg"]
        assert xlsx_column(xlsx_blob, "B") == ["Apple", "Pear", "Kale"]
        assert xlsx_column(xlsx_blob, "C", first_row=1) == ["Sales", 30.0, 12.0, 7.0]

    def it_writes_values_alone_when_there_are_no_categories(self):
        chart_data = CategoryChartData()
        chart_data.add_series("Scores", (1, 2, 3))

        xlsx_blob = ChartExWorkbookWriter(chart_data).xlsx_blob

        assert xlsx_column(xlsx_blob, "A", first_row=1) == ["Scores", 1.0, 2.0, 3.0]
