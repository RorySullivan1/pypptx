"""Unit-test suite for `pptx.oxml.chart.chartex` module."""

from __future__ import annotations

import pytest

from pptx.oxml.chart.chartex import (
    CT_ChartExChart,
    CT_ChartExDataItem,
    CT_ChartExSeries,
    CT_ChartExSpace,
)

from ..unitutil.cxml import element

SERIES_CXML = (
    "cx:chart/cx:plotArea/cx:plotAreaRegion/("
    "cx:series{layoutId=clusteredColumn}/cx:dataId{val=0},"
    "cx:series{layoutId=paretoLine,ownerIdx=0,hidden=1})"
)


class DescribeCT_ChartExSpace:
    """Unit-test suite for `pptx.oxml.chart.chartex.CT_ChartExSpace` objects."""

    def it_provides_access_to_its_series(self):
        chartSpace = element("cx:chartSpace/%s" % SERIES_CXML)

        assert isinstance(chartSpace, CT_ChartExSpace)
        assert [ser.layoutId for ser in chartSpace.series_lst] == ["clusteredColumn", "paretoLine"]

    def it_finds_the_data_for_an_id(self):
        chartSpace = element("cx:chartSpace/cx:chartData/(cx:data{id=0},cx:data{id=3})")

        data = chartSpace.data_for(3)

        assert isinstance(data, CT_ChartExDataItem)
        assert data.id == 3
        assert chartSpace.data_for(1) is None
        assert element("cx:chartSpace").data_for(0) is None

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("cx:chartSpace/cx:chartData/cx:externalData{r:id=rId4}", "rId4"),
            ("cx:chartSpace/cx:chartData", None),
            ("cx:chartSpace", None),
        ],
    )
    def it_knows_its_workbook_rId(self, cxml: str, expected_value: str | None):
        assert element(cxml).xlsx_part_rId == expected_value

    def it_adds_the_workbook_reference_ahead_of_the_data(self):
        chartSpace = element("cx:chartSpace/cx:chartData/cx:data{id=0}")

        externalData = chartSpace.get_or_add_chartData().get_or_add_externalData()
        externalData.rId = "rId1"

        assert [child.tag.split("}")[1] for child in chartSpace.chartData] == [
            "externalData",
            "data",
        ]
        assert chartSpace.xlsx_part_rId == "rId1"


class DescribeCT_ChartExDataItem:
    """Unit-test suite for `pptx.oxml.chart.chartex.CT_ChartExDataItem` objects."""

    def it_finds_its_category_and_value_dimensions(self):
        data = element(
            "cx:data{id=0}/(cx:numDim{type=size},cx:strDim{type=cat}/cx:f\"Sheet1!$A$2:$B$4\")"
        )

        assert data.cat_dim.formula == "Sheet1!$A$2:$B$4"
        assert data.val_dim.type == "size"

    def it_accepts_a_numeric_category_dimension(self):
        data = element("cx:data{id=0}/(cx:numDim{type=cat},cx:numDim{type=val})")

        assert data.cat_dim.type == "cat"
        assert data.val_dim.type == "val"

    def it_has_no_dimensions_when_they_are_absent(self):
        data = element("cx:data{id=0}")

        assert data.cat_dim is None
        assert data.val_dim is None


class DescribeCT_ChartExLevel:
    """Unit-test suite for `pptx.oxml.chart.chartex.CT_ChartExLevel` objects."""

    def it_lists_its_point_texts_by_index(self):
        lvl = element(
            'cx:lvl{ptCount=4}/(cx:pt{idx=2}"c",cx:pt{idx=0}"a",cx:pt{idx=1},cx:pt{idx=9}"x")'
        )

        assert lvl.texts() == ["a", "", "c", None]


class DescribeCT_ChartExChart:
    """Unit-test suite for `pptx.oxml.chart.chartex.CT_ChartExChart` objects."""

    def it_is_also_the_graphic_frame_reference_to_the_part(self):
        chart = element("cx:chart{r:id=rId2}")

        assert isinstance(chart, CT_ChartExChart)
        assert chart.rId == "rId2"
        assert chart.plotArea is None


class DescribeCT_ChartExSeries:
    """Unit-test suite for `pptx.oxml.chart.chartex.CT_ChartExSeries` objects."""

    def it_knows_its_layout_and_data_id(self):
        chart = element(SERIES_CXML)
        column, pareto = chart.xpath(".//cx:series")

        assert isinstance(column, CT_ChartExSeries)
        assert (column.layoutId, column.data_id, column.hidden, column.ownerIdx) == (
            "clusteredColumn",
            0,
            False,
            None,
        )
        assert (pareto.data_id, pareto.hidden, pareto.ownerIdx) == (None, True, 0)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ('cx:series{layoutId=funnel}/cx:tx/cx:txData/(cx:f"Sheet1!$B$1",cx:v"Leads")', "Leads"),
            ("cx:series{layoutId=funnel}/cx:tx/cx:txData/cx:v", None),
            ("cx:series{layoutId=funnel}", None),
        ],
    )
    def it_knows_its_name(self, cxml: str, expected_value: str | None):
        assert element(cxml).name == expected_value
