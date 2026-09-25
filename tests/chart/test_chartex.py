"""Unit-test suite for `pptx.chart.chartex` module."""

from __future__ import annotations

import pytest

from pptx.chart.chartex import ChartEx, ChartExSeries
from pptx.enum.chart import XL_CHART_TYPE

from ..unitutil.cxml import element


def _chartSpace(*layout_ids: str):
    series = ",".join("cx:series{layoutId=%s}" % layout_id for layout_id in layout_ids)
    cxml = "cx:chartSpace/cx:chart/cx:plotArea/cx:plotAreaRegion"
    return element("%s/(%s)" % (cxml, series) if series else cxml)


class DescribeChartEx:
    """Unit-test suite for `pptx.chart.chartex.ChartEx` objects."""

    @pytest.mark.parametrize(
        ("layout_ids", "expected_value"),
        [
            (("waterfall",), XL_CHART_TYPE.WATERFALL),
            (("clusteredColumn",), XL_CHART_TYPE.HISTOGRAM),
            (("clusteredColumn", "paretoLine"), XL_CHART_TYPE.PARETO),
            (("boxWhisker", "boxWhisker"), XL_CHART_TYPE.BOX_WHISKER),
            (("treemap",), XL_CHART_TYPE.TREEMAP),
            (("sunburst",), XL_CHART_TYPE.SUNBURST),
            (("funnel",), XL_CHART_TYPE.FUNNEL),
            (("regionMap",), XL_CHART_TYPE.REGION_MAP),
            (("somethingNew",), None),
            ((), None),
        ],
    )
    def it_knows_its_chart_type(self, layout_ids, expected_value):
        assert ChartEx(_chartSpace(*layout_ids), None).chart_type is expected_value

    def it_provides_access_to_its_series(self):
        chartex = ChartEx(_chartSpace("clusteredColumn", "paretoLine"), None)

        series = chartex.series

        assert all(isinstance(s, ChartExSeries) for s in series)
        assert [s.layout_id for s in series] == ["clusteredColumn", "paretoLine"]

    def it_provides_access_to_its_part(self):
        part = object()
        assert ChartEx(_chartSpace(), part).part is part


class DescribeChartExSeries:
    """Unit-test suite for `pptx.chart.chartex.ChartExSeries` objects."""

    def it_reads_its_name_values_and_categories(self):
        chartSpace = element(
            "cx:chartSpace/("
            "cx:chartData/cx:data{id=4}/("
            'cx:strDim{type=cat}/cx:lvl{ptCount=3}/(cx:pt{idx=0}"a",cx:pt{idx=1}"b",'
            'cx:pt{idx=2}"c"),'
            'cx:numDim{type=val}/cx:lvl{ptCount=3}/(cx:pt{idx=0}"1.5",cx:pt{idx=2}"-3")),'
            "cx:chart/cx:plotArea/cx:plotAreaRegion/cx:series{layoutId=waterfall}/("
            'cx:tx/cx:txData/cx:v"Flow",cx:dataId{val=4}))'
        )
        series = ChartEx(chartSpace, None).series[0]

        assert series.name == "Flow"
        assert series.values == (1.5, None, -3.0)
        assert series.categories == ("a", "b", "c")
        assert series.category_paths == (("a",), ("b",), ("c",))
        assert series.hidden is False

    def it_fills_missing_parent_labels_from_the_point_above(self):
        chartSpace = element(
            "cx:chartSpace/("
            "cx:chartData/cx:data{id=0}/cx:strDim{type=cat}/("
            'cx:lvl{ptCount=3}/(cx:pt{idx=0}"Apple",cx:pt{idx=1}"Pear",cx:pt{idx=2}"Kale"),'
            'cx:lvl{ptCount=3}/(cx:pt{idx=0}"Fruit",cx:pt{idx=2}"Veg")),'
            "cx:chart/cx:plotArea/cx:plotAreaRegion/cx:series{layoutId=treemap}/"
            "cx:dataId{val=0})"
        )
        series = ChartEx(chartSpace, None).series[0]

        assert series.category_paths == (
            ("Fruit", "Apple"),
            ("Fruit", "Pear"),
            ("Veg", "Kale"),
        )
        assert series.categories == ("Apple", "Pear", "Kale")
        assert series.values == ()

    def it_has_no_data_when_it_has_no_data_id(self):
        chartSpace = element(
            "cx:chartSpace/cx:chart/cx:plotArea/cx:plotAreaRegion/"
            "cx:series{layoutId=paretoLine,ownerIdx=0}"
        )
        series = ChartEx(chartSpace, None).series[0]

        assert (series.name, series.values, series.categories) == (None, (), ())
