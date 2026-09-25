# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.chartex` module."""

from __future__ import annotations

from pptx import Presentation
from pptx.chart.chartex import ChartEx
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import PartFactory
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.parts.chartex import ChartExPart
from pptx.parts.embeddedpackage import EmbeddedXlsxPart

from ..unitutil.chartex import WATERFALL_XML, xlsx_column
from ..unitutil.mock import function_mock, instance_mock


class DescribeChartExPart:
    """Unit-test suite for `pptx.parts.chartex.ChartExPart` objects."""

    def it_is_registered_for_the_chartex_content_type(self):
        assert PartFactory.part_type_for[CT.OFC_CHART_EX] is ChartExPart

    def it_can_create_a_new_chart_with_its_workbook(self):
        package = Presentation().part.package
        chart_data = CategoryChartData()
        chart_data.categories = ["Start", "End"]
        chart_data.add_series("Flow", (10, 12))

        part = ChartExPart.new(XL_CHART_TYPE.WATERFALL, chart_data, package)

        assert part.partname == PackURI("/ppt/charts/chartEx1.xml")
        assert part.content_type == CT.OFC_CHART_EX
        xlsx_part = part.xlsx_part
        assert isinstance(xlsx_part, EmbeddedXlsxPart)
        assert xlsx_column(xlsx_part.blob, "B") == [10.0, 12.0]
        externalData = part.chartSpace.xpath("./cx:chartData/cx:externalData")[0]
        assert externalData.get(qn("cx:autoUpdate")) == "0"
        assert part.chartex.chart_type is XL_CHART_TYPE.WATERFALL

    def it_provides_access_to_the_chart(self, request):
        package_ = instance_mock(request, "pptx.package.Package")
        part = ChartExPart.load(
            PackURI("/ppt/charts/chartEx1.xml"),
            CT.OFC_CHART_EX,
            package_,
            WATERFALL_XML.encode("utf-8"),
        )

        chartex = part.chartex

        assert isinstance(chartex, ChartEx)
        assert chartex.part is part
        assert chartex.chart_type is XL_CHART_TYPE.WATERFALL

    def it_has_no_workbook_when_there_is_no_external_data_reference(self, request):
        package_ = instance_mock(request, "pptx.package.Package")
        blob = (
            b'<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
            b"/>"
        )
        part = ChartExPart.load(
            PackURI("/ppt/charts/chartEx1.xml"), CT.OFC_CHART_EX, package_, blob
        )

        assert part.xlsx_part is None

    def it_never_parses_when_only_the_blob_is_read_back(self, request):
        parse_xml_ = function_mock(request, "pptx.opc.package.parse_xml")
        package_ = instance_mock(request, "pptx.package.Package")
        blob = WATERFALL_XML.encode("utf-8")

        part = ChartExPart.load(
            PackURI("/ppt/charts/chartEx1.xml"), CT.OFC_CHART_EX, package_, blob
        )

        assert part.blob == blob
        parse_xml_.assert_not_called()
