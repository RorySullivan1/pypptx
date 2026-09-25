"""Chartex part: an Office 2016+ chart (waterfall, histogram, treemap, ...) and its workbook."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.chart.chartex import ChartEx
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import LazyXmlPart
from pptx.oxml.ns import qn
from pptx.parts.embeddedpackage import EmbeddedXlsxPart

if TYPE_CHECKING:
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.oxml.chart.chartex import CT_ChartExSpace
    from pptx.package import Package


class ChartExPart(LazyXmlPart):
    """A chartex part, e.g. ``/ppt/charts/chartEx1.xml``, holding a `cx:chartSpace`.

    Loaded lazily (see |LazyXmlPart|), so a chartex chart that is opened and re-saved without
    being read writes its bytes back unchanged.
    """

    partname_template = "/ppt/charts/chartEx%d.xml"

    @classmethod
    def new(
        cls, chart_type: XL_CHART_TYPE, chart_data: CategoryChartData, package: Package
    ) -> ChartExPart:
        """Return a new |ChartExPart| depicting `chart_data` as a chart of `chart_type`.

        The chart's data is also written to a new embedded workbook related to the part, so the
        chart can be edited in PowerPoint.
        """
        from pptx.chart.chartexwriter import ChartExWorkbookWriter, ChartExXmlWriter

        chart_part = cls.load(
            package.next_partname(cls.partname_template),
            CT.OFC_CHART_EX,
            package,
            ChartExXmlWriter(chart_type, chart_data).xml.encode("utf-8"),
        )
        xlsx_part = EmbeddedXlsxPart.new(ChartExWorkbookWriter(chart_data).xlsx_blob, package)
        rId = chart_part.relate_to(xlsx_part, RT.PACKAGE)
        chartData = chart_part.chartSpace.get_or_add_chartData()
        externalData = chartData.get_or_add_externalData()
        externalData.rId = rId
        externalData.set(qn("cx:autoUpdate"), "0")
        return chart_part

    @property
    def chartex(self) -> ChartEx:
        """|ChartEx| object giving read access to the chart in this part."""
        return ChartEx(self.chartSpace, self)

    @property
    def chartSpace(self) -> CT_ChartExSpace:
        """The `cx:chartSpace` root element of this part (parsed on first access)."""
        return cast("CT_ChartExSpace", self._element)

    @property
    def xlsx_part(self) -> EmbeddedXlsxPart | None:
        """The embedded workbook holding this chart's data, or |None| if it has none."""
        rId = self.chartSpace.xlsx_part_rId
        return None if rId is None else cast(EmbeddedXlsxPart, self.related_part(rId))
