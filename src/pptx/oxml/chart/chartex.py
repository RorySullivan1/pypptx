"""Custom element classes for Office 2016+ "chartex" charts (the `cx:` namespace).

Waterfall, histogram, Pareto, box & whisker, treemap, sunburst, funnel and region-map charts are
stored in a chartex part (`cx:chartSpace`) rather than a classic chart part (`c:chartSpace`).
The data lives in `cx:chartSpace/cx:chartData/cx:data`, one `cx:data` per series, each holding a
category dimension (`cx:strDim` or `cx:numDim` of type "cat") and a value dimension (a
`cx:numDim`); the series in `cx:chart/cx:plotArea/cx:plotAreaRegion` point at their data by id.
"""

from __future__ import annotations

from typing import cast

from pptx.oxml.simpletypes import XsdBoolean, XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_ChartExSpace(BaseOxmlElement):
    """`cx:chartSpace` element, root of a chartex part."""

    chartData: CT_ChartExData | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:chartData", successors=("cx:chart", "cx:spPr", "cx:txPr", "cx:clrMapOvr")
    )
    chart: CT_ChartExChart | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:chart", successors=("cx:spPr", "cx:txPr", "cx:clrMapOvr")
    )

    @property
    def series_lst(self) -> list[CT_ChartExSeries]:
        """Every `cx:series` of this chart, in document order."""
        return cast(
            "list[CT_ChartExSeries]",
            self.xpath("./cx:chart/cx:plotArea/cx:plotAreaRegion/cx:series"),
        )

    def data_for(self, id_: int) -> CT_ChartExDataItem | None:
        """The `cx:data` element with `id` of `id_`, or |None| if there is none."""
        chartData = self.chartData
        if chartData is None:
            return None
        for data in chartData.data_lst:
            if data.id == id_:
                return data
        return None

    @property
    def xlsx_part_rId(self) -> str | None:
        """rId of the embedded-workbook relationship, from `cx:chartData/cx:externalData`."""
        rIds = cast("list[str]", self.xpath("./cx:chartData/cx:externalData/@r:id"))
        return rIds[0] if rIds else None


class CT_ChartExData(BaseOxmlElement):
    """`cx:chartData` element, the data (and workbook reference) of a chartex chart."""

    data_lst: list[CT_ChartExDataItem]

    externalData: CT_ChartExExternalData | None = ZeroOrOne(  # pyright: ignore
        "cx:externalData", successors=("cx:data",)
    )
    data = ZeroOrMore("cx:data")


class CT_ChartExExternalData(BaseOxmlElement):
    """`cx:externalData` element, pointing at the embedded workbook part."""

    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_ChartExDataItem(BaseOxmlElement):
    """`cx:data` element, the category and value dimensions for one series."""

    id: int = RequiredAttribute("id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]

    @property
    def cat_dim(self) -> CT_ChartExDimension | None:
        """The category dimension (`cx:strDim` or `cx:numDim` of type "cat"), if any."""
        dims = cast(
            "list[CT_ChartExDimension]",
            self.xpath("./cx:strDim[@type='cat'] | ./cx:numDim[@type='cat']"),
        )
        return dims[0] if dims else None

    @property
    def val_dim(self) -> CT_ChartExDimension | None:
        """The value dimension: the first `cx:numDim` whose type is not "cat", if any.

        Its type is "val" for most charts and "size" for treemap and sunburst.
        """
        dims = cast("list[CT_ChartExDimension]", self.xpath("./cx:numDim[@type!='cat']"))
        return dims[0] if dims else None


class CT_ChartExDimension(BaseOxmlElement):
    """`cx:strDim` or `cx:numDim` element, one dimension of a series' data.

    Holds the workbook formula (`cx:f`) and one cached level (`cx:lvl`) per hierarchy level. For
    hierarchical categories (treemap, sunburst) the first level is the leaf level and each
    following level is the next one up.
    """

    lvl_lst: list[CT_ChartExLevel]

    type: str = RequiredAttribute("type", XsdString)  # pyright: ignore[reportAssignmentType]
    lvl = ZeroOrMore("cx:lvl", successors=("cx:extLst",))

    @property
    def formula(self) -> str | None:
        """Text of the `cx:f` child, e.g. "Sheet1!$B$2:$B$7", or |None| if there is none."""
        fs = cast("list[str]", self.xpath("./cx:f/text()"))
        return fs[0] if fs else None


class CT_ChartExLevel(BaseOxmlElement):
    """`cx:lvl` element, the cached points of one level of a dimension."""

    pt_lst: list[CT_ChartExPoint]

    ptCount: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "ptCount", XsdUnsignedInt
    )
    formatCode: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "formatCode", XsdString
    )
    pt = ZeroOrMore("cx:pt")

    def texts(self) -> list[str | None]:
        """The text of each point, by index; |None| where the level has no point."""
        texts: list[str | None] = [None] * self.ptCount
        for pt in self.pt_lst:
            if 0 <= pt.idx < self.ptCount:
                texts[pt.idx] = pt.text or ""
        return texts


class CT_ChartExPoint(BaseOxmlElement):
    """`cx:pt` element, one cached value; its text is the value."""

    idx: int = RequiredAttribute("idx", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_ChartExChart(BaseOxmlElement):
    """`cx:chart` element.

    The tag has two roles: inside a chartex part it is the chart (title, plot area, legend);
    inside a slide's `a:graphicData` it is the reference to the chartex part, carrying `r:id`.
    """

    plotArea: CT_ChartExPlotArea | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:plotArea", successors=("cx:legend", "cx:extLst")
    )
    rId: str | None = OptionalAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_ChartExPlotArea(BaseOxmlElement):
    """`cx:plotArea` element, holding the plot-area region (the series) and the axes."""

    plotAreaRegion: CT_ChartExPlotAreaRegion | None = ZeroOrOne(  # pyright: ignore
        "cx:plotAreaRegion", successors=("cx:axis", "cx:spPr", "cx:extLst")
    )


class CT_ChartExPlotAreaRegion(BaseOxmlElement):
    """`cx:plotAreaRegion` element, holding the series."""

    series_lst: list[CT_ChartExSeries]

    series = ZeroOrMore("cx:series", successors=("cx:extLst",))


class CT_ChartExSeries(BaseOxmlElement):
    """`cx:series` element, one series of a chartex chart.

    `layoutId` names the chart type the series is drawn as, e.g. "waterfall", "clusteredColumn"
    (histogram), "paretoLine", "boxWhisker", "treemap", "sunburst", "funnel" or "regionMap".
    """

    layoutId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "layoutId", XsdString
    )
    hidden: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "hidden", XsdBoolean, default=False
    )
    ownerIdx: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "ownerIdx", XsdUnsignedInt
    )

    @property
    def data_id(self) -> int | None:
        """`val` of the `cx:dataId` child, the id of this series' `cx:data`, or |None|."""
        vals = cast("list[str]", self.xpath("./cx:dataId/@val"))
        return int(vals[0]) if vals else None

    @property
    def name(self) -> str | None:
        """Cached series name, from `cx:tx/cx:txData/cx:v`, or |None| if there is none."""
        vs = cast("list[str]", self.xpath("./cx:tx/cx:txData/cx:v/text()"))
        return vs[0] if vs else None
