"""Graphic Frame shape and related objects.

A graphic frame is a common container for table, chart, smart art, and media
objects.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.exc import ShapeError
from pptx.shapes.base import BaseShape
from pptx.shared import ParentedElementProxy
from pptx.smartart import SmartArt
from pptx.spec import (
    GRAPHIC_DATA_URI_CHART,
    GRAPHIC_DATA_URI_CHARTEX,
    GRAPHIC_DATA_URI_DIAGRAM,
    GRAPHIC_DATA_URI_OLEOBJ,
    GRAPHIC_DATA_URI_TABLE,
)
from pptx.table import Table
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.chart.chart import Chart
    from pptx.chart.chartex import ChartEx
    from pptx.dml.effect import ShadowFormat
    from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectData, CT_GraphicalObjectFrame
    from pptx.parts.chart import ChartPart
    from pptx.parts.chartex import ChartExPart
    from pptx.parts.diagram import DiagramDataPart
    from pptx.parts.slide import BaseSlidePart
    from pptx.types import ProvidesPart


class GraphicFrame(BaseShape):
    """Container shape for table, chart, SmartArt, and media objects.

    Corresponds to a `p:graphicFrame` element in the shape tree.
    """

    def __init__(self, graphicFrame: CT_GraphicalObjectFrame, parent: ProvidesPart) -> None:
        super().__init__(graphicFrame, parent)
        self._graphicFrame = graphicFrame

    @property
    def chart(self) -> Chart:
        """The |Chart| object containing the chart in this graphic frame.

        Raises |ValueError| if this graphic frame does not contain a chart.
        """
        if not self.has_chart:
            raise ShapeError("shape does not contain a chart")
        return self.chart_part.chart

    @property
    def chart_part(self) -> ChartPart:
        """The |ChartPart| object containing the chart in this graphic frame."""
        chart_rId = self._graphicFrame.chart_rId
        if chart_rId is None:
            raise ShapeError("this graphic frame does not contain a chart")
        return cast("ChartPart", self.part.related_part(chart_rId))

    @property
    def chartex(self) -> ChartEx:
        """The |ChartEx| object for the chartex chart (waterfall, treemap, ...) in this frame.

        Raises |ShapeError| if this graphic frame does not contain a chartex chart.
        """
        return self.chartex_part.chartex

    @property
    def chartex_part(self) -> ChartExPart:
        """The |ChartExPart| holding the chartex chart in this graphic frame.

        Raises |ShapeError| if this graphic frame does not contain a chartex chart.
        """
        rId = self._graphicFrame.chartex_rId
        if not self.has_chartex or rId is None:
            raise ShapeError("shape does not contain a chartex chart")
        return cast("ChartExPart", self.part.related_part(rId))

    @property
    def has_chart(self) -> bool:
        """|True| if this graphic frame contains a chart object. |False| otherwise.

        When |True|, the chart object can be accessed using the `.chart` property.
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_CHART

    @property
    def has_chartex(self) -> bool:
        """|True| if this graphic frame contains a chartex chart, |False| otherwise.

        Chartex charts are the Office 2016+ types: waterfall, histogram, Pareto, box & whisker,
        treemap, sunburst, funnel and region map. They are read through `.chartex`, not `.chart`,
        so `has_chart` is |False| for them.
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_CHARTEX

    @property
    def has_smartart(self) -> bool:
        """|True| if this graphic frame contains a SmartArt graphic, |False| otherwise.

        When |True|, the SmartArt content can be accessed using the `.smartart` property.
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_DIAGRAM

    @property
    def has_table(self) -> bool:
        """|True| if this graphic frame contains a table object, |False| otherwise.

        When |True|, the table object can be accessed using the `.table` property.
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_TABLE

    @property
    def ole_format(self) -> _OleFormat:
        """_OleFormat object for this graphic-frame shape.

        Raises `ValueError` on a GraphicFrame instance that does not contain an OLE object.

        An shape that contains an OLE object will have `.shape_type` of either
        `EMBEDDED_OLE_OBJECT` or `LINKED_OLE_OBJECT`.
        """
        if not self._graphicFrame.has_oleobj:
            raise ShapeError("not an OLE-object shape")
        return _OleFormat(self._graphicFrame.graphicData, self._parent)

    @lazyproperty
    def shadow(self) -> ShadowFormat:
        """Unconditionally raises |NotImplementedError|.

        Access to the shadow effect for graphic-frame objects is content-specific (i.e. different
        for charts, tables, etc.) and has not yet been implemented.
        """
        raise NotImplementedError("shadow property on GraphicFrame not yet supported")

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """Optional member of `MSO_SHAPE_TYPE` identifying the type of this shape.

        Possible values are `MSO_SHAPE_TYPE.CHART`, `MSO_SHAPE_TYPE.TABLE`,
        `MSO_SHAPE_TYPE.SMART_ART`, `MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT`,
        `MSO_SHAPE_TYPE.LINKED_OLE_OBJECT`. A chartex chart (see `has_chartex`) is also
        `MSO_SHAPE_TYPE.CHART`.

        This value is `None` when none of these five types apply.
        """
        graphicData_uri = self._graphicFrame.graphicData_uri
        if graphicData_uri in (GRAPHIC_DATA_URI_CHART, GRAPHIC_DATA_URI_CHARTEX):
            return MSO_SHAPE_TYPE.CHART
        elif graphicData_uri == GRAPHIC_DATA_URI_DIAGRAM:
            return MSO_SHAPE_TYPE.SMART_ART
        elif graphicData_uri == GRAPHIC_DATA_URI_TABLE:
            return MSO_SHAPE_TYPE.TABLE
        elif graphicData_uri == GRAPHIC_DATA_URI_OLEOBJ:
            return (
                MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
                if self._graphicFrame.is_embedded_ole_obj
                else MSO_SHAPE_TYPE.LINKED_OLE_OBJECT
            )
        else:
            return None  # pyright: ignore[reportReturnType]

    @property
    def smartart(self) -> SmartArt:
        """The |SmartArt| object holding the node tree of the SmartArt in this graphic frame.

        Raises |ShapeError| if this graphic frame does not contain SmartArt.
        """
        relIds = self._graphicFrame.graphicData.relIds
        if not self.has_smartart or relIds is None:
            raise ShapeError("shape does not contain SmartArt")
        data_part = cast("DiagramDataPart", self.part.related_part(relIds.dm))
        return SmartArt(data_part, self.part)

    @property
    def table(self) -> Table:
        """The |Table| object contained in this graphic frame.

        Raises |ValueError| if this graphic frame does not contain a table.
        """
        if not self.has_table:
            raise ShapeError("shape does not contain a table")
        tbl = self._graphicFrame.graphic.graphicData.tbl
        return Table(tbl, self)


class _OleFormat(ParentedElementProxy):
    """Provides attributes on an embedded OLE object."""

    part: BaseSlidePart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, graphicData: CT_GraphicalObjectData, parent: ProvidesPart) -> None:
        super().__init__(graphicData, parent)
        self._graphicData = graphicData

    @property
    def blob(self) -> bytes | None:
        """Optional bytes of OLE object, suitable for loading or saving as a file.

        This value is `None` if the embedded object does not represent a "file".
        """
        blob_rId = self._graphicData.blob_rId
        if blob_rId is None:
            return None
        return self.part.related_part(blob_rId).blob

    @property
    def prog_id(self) -> str | None:
        """str "progId" attribute of this embedded OLE object.

        The progId is a str like "Excel.Sheet.12" that identifies the "file-type" of the embedded
        object, or perhaps more precisely, the application (aka. "server" in OLE parlance) to be
        used to open this object.
        """
        return self._graphicData.progId

    @property
    def show_as_icon(self) -> bool | None:
        """True when OLE object should appear as an icon (rather than preview)."""
        return self._graphicData.showAsIcon
