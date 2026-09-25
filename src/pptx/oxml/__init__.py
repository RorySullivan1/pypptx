"""Initializes lxml parser, particularly the custom element classes.

Also makes available a handful of functions that wrap its typical uses.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Type

from lxml import etree

from pptx.oxml.ns import NamespacePrefixedTag

if TYPE_CHECKING:
    from pptx.oxml.xmlchemy import BaseOxmlElement


# -- configure etree XML parser ----------------------------
element_class_lookup = etree.ElementNamespaceClassLookup()
oxml_parser = etree.XMLParser(remove_blank_text=True, resolve_entities=False)
oxml_parser.set_element_class_lookup(element_class_lookup)


def parse_from_template(template_file_name: str):
    """Return an element loaded from the XML in the template file identified by `template_name`."""
    thisdir = os.path.split(__file__)[0]
    filename = os.path.join(thisdir, "..", "templates", "%s.xml" % template_file_name)
    with open(filename, "rb") as f:
        xml = f.read()
    return parse_xml(xml)


def parse_xml(xml: str | bytes):
    """Return root lxml element obtained by parsing XML character string in `xml`."""
    return etree.fromstring(xml, oxml_parser)


def register_element_cls(nsptagname: str, cls: Type[BaseOxmlElement]):
    """Register `cls` to be constructed when oxml parser encounters element having `nsptag_name`.

    `nsptag_name` is a string of the form `nspfx:tagroot`, e.g. `"w:document"`.
    """
    nsptag = NamespacePrefixedTag(nsptagname)
    namespace = element_class_lookup.get_namespace(nsptag.nsuri)
    namespace[nsptag.local_part] = cls


from pptx.oxml.action import CT_Hyperlink  # noqa: E402

register_element_cls("a:hlinkClick", CT_Hyperlink)
register_element_cls("a:hlinkHover", CT_Hyperlink)


from pptx.oxml.chart.axis import (  # noqa: E402
    CT_AxPos,
    CT_AxisUnit,
    CT_BuiltInUnit,
    CT_CatAx,
    CT_ChartLines,
    CT_CrossBetween,
    CT_Crosses,
    CT_DateAx,
    CT_DispUnits,
    CT_LblAlgn,
    CT_LblOffset,
    CT_Orientation,
    CT_Scaling,
    CT_TickLblPos,
    CT_TickMark,
    CT_TimeUnit,
    CT_ValAx,
)

register_element_cls("c:axPos", CT_AxPos)
register_element_cls("c:baseTimeUnit", CT_TimeUnit)
register_element_cls("c:builtInUnit", CT_BuiltInUnit)
register_element_cls("c:catAx", CT_CatAx)
register_element_cls("c:crossBetween", CT_CrossBetween)
register_element_cls("c:crosses", CT_Crosses)
register_element_cls("c:dateAx", CT_DateAx)
register_element_cls("c:dispUnits", CT_DispUnits)
register_element_cls("c:lblAlgn", CT_LblAlgn)
register_element_cls("c:lblOffset", CT_LblOffset)
register_element_cls("c:majorGridlines", CT_ChartLines)
register_element_cls("c:majorTickMark", CT_TickMark)
register_element_cls("c:majorTimeUnit", CT_TimeUnit)
register_element_cls("c:majorUnit", CT_AxisUnit)
register_element_cls("c:minorGridlines", CT_ChartLines)
register_element_cls("c:minorTickMark", CT_TickMark)
register_element_cls("c:minorTimeUnit", CT_TimeUnit)
register_element_cls("c:minorUnit", CT_AxisUnit)
register_element_cls("c:orientation", CT_Orientation)
register_element_cls("c:scaling", CT_Scaling)
register_element_cls("c:tickLblPos", CT_TickLblPos)
register_element_cls("c:valAx", CT_ValAx)

# Reuse CT_ChartLines for dropLines, hiLowLines, serLines (all just spPr containers)
register_element_cls("c:dropLines", CT_ChartLines)
register_element_cls("c:hiLowLines", CT_ChartLines)
register_element_cls("c:leaderLines", CT_ChartLines)
register_element_cls("c:serLines", CT_ChartLines)


from pptx.oxml.chart.chartlines import CT_UpDownBar, CT_UpDownBars  # noqa: E402

register_element_cls("c:downBars", CT_UpDownBar)
register_element_cls("c:upBars", CT_UpDownBar)
register_element_cls("c:upDownBars", CT_UpDownBars)


from pptx.oxml.chart.chart import (  # noqa: E402
    CT_Chart,
    CT_ChartSpace,
    CT_DispBlanksAs,
    CT_ExternalData,
    CT_PlotArea,
    CT_Style,
)

register_element_cls("c:chart", CT_Chart)
register_element_cls("c:chartSpace", CT_ChartSpace)
register_element_cls("c:dispBlanksAs", CT_DispBlanksAs)
register_element_cls("c:externalData", CT_ExternalData)
register_element_cls("c:plotArea", CT_PlotArea)
register_element_cls("c:style", CT_Style)


from pptx.oxml.chart.datatable import CT_DTable  # noqa: E402
from pptx.oxml.chart.shared import CT_Boolean as _CT_Boolean  # noqa: E402

register_element_cls("c:dTable", CT_DTable)
register_element_cls("c:showHorzBorder", _CT_Boolean)
register_element_cls("c:showVertBorder", _CT_Boolean)
register_element_cls("c:showOutline", _CT_Boolean)
register_element_cls("c:showKeys", _CT_Boolean)


from pptx.oxml.chart.datalabel import CT_DLbl, CT_DLblPos, CT_DLbls  # noqa: E402

register_element_cls("c:dLbl", CT_DLbl)
register_element_cls("c:dLblPos", CT_DLblPos)
register_element_cls("c:dLbls", CT_DLbls)


from pptx.oxml.chart.legend import CT_Legend, CT_LegendPos  # noqa: E402

register_element_cls("c:legend", CT_Legend)
register_element_cls("c:legendPos", CT_LegendPos)


from pptx.oxml.chart.marker import CT_Marker, CT_MarkerSize, CT_MarkerStyle  # noqa: E402

register_element_cls("c:marker", CT_Marker)
register_element_cls("c:size", CT_MarkerSize)
register_element_cls("c:symbol", CT_MarkerStyle)


from pptx.oxml.chart.errbar import (  # noqa: E402
    CT_ErrBarType,
    CT_ErrBars,
    CT_ErrDir,
    CT_ErrValType,
)

register_element_cls("c:errBarType", CT_ErrBarType)
register_element_cls("c:errBars", CT_ErrBars)
register_element_cls("c:errDir", CT_ErrDir)
register_element_cls("c:errValType", CT_ErrValType)


from pptx.oxml.chart.trendline import (  # noqa: E402
    CT_Trendline,
    CT_TrendlineLabel,
    CT_TrendlineType,
)

register_element_cls("c:trendline", CT_Trendline)
register_element_cls("c:trendlineLbl", CT_TrendlineLabel)
register_element_cls("c:trendlineType", CT_TrendlineType)


from pptx.oxml.chart.plot import (  # noqa: E402
    CT_Area3DChart,
    CT_AreaChart,
    CT_Bar3DChart,
    CT_BarChart,
    CT_BarDir,
    CT_BubbleChart,
    CT_BubbleScale,
    CT_DoughnutChart,
    CT_FirstSliceAng,
    CT_GapAmount,
    CT_Grouping,
    CT_HoleSize,
    CT_Line3DChart,
    CT_LineChart,
    CT_OfPieChart,
    CT_OfPieType,
    CT_Overlap,
    CT_Pie3DChart,
    CT_PieChart,
    CT_RadarChart,
    CT_RadarStyle,
    CT_ScatterChart,
    CT_ScatterStyle,
    CT_SecondPieSize,
    CT_Shape,
    CT_SizeRepresents,
    CT_SplitType,
    CT_StockChart,
    CT_Surface3DChart,
    CT_SurfaceChart,
)

register_element_cls("c:area3DChart", CT_Area3DChart)
register_element_cls("c:areaChart", CT_AreaChart)
register_element_cls("c:bar3DChart", CT_Bar3DChart)
register_element_cls("c:barChart", CT_BarChart)
register_element_cls("c:barDir", CT_BarDir)
register_element_cls("c:bubbleChart", CT_BubbleChart)
register_element_cls("c:bubbleScale", CT_BubbleScale)
register_element_cls("c:doughnutChart", CT_DoughnutChart)
register_element_cls("c:firstSliceAng", CT_FirstSliceAng)
register_element_cls("c:gapDepth", CT_GapAmount)
register_element_cls("c:gapWidth", CT_GapAmount)
register_element_cls("c:grouping", CT_Grouping)
register_element_cls("c:holeSize", CT_HoleSize)
register_element_cls("c:line3DChart", CT_Line3DChart)
register_element_cls("c:lineChart", CT_LineChart)
register_element_cls("c:ofPieChart", CT_OfPieChart)
register_element_cls("c:ofPieType", CT_OfPieType)
register_element_cls("c:overlap", CT_Overlap)
register_element_cls("c:pie3DChart", CT_Pie3DChart)
register_element_cls("c:pieChart", CT_PieChart)
register_element_cls("c:radarChart", CT_RadarChart)
register_element_cls("c:radarStyle", CT_RadarStyle)
register_element_cls("c:scatterChart", CT_ScatterChart)
register_element_cls("c:scatterStyle", CT_ScatterStyle)
register_element_cls("c:secondPieSize", CT_SecondPieSize)
register_element_cls("c:shape", CT_Shape)
register_element_cls("c:sizeRepresents", CT_SizeRepresents)
register_element_cls("c:splitType", CT_SplitType)
register_element_cls("c:stockChart", CT_StockChart)
register_element_cls("c:surface3DChart", CT_Surface3DChart)
register_element_cls("c:surfaceChart", CT_SurfaceChart)


from pptx.oxml.chart.series import (  # noqa: E402
    CT_AxDataSource,
    CT_DPt,
    CT_Lvl,
    CT_NumDataSource,
    CT_SeriesComposite,
    CT_StrVal_NumVal_Composite,
)

register_element_cls("c:bubbleSize", CT_NumDataSource)
register_element_cls("c:cat", CT_AxDataSource)
register_element_cls("c:dPt", CT_DPt)
register_element_cls("c:lvl", CT_Lvl)
register_element_cls("c:pt", CT_StrVal_NumVal_Composite)
register_element_cls("c:ser", CT_SeriesComposite)
register_element_cls("c:val", CT_NumDataSource)
register_element_cls("c:xVal", CT_NumDataSource)
register_element_cls("c:yVal", CT_NumDataSource)


from pptx.oxml.chart.shared import (  # noqa: E402
    CT_Boolean,
    CT_Boolean_Explicit,
    CT_Double,
    CT_Layout,
    CT_LayoutMode,
    CT_ManualLayout,
    CT_NumFmt,
    CT_Title,
    CT_Tx,
    CT_UnsignedInt,
)

register_element_cls("c:axId", CT_UnsignedInt)
register_element_cls("c:auto", CT_Boolean)
register_element_cls("c:autoTitleDeleted", CT_Boolean_Explicit)
register_element_cls("c:autoUpdate", CT_Boolean)
register_element_cls("c:backward", CT_Double)
register_element_cls("c:bubble3D", CT_Boolean)
register_element_cls("c:crossAx", CT_UnsignedInt)
register_element_cls("c:crossesAt", CT_Double)
register_element_cls("c:date1904", CT_Boolean)
register_element_cls("c:delete", CT_Boolean)
register_element_cls("c:dispEq", CT_Boolean)
register_element_cls("c:explosion", CT_UnsignedInt)
register_element_cls("c:dispRSqr", CT_Boolean)
register_element_cls("c:forward", CT_Double)
register_element_cls("c:idx", CT_UnsignedInt)
register_element_cls("c:intercept", CT_Double)
register_element_cls("c:invertIfNegative", CT_Boolean_Explicit)
register_element_cls("c:layout", CT_Layout)
register_element_cls("c:manualLayout", CT_ManualLayout)
register_element_cls("c:logBase", CT_Double)
register_element_cls("c:max", CT_Double)
register_element_cls("c:min", CT_Double)
register_element_cls("c:noEndCap", CT_Boolean)
register_element_cls("c:noMultiLvlLbl", CT_Boolean)
register_element_cls("c:numFmt", CT_NumFmt)
register_element_cls("c:order", CT_UnsignedInt)
register_element_cls("c:overlay", CT_Boolean_Explicit)
register_element_cls("c:period", CT_UnsignedInt)
register_element_cls("c:plotVisOnly", CT_Boolean)
register_element_cls("c:roundedCorners", CT_Boolean)
register_element_cls("c:ptCount", CT_UnsignedInt)
register_element_cls("c:showCatName", CT_Boolean_Explicit)
register_element_cls("c:showLegendKey", CT_Boolean_Explicit)
register_element_cls("c:showPercent", CT_Boolean_Explicit)
register_element_cls("c:showSerName", CT_Boolean_Explicit)
register_element_cls("c:showBubbleSize", CT_Boolean_Explicit)
register_element_cls("c:showDLblsOverMax", CT_Boolean)
register_element_cls("c:showLeaderLines", CT_Boolean)
register_element_cls("c:showNegBubbles", CT_Boolean)
register_element_cls("c:showVal", CT_Boolean_Explicit)
register_element_cls("c:smooth", CT_Boolean)
register_element_cls("c:splitPos", CT_Double)
register_element_cls("c:tickLblSkip", CT_UnsignedInt)
register_element_cls("c:tickMarkSkip", CT_UnsignedInt)
register_element_cls("c:title", CT_Title)
register_element_cls("c:tx", CT_Tx)
register_element_cls("c:varyColors", CT_Boolean)
register_element_cls("c:wireframe", CT_Boolean)
register_element_cls("c:h", CT_Double)
register_element_cls("c:hMode", CT_LayoutMode)
register_element_cls("c:w", CT_Double)
register_element_cls("c:wMode", CT_LayoutMode)
register_element_cls("c:x", CT_Double)
register_element_cls("c:xMode", CT_LayoutMode)
register_element_cls("c:y", CT_Double)
register_element_cls("c:yMode", CT_LayoutMode)


from pptx.oxml.chart.view3d import (  # noqa: E402
    CT_DepthPercent,
    CT_HPercent,
    CT_Perspective,
    CT_RotX,
    CT_RotY,
    CT_Surface,
    CT_Thickness,
    CT_View3D,
)

register_element_cls("c:backWall", CT_Surface)
register_element_cls("c:depthPercent", CT_DepthPercent)
register_element_cls("c:floor", CT_Surface)
register_element_cls("c:hPercent", CT_HPercent)
register_element_cls("c:perspective", CT_Perspective)
register_element_cls("c:rAngAx", CT_Boolean)
register_element_cls("c:rotX", CT_RotX)
register_element_cls("c:rotY", CT_RotY)
register_element_cls("c:sideWall", CT_Surface)
register_element_cls("c:thickness", CT_Thickness)
register_element_cls("c:view3D", CT_View3D)


from pptx.oxml.coreprops import CT_CoreProperties  # noqa: E402

register_element_cls("cp:coreProperties", CT_CoreProperties)


from pptx.oxml.dml.color import (  # noqa: E402
    CT_Color,
    CT_HslColor,
    CT_Percentage,
    CT_PresetColor,
    CT_SchemeColor,
    CT_ScRgbColor,
    CT_SRgbColor,
    CT_SystemColor,
)

register_element_cls("a:bgClr", CT_Color)
register_element_cls("a:clrFrom", CT_Color)
register_element_cls("a:clrTo", CT_Color)
register_element_cls("a:fgClr", CT_Color)
register_element_cls("a:hslClr", CT_HslColor)
register_element_cls("a:alpha", CT_Percentage)
register_element_cls("a:lumMod", CT_Percentage)
register_element_cls("a:lumOff", CT_Percentage)
register_element_cls("a:satMod", CT_Percentage)
register_element_cls("a:satOff", CT_Percentage)
register_element_cls("a:shade", CT_Percentage)
register_element_cls("a:tint", CT_Percentage)
register_element_cls("a:prstClr", CT_PresetColor)
register_element_cls("a:schemeClr", CT_SchemeColor)
register_element_cls("a:scrgbClr", CT_ScRgbColor)
register_element_cls("a:srgbClr", CT_SRgbColor)
register_element_cls("a:sysClr", CT_SystemColor)


from pptx.oxml.dml.fill import (  # noqa: E402
    CT_Blip,
    CT_BlipFillProperties,
    CT_ColorChangeEffect,
    CT_GradientFillProperties,
    CT_GradientStop,
    CT_GradientStopList,
    CT_GroupFillProperties,
    CT_LinearShadeProperties,
    CT_NoFillProperties,
    CT_PatternFillProperties,
    CT_RelativeRect,
    CT_SolidColorFillProperties,
    CT_StretchInfoProperties,
    CT_TileInfoProperties,
)

register_element_cls("a:blip", CT_Blip)
register_element_cls("a:blipFill", CT_BlipFillProperties)
register_element_cls("a:clrChange", CT_ColorChangeEffect)
register_element_cls("a:fillRect", CT_RelativeRect)
register_element_cls("a:fillToRect", CT_RelativeRect)
register_element_cls("a:gradFill", CT_GradientFillProperties)
register_element_cls("a:grpFill", CT_GroupFillProperties)
register_element_cls("a:gs", CT_GradientStop)
register_element_cls("a:gsLst", CT_GradientStopList)
register_element_cls("a:lin", CT_LinearShadeProperties)
register_element_cls("a:noFill", CT_NoFillProperties)
register_element_cls("a:pattFill", CT_PatternFillProperties)
register_element_cls("a:solidFill", CT_SolidColorFillProperties)
register_element_cls("a:srcRect", CT_RelativeRect)
register_element_cls("a:stretch", CT_StretchInfoProperties)
register_element_cls("a:tile", CT_TileInfoProperties)


from pptx.oxml.dml.effect import (  # noqa: E402
    CT_EffectList,
    CT_GlowEffect,
    CT_InnerShadowEffect,
    CT_OuterShadowEffect,
    CT_ReflectionEffect,
    CT_SoftEdgesEffect,
)

register_element_cls("a:effectLst", CT_EffectList)
register_element_cls("a:glow", CT_GlowEffect)
register_element_cls("a:innerShdw", CT_InnerShadowEffect)
register_element_cls("a:outerShdw", CT_OuterShadowEffect)
register_element_cls("a:reflection", CT_ReflectionEffect)
register_element_cls("a:softEdge", CT_SoftEdgesEffect)


from pptx.oxml.dml.line import CT_PresetLineDashProperties  # noqa: E402

register_element_cls("a:prstDash", CT_PresetLineDashProperties)


from pptx.oxml.dml.picture import (  # noqa: E402
    CT_DuotoneEffect,
    CT_GrayscaleEffect,
    CT_LuminanceEffect,
)

register_element_cls("a:duotone", CT_DuotoneEffect)
register_element_cls("a:grayscl", CT_GrayscaleEffect)
register_element_cls("a:lum", CT_LuminanceEffect)


from pptx.oxml.dml.threed import (  # noqa: E402
    CT_Bevel,
    CT_Camera,
    CT_LightRig,
    CT_Scene3D,
    CT_Shape3D,
)

register_element_cls("a:bevelB", CT_Bevel)
register_element_cls("a:bevelT", CT_Bevel)
register_element_cls("a:camera", CT_Camera)
register_element_cls("a:lightRig", CT_LightRig)
register_element_cls("a:scene3d", CT_Scene3D)
register_element_cls("a:sp3d", CT_Shape3D)


from pptx.oxml.presentation import (  # noqa: E402
    CT_CustomShow,
    CT_CustomShowList,
    CT_CustomShowSlideList,
    CT_NotesSize,
    CT_Presentation,
    CT_SlideId,
    CT_SlideIdList,
    CT_SlideMasterIdList,
    CT_SlideMasterIdListEntry,
    CT_SlideSize,
)

register_element_cls("p:presentation", CT_Presentation)
register_element_cls("p:custShow", CT_CustomShow)
register_element_cls("p:custShowLst", CT_CustomShowList)
register_element_cls("p:notesSz", CT_NotesSize)
register_element_cls("p:sldId", CT_SlideId)
register_element_cls("p:sldIdLst", CT_SlideIdList)
register_element_cls("p:sldLst", CT_CustomShowSlideList)
register_element_cls("p:sldMasterId", CT_SlideMasterIdListEntry)
register_element_cls("p:sldMasterIdLst", CT_SlideMasterIdList)
register_element_cls("p:sldSz", CT_SlideSize)


from pptx.oxml.embeddedfont import (  # noqa: E402
    CT_EmbeddedFontDataId,
    CT_EmbeddedFontList,
    CT_EmbeddedFontListEntry,
    CT_Font,
)

register_element_cls("p:embeddedFontLst", CT_EmbeddedFontList)
register_element_cls("p:embeddedFont", CT_EmbeddedFontListEntry)
register_element_cls("p:font", CT_Font)
register_element_cls("p:regular", CT_EmbeddedFontDataId)
register_element_cls("p:bold", CT_EmbeddedFontDataId)
register_element_cls("p:italic", CT_EmbeddedFontDataId)
register_element_cls("p:boldItalic", CT_EmbeddedFontDataId)


from pptx.oxml.shapes.autoshape import (  # noqa: E402
    CT_AdjPoint2D,
    CT_CustomGeometry2D,
    CT_GeomGuide,
    CT_GeomGuideList,
    CT_NonVisualDrawingShapeProps,
    CT_Path2D,
    CT_Path2DArcTo,
    CT_Path2DClose,
    CT_Path2DCubicBezierTo,
    CT_Path2DLineTo,
    CT_Path2DList,
    CT_Path2DMoveTo,
    CT_Path2DQuadBezierTo,
    CT_PresetGeometry2D,
    CT_Shape,
    CT_ShapeNonVisual,
)

register_element_cls("a:avLst", CT_GeomGuideList)
register_element_cls("a:custGeom", CT_CustomGeometry2D)
register_element_cls("a:gd", CT_GeomGuide)
register_element_cls("a:arcTo", CT_Path2DArcTo)
register_element_cls("a:close", CT_Path2DClose)
register_element_cls("a:cubicBezTo", CT_Path2DCubicBezierTo)
register_element_cls("a:lnTo", CT_Path2DLineTo)
register_element_cls("a:moveTo", CT_Path2DMoveTo)
register_element_cls("a:path", CT_Path2D)
register_element_cls("a:pathLst", CT_Path2DList)
register_element_cls("a:quadBezTo", CT_Path2DQuadBezierTo)
register_element_cls("a:prstGeom", CT_PresetGeometry2D)
register_element_cls("a:pt", CT_AdjPoint2D)
register_element_cls("p:cNvSpPr", CT_NonVisualDrawingShapeProps)
register_element_cls("p:nvSpPr", CT_ShapeNonVisual)
register_element_cls("p:sp", CT_Shape)


from pptx.oxml.shapes.connector import (  # noqa: E402
    CT_Connection,
    CT_Connector,
    CT_ConnectorNonVisual,
    CT_NonVisualConnectorProperties,
)

register_element_cls("a:endCxn", CT_Connection)
register_element_cls("a:stCxn", CT_Connection)
register_element_cls("p:cNvCxnSpPr", CT_NonVisualConnectorProperties)
register_element_cls("p:cxnSp", CT_Connector)
register_element_cls("p:nvCxnSpPr", CT_ConnectorNonVisual)


from pptx.oxml.shapes.graphfrm import (  # noqa: E402
    CT_GraphicalObject,
    CT_GraphicalObjectData,
    CT_GraphicalObjectFrame,
    CT_GraphicalObjectFrameNonVisual,
    CT_OleObject,
)

register_element_cls("a:graphic", CT_GraphicalObject)
register_element_cls("a:graphicData", CT_GraphicalObjectData)
register_element_cls("p:graphicFrame", CT_GraphicalObjectFrame)
register_element_cls("p:nvGraphicFramePr", CT_GraphicalObjectFrameNonVisual)
register_element_cls("p:oleObj", CT_OleObject)


from pptx.oxml.shapes.groupshape import (  # noqa: E402
    CT_GroupShape,
    CT_GroupShapeNonVisual,
    CT_GroupShapeProperties,
)

register_element_cls("p:grpSp", CT_GroupShape)
register_element_cls("p:grpSpPr", CT_GroupShapeProperties)
register_element_cls("p:nvGrpSpPr", CT_GroupShapeNonVisual)
register_element_cls("p:spTree", CT_GroupShape)


from pptx.oxml.shapes.picture import CT_Picture, CT_PictureNonVisual  # noqa: E402

register_element_cls("p:blipFill", CT_BlipFillProperties)
register_element_cls("p:nvPicPr", CT_PictureNonVisual)
register_element_cls("p:pic", CT_Picture)


from pptx.oxml.shapes.shared import (  # noqa: E402
    CT_ApplicationNonVisualDrawingProps,
    CT_LineEndProperties,
    CT_LineJoinMiterProperties,
    CT_LineProperties,
    CT_Locking,
    CT_NonVisualDrawingProps,
    CT_Placeholder,
    CT_Point2D,
    CT_PositiveSize2D,
    CT_ShapeProperties,
    CT_Transform2D,
)

register_element_cls("a:chExt", CT_PositiveSize2D)
register_element_cls("a:chOff", CT_Point2D)
register_element_cls("a:ext", CT_PositiveSize2D)
register_element_cls("a:headEnd", CT_LineEndProperties)
register_element_cls("a:miter", CT_LineJoinMiterProperties)
register_element_cls("a:ln", CT_LineProperties)
register_element_cls("a:lnB", CT_LineProperties)
register_element_cls("a:lnL", CT_LineProperties)
register_element_cls("a:lnR", CT_LineProperties)
register_element_cls("a:lnT", CT_LineProperties)
register_element_cls("a:tailEnd", CT_LineEndProperties)
register_element_cls("a:off", CT_Point2D)
register_element_cls("p188:pos", CT_Point2D)
register_element_cls("a:xfrm", CT_Transform2D)
register_element_cls("c:spPr", CT_ShapeProperties)
register_element_cls("a:cxnSpLocks", CT_Locking)
register_element_cls("a:graphicFrameLocks", CT_Locking)
register_element_cls("a:grpSpLocks", CT_Locking)
register_element_cls("a:picLocks", CT_Locking)
register_element_cls("a:spLocks", CT_Locking)
register_element_cls("p:cNvPr", CT_NonVisualDrawingProps)
register_element_cls("p:nvPr", CT_ApplicationNonVisualDrawingProps)
register_element_cls("p:ph", CT_Placeholder)
register_element_cls("p:spPr", CT_ShapeProperties)
register_element_cls("p:xfrm", CT_Transform2D)


from pptx.oxml.comment import (  # noqa: E402
    CT_Comment,
    CT_CommentAuthor,
    CT_CommentAuthorList,
    CT_CommentList,
    CT_ModernAuthor,
    CT_ModernAuthorList,
    CT_ModernComment,
    CT_ModernCommentList,
    CT_ModernCommentReply,
    CT_ModernCommentReplyList,
    CT_SlideMoniker,
    CT_SlideMonikerList,
)

register_element_cls("p:cm", CT_Comment)
register_element_cls("p:cmAuthor", CT_CommentAuthor)
register_element_cls("p:cmAuthorLst", CT_CommentAuthorList)
register_element_cls("p:cmLst", CT_CommentList)
register_element_cls("p188:author", CT_ModernAuthor)
register_element_cls("p188:authorLst", CT_ModernAuthorList)
register_element_cls("p188:cm", CT_ModernComment)
register_element_cls("p188:cmLst", CT_ModernCommentList)
register_element_cls("p188:reply", CT_ModernCommentReply)
register_element_cls("p188:replyLst", CT_ModernCommentReplyList)
register_element_cls("pc:sldMk", CT_SlideMoniker)
register_element_cls("pc:sldMkLst", CT_SlideMonikerList)


from pptx.oxml.section import (  # noqa: E402
    CT_Section,
    CT_SectionList,
    CT_SectionSlideIdListEntry,
)

register_element_cls("p14:section", CT_Section)
register_element_cls("p14:sectionLst", CT_SectionList)
register_element_cls("p14:sldId", CT_SectionSlideIdListEntry)


from pptx.oxml.custprops import CT_CustomProperties, CT_CustomProperty  # noqa: E402

register_element_cls("cust:Properties", CT_CustomProperties)
register_element_cls("cust:property", CT_CustomProperty)


from pptx.oxml.tags import CT_StringTag, CT_TagList  # noqa: E402

register_element_cls("p:tag", CT_StringTag)
register_element_cls("p:tagLst", CT_TagList)


from pptx.oxml.presprops import (  # noqa: E402
    CT_IndexRange,
    CT_PresentationProperties,
    CT_ShowInfoBrowse,
    CT_ShowInfoKiosk,
    CT_ShowProperties,
)

register_element_cls("p:presentationPr", CT_PresentationProperties)
register_element_cls("p:showPr", CT_ShowProperties)
register_element_cls("p:browse", CT_ShowInfoBrowse)
register_element_cls("p:kiosk", CT_ShowInfoKiosk)
register_element_cls("p:sldRg", CT_IndexRange)
register_element_cls("p:penClr", CT_Color)


from pptx.oxml.tablestyles import CT_TableStyle, CT_TableStyleList  # noqa: E402

register_element_cls("a:tblStyle", CT_TableStyle)
register_element_cls("a:tblStyleLst", CT_TableStyleList)


from pptx.oxml.slide import (  # noqa: E402
    CT_Background,
    CT_BackgroundProperties,
    CT_CommonSlideData,
    CT_HeaderFooter,
    CT_NotesMaster,
    CT_NotesSlide,
    CT_Slide,
    CT_SlideLayout,
    CT_SlideLayoutIdList,
    CT_SlideLayoutIdListEntry,
    CT_SlideMaster,
    CT_SlideMasterTextStyles,
    CT_SlideTiming,
    CT_TimeNodeList,
    CT_TLMediaNodeVideo,
)

register_element_cls("p:bg", CT_Background)
register_element_cls("p:bgPr", CT_BackgroundProperties)
register_element_cls("p:childTnLst", CT_TimeNodeList)
register_element_cls("p:cSld", CT_CommonSlideData)
register_element_cls("p:hf", CT_HeaderFooter)
register_element_cls("p:notes", CT_NotesSlide)
register_element_cls("p:notesMaster", CT_NotesMaster)
register_element_cls("p:sld", CT_Slide)
register_element_cls("p:sldLayout", CT_SlideLayout)
register_element_cls("p:sldLayoutId", CT_SlideLayoutIdListEntry)
register_element_cls("p:sldLayoutIdLst", CT_SlideLayoutIdList)
register_element_cls("p:sldMaster", CT_SlideMaster)
register_element_cls("p:txStyles", CT_SlideMasterTextStyles)
register_element_cls("p:timing", CT_SlideTiming)
register_element_cls("p:video", CT_TLMediaNodeVideo)


from pptx.oxml.table import (  # noqa: E402
    CT_Table,
    CT_TableCell,
    CT_TableCellProperties,
    CT_TableCol,
    CT_TableGrid,
    CT_TableProperties,
    CT_TableRow,
)

register_element_cls("a:gridCol", CT_TableCol)
register_element_cls("a:tbl", CT_Table)
register_element_cls("a:tblGrid", CT_TableGrid)
register_element_cls("a:tblPr", CT_TableProperties)
register_element_cls("a:tc", CT_TableCell)
register_element_cls("a:tcPr", CT_TableCellProperties)
register_element_cls("a:tr", CT_TableRow)


from pptx.oxml.text import (  # noqa: E402
    CT_PresetTextShape,
    CT_RegularTextRun,
    CT_TabStop,
    CT_TabStopList,
    CT_TextBody,
    CT_TextBodyProperties,
    CT_TextCharacterProperties,
    CT_TextField,
    CT_TextFont,
    CT_TextLineBreak,
    CT_TextNormalAutofit,
    CT_TextParagraph,
    CT_TextListStyle,
    CT_TextParagraphProperties,
    CT_TextSpacing,
    CT_TextSpacingPercent,
    CT_TextSpacingPoint,
)

register_element_cls("a:bodyPr", CT_TextBodyProperties)
register_element_cls("a:br", CT_TextLineBreak)
register_element_cls("a:defRPr", CT_TextCharacterProperties)
register_element_cls("a:endParaRPr", CT_TextCharacterProperties)
register_element_cls("a:fld", CT_TextField)
register_element_cls("a:cs", CT_TextFont)
register_element_cls("a:ea", CT_TextFont)
register_element_cls("a:latin", CT_TextFont)
register_element_cls("a:sym", CT_TextFont)
register_element_cls("a:lnSpc", CT_TextSpacing)
register_element_cls("a:normAutofit", CT_TextNormalAutofit)
register_element_cls("a:prstTxWarp", CT_PresetTextShape)
register_element_cls("a:r", CT_RegularTextRun)
register_element_cls("a:tab", CT_TabStop)
register_element_cls("a:tabLst", CT_TabStopList)
register_element_cls("a:p", CT_TextParagraph)
register_element_cls("a:pPr", CT_TextParagraphProperties)
register_element_cls("a:defPPr", CT_TextParagraphProperties)
for _n in range(1, 10):
    register_element_cls("a:lvl%dpPr" % _n, CT_TextParagraphProperties)
del _n
register_element_cls("a:lstStyle", CT_TextListStyle)
register_element_cls("p:bodyStyle", CT_TextListStyle)
register_element_cls("p:defaultTextStyle", CT_TextListStyle)
register_element_cls("p:otherStyle", CT_TextListStyle)
register_element_cls("p:titleStyle", CT_TextListStyle)
register_element_cls("c:rich", CT_TextBody)
register_element_cls("a:rPr", CT_TextCharacterProperties)
register_element_cls("a:spcAft", CT_TextSpacing)
register_element_cls("a:spcBef", CT_TextSpacing)
register_element_cls("a:spcPct", CT_TextSpacingPercent)
register_element_cls("a:spcPts", CT_TextSpacingPoint)
register_element_cls("a:txBody", CT_TextBody)
register_element_cls("c:txPr", CT_TextBody)
register_element_cls("p:txBody", CT_TextBody)
register_element_cls("p188:txBody", CT_TextBody)


from pptx.oxml.diagram import (  # noqa: E402
    CT_CxnList,
    CT_Cxn,
    CT_DataModel,
    CT_DiagramRelIds,
    CT_ElemPropSet,
    CT_Pt,
    CT_PtList,
)

register_element_cls("dgm:cxn", CT_Cxn)
register_element_cls("dgm:cxnLst", CT_CxnList)
register_element_cls("dgm:dataModel", CT_DataModel)
register_element_cls("dgm:prSet", CT_ElemPropSet)
register_element_cls("dgm:pt", CT_Pt)
register_element_cls("dgm:ptLst", CT_PtList)
register_element_cls("dgm:relIds", CT_DiagramRelIds)
register_element_cls("dgm:t", CT_TextBody)

from pptx.oxml.theme import (  # noqa: E402
    CT_BaseStyles,
    CT_BaseStylesOverride,
    CT_ColorScheme,
    CT_EffectStyleItem,
    CT_EffectStyleList,
    CT_FillStyleList,
    CT_FontCollection,
    CT_FontScheme,
    CT_LineStyleList,
    CT_OfficeStyleSheet,
    CT_StyleMatrix,
)

register_element_cls("a:bgFillStyleLst", CT_FillStyleList)
register_element_cls("a:fillStyleLst", CT_FillStyleList)
register_element_cls("a:lnStyleLst", CT_LineStyleList)

register_element_cls("a:clrScheme", CT_ColorScheme)
register_element_cls("a:effectStyle", CT_EffectStyleItem)
register_element_cls("a:effectStyleLst", CT_EffectStyleList)
register_element_cls("a:fmtScheme", CT_StyleMatrix)
register_element_cls("a:fontScheme", CT_FontScheme)
register_element_cls("a:majorFont", CT_FontCollection)
register_element_cls("a:minorFont", CT_FontCollection)
register_element_cls("a:theme", CT_OfficeStyleSheet)
register_element_cls("a:themeOverride", CT_BaseStylesOverride)
register_element_cls("a:themeElements", CT_BaseStyles)
