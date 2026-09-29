"""Builds a `.pptx` holding a 3-D clustered column chart, written the way PowerPoint saves one.

The chart part carries what PowerPoint 2016 writes and pypptx's own writer does not: a
`c:lang`, an `mc:AlternateContent` chart style, formatted floor and walls (`c:spPr` with
`a:sp3d`), series `c:spPr`/`c:invertIfNegative`/`c:extLst`, a `c:dLbls` block, the third
`c:axId val="0"` PowerPoint adds to a clustered 3D bar plot, and a `c:externalData` link to
the embedded workbook. Reading and editing such a chart must keep all of it.

This is hand-built to PowerPoint's layout, not saved by PowerPoint.
"""

from __future__ import annotations

import zipfile
from io import BytesIO

CATEGORIES = ("Q1", "Q2", "Q3")
SERIES = (("Revenue", (12.0, 15.5, 19.0)), ("Cost", (8.0, 9.25, 11.0)))


def _str_cache(ref: str, values: tuple[str, ...]) -> str:
    pts = "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, v) for i, v in enumerate(values))
    return '<c:strRef><c:f>%s</c:f><c:strCache><c:ptCount val="%d"/>%s</c:strCache></c:strRef>' % (
        ref,
        len(values),
        pts,
    )


def _ser(idx: int, name: str, values: tuple[float, ...]) -> str:
    col = "BC"[idx]
    pts = "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, v) for i, v in enumerate(values))
    return (
        '<c:ser><c:idx val="%d"/><c:order val="%d"/>' % (idx, idx)
        + "<c:tx>%s</c:tx>" % _str_cache("Sheet1!$%s$1" % col, (name,))
        + '<c:spPr><a:solidFill><a:schemeClr val="accent%d"/></a:solidFill>' % (idx + 1)
        + "<a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/></c:spPr>"
        + '<c:invertIfNegative val="0"/>'
        + "<c:cat>%s</c:cat>" % _str_cache("Sheet1!$A$2:$A$4", CATEGORIES)
        + "<c:val><c:numRef><c:f>Sheet1!$%s$2:$%s$4</c:f><c:numCache>" % (col, col)
        + '<c:formatCode>General</c:formatCode><c:ptCount val="%d"/>%s' % (len(values), pts)
        + "</c:numCache></c:numRef></c:val>"
        + '<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" '
        'xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">'
        '<c16:uniqueId val="{0000000%d-0001-0000-0000-000000000000}"/></c:ext></c:extLst>' % idx
        + "</c:ser>"
    )


def _surface(tag: str) -> str:
    return (
        '<c:%s><c:thickness val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln>'
        "<a:effectLst/><a:sp3d/></c:spPr></c:%s>" % (tag, tag)
    )


CHART_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>'
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    '<mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">'
    '<c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback>'
    "</mc:AlternateContent>"
    '<c:chart><c:autoTitleDeleted val="1"/>'
    '<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:depthPercent val="100"/>'
    '<c:rAngAx val="1"/></c:view3D>'
    + _surface("floor")
    + _surface("sideWall")
    + _surface("backWall")
    + '<c:plotArea><c:layout/><c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>'
    '<c:varyColors val="0"/>'
    + "".join(_ser(i, name, values) for i, (name, values) in enumerate(SERIES))
    + '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>'
    '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
    '<c:gapWidth val="150"/><c:shape val="box"/>'
    '<c:axId val="510491184"/><c:axId val="510494464"/><c:axId val="0"/></c:bar3DChart>'
    '<c:catAx><c:axId val="510491184"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
    '<c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/>'
    '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
    '<c:crossAx val="510494464"/><c:crosses val="autoZero"/><c:auto val="1"/>'
    '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
    '<c:valAx><c:axId val="510494464"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
    '<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>'
    '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>'
    '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="510491184"/>'
    '<c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>'
    '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr></c:plotArea>'
    '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/>'
    "</a:p></c:txPr>"
    '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'
    "</c:chartSpace>"
)


def chart_3d_pptx() -> BytesIO:
    """Return a stream holding a one-slide `.pptx` whose one chart is `CHART_XML` above."""
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    chart_data = CategoryChartData()
    chart_data.categories = CATEGORIES
    for name, values in SERIES:
        chart_data.add_series(name, values)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(6), Inches(4), chart_data
    )
    base = BytesIO()
    prs.save(base)

    out = BytesIO()
    with zipfile.ZipFile(base) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "ppt/charts/chart1.xml":
                data = CHART_XML.encode("utf-8")
            dst.writestr(item, data)
    out.seek(0)
    return out
