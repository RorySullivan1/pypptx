"""Builds a `.pptx` package holding chartex charts laid out the way PowerPoint saves them.

The package starts as a one-slide pypptx presentation. Three chartex parts are then written in at
the ZIP level, each with its embedded workbook, relationships, content-type overrides and an
`mc:AlternateContent` graphic frame on the slide, so loading exercises the path a PowerPoint file
does:

* a waterfall (with a title and subtotals, as PowerPoint writes them),
* a Pareto chart -- a histogram series plus a "paretoLine" series with no data of its own,
* a treemap with two-level categories whose parent labels are cached only at the first point
  of each group (blank cells below, as a hand-made worksheet has).

Each workbook holds the same numbers as the chart's cache, so tests can check one against the
other.
"""

from __future__ import annotations

import io
import re
import zipfile
from io import BytesIO

from xlsxwriter import Workbook

CX = "http://schemas.microsoft.com/office/drawing/2014/chartex"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
CX1 = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"

WATERFALL_CATEGORIES = ["Start", "Sales", "Costs", "End"]
WATERFALL_VALUES = [100.0, 45.5, -30.0, 115.5]
HISTOGRAM_VALUES = [1.0, 2.0, 2.0, 3.0, 5.0, 8.0]
TREEMAP_PATHS = [("Fruit", "Apple"), ("Fruit", "Pear"), ("Veg", "Kale")]
TREEMAP_VALUES = [30.0, 12.0, 7.0]

_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_ROOT = '<cx:chartSpace xmlns:a="%s" xmlns:r="%s" xmlns:cx="%s">' % (A, R, CX)
_EXTERNAL_DATA = '<cx:externalData r:id="rId1" cx:autoUpdate="0"/>'
_AXES = (
    '<cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis>'
    '<cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>'
)


def _pts(values) -> str:
    return "".join(
        '<cx:pt idx="%d">%s</cx:pt>' % (idx, value)
        for idx, value in enumerate(values)
        if value is not None
    )


def _num(value: float) -> str:
    return ("%f" % value).rstrip("0").rstrip(".")


WATERFALL_XML = (
    _DECL + _ROOT + "<cx:chartData>" + _EXTERNAL_DATA + '<cx:data id="0">'
    '<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$5</cx:f><cx:lvl ptCount="4">'
    + _pts(WATERFALL_CATEGORIES)
    + "</cx:lvl></cx:strDim>"
    '<cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$5</cx:f><cx:lvl ptCount="4" formatCode="General">'
    + _pts(_num(v) for v in WATERFALL_VALUES)
    + "</cx:lvl></cx:numDim></cx:data></cx:chartData>"
    '<cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>Cash flow'
    "</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion>"
    '<cx:series layoutId="waterfall" uniqueId="{0A1B2C3D-0000-4000-8000-000000000001}">'
    "<cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Flow</cx:v></cx:txData></cx:tx>"
    '<cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="0" value="1"/>'
    '</cx:dataLabels><cx:dataId val="0"/><cx:layoutPr><cx:subtotals><cx:idx val="0"/>'
    '<cx:idx val="3"/></cx:subtotals></cx:layoutPr></cx:series></cx:plotAreaRegion>'
    + _AXES
    + '</cx:plotArea><cx:legend pos="t" align="ctr" overlay="0"/></cx:chart></cx:chartSpace>'
)

PARETO_XML = (
    _DECL + _ROOT + "<cx:chartData>" + _EXTERNAL_DATA + '<cx:data id="0">'
    '<cx:numDim type="val"><cx:f>Sheet1!$A$2:$A$7</cx:f><cx:lvl ptCount="6" formatCode="General">'
    + _pts(_num(v) for v in HISTOGRAM_VALUES)
    + "</cx:lvl></cx:numDim></cx:data></cx:chartData>"
    "<cx:chart><cx:plotArea><cx:plotAreaRegion>"
    '<cx:series layoutId="clusteredColumn" uniqueId="{0A1B2C3D-0000-4000-8000-000000000002}">'
    "<cx:tx><cx:txData><cx:f>Sheet1!$A$1</cx:f><cx:v>Scores</cx:v></cx:txData></cx:tx>"
    '<cx:dataId val="0"/><cx:layoutPr><cx:binning intervalClosed="r"/></cx:layoutPr>'
    '<cx:axisId val="0"/><cx:axisId val="1"/></cx:series>'
    '<cx:series layoutId="paretoLine" ownerIdx="0"'
    ' uniqueId="{0A1B2C3D-0000-4000-8000-000000000003}"><cx:axisId val="2"/></cx:series>'
    "</cx:plotAreaRegion>"
    + _AXES
    + '<cx:axis id="2"><cx:valScaling max="1" min="0"/><cx:units unit="percentage"/>'
    "<cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>"
)

TREEMAP_XML = (
    _DECL + _ROOT + "<cx:chartData>" + _EXTERNAL_DATA + '<cx:data id="0">'
    '<cx:strDim type="cat"><cx:f dir="row">Sheet1!$A$2:$B$4</cx:f>'
    '<cx:lvl ptCount="3">'
    + _pts(leaf for _, leaf in TREEMAP_PATHS)
    + '</cx:lvl><cx:lvl ptCount="3"><cx:pt idx="0">Fruit</cx:pt><cx:pt idx="2">Veg</cx:pt>'
    "</cx:lvl></cx:strDim>"
    '<cx:numDim type="size"><cx:f>Sheet1!$C$2:$C$4</cx:f><cx:lvl ptCount="3"'
    ' formatCode="General">'
    + _pts(_num(v) for v in TREEMAP_VALUES)
    + "</cx:lvl></cx:numDim></cx:data></cx:chartData>"
    "<cx:chart><cx:plotArea><cx:plotAreaRegion>"
    '<cx:series layoutId="treemap" uniqueId="{0A1B2C3D-0000-4000-8000-000000000004}">'
    "<cx:tx><cx:txData><cx:f>Sheet1!$C$1</cx:f><cx:v>Sales</cx:v></cx:txData></cx:tx>"
    '<cx:dataLabels pos="inEnd"><cx:visibility seriesName="0" categoryName="1" value="0"/>'
    '</cx:dataLabels><cx:dataId val="0"/><cx:layoutPr><cx:parentLabelLayout val="overlapping"/>'
    "</cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"
)


def _xlsx(rows: list[list[object]]) -> bytes:
    stream = io.BytesIO()
    workbook = Workbook(stream, {"in_memory": True})
    worksheet = workbook.add_worksheet()
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            if value is not None:
                worksheet.write(r, c, value)
    workbook.close()
    return stream.getvalue()


WATERFALL_XLSX = _xlsx(
    [["", "Flow"]] + [[c, v] for c, v in zip(WATERFALL_CATEGORIES, WATERFALL_VALUES)]
)
PARETO_XLSX = _xlsx([["Scores"]] + [[v] for v in HISTOGRAM_VALUES])
TREEMAP_XLSX = _xlsx(
    [["", "", "Sales"]]
    + [
        [parent if idx == 0 or TREEMAP_PATHS[idx - 1][0] != parent else None, leaf, value]
        for idx, ((parent, leaf), value) in enumerate(zip(TREEMAP_PATHS, TREEMAP_VALUES))
    ]
)

CHARTS = (
    # -- (chart partname, chart XML, workbook partname, workbook bytes, slide rId, shape id) --
    ("ppt/charts/chartEx1.xml", WATERFALL_XML, "ppt/embeddings/cx1.xlsx", WATERFALL_XLSX, 1),
    ("ppt/charts/chartEx2.xml", PARETO_XML, "ppt/embeddings/cx2.xlsx", PARETO_XLSX, 2),
    ("ppt/charts/chartEx3.xml", TREEMAP_XML, "ppt/embeddings/cx3.xlsx", TREEMAP_XLSX, 3),
)


def _frame(n: int) -> str:
    return (
        '<mc:AlternateContent xmlns:mc="%s"><mc:Choice xmlns:cx1="%s" Requires="cx1">'
        "<p:graphicFrame><p:nvGraphicFramePr>"
        '<p:cNvPr id="%d" name="Chart %d"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
        '<p:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="2000000"/></p:xfrm>'
        '<a:graphic><a:graphicData uri="%s"><cx:chart xmlns:cx="%s" r:id="rId9%02d"/>'
        "</a:graphicData></a:graphic></p:graphicFrame></mc:Choice>"
        '<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="%d" name="Chart %d"/><p:cNvSpPr/><p:nvPr/>'
        "</p:nvSpPr><p:spPr/></p:sp></mc:Fallback></mc:AlternateContent>"
    ) % (MC, CX1, 100 + n, n, CX, CX, n, 100 + n, n)


def chartex_pptx() -> BytesIO:
    """Return a stream holding a one-slide `.pptx` whose slide has the three chartex charts."""
    from pptx import Presentation

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    base = BytesIO()
    prs.save(base)

    overrides = "".join(
        '<Override PartName="/%s" ContentType="application/vnd.ms-office.chartex+xml"/>'
        '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument'
        '.spreadsheetml.sheet"/>' % (chart_partname, xlsx_partname)
        for chart_partname, _, xlsx_partname, _, _ in CHARTS
    )
    slide_rels = "".join(
        '<Relationship Id="rId9%02d" Type="http://schemas.microsoft.com/office/2014/'
        'relationships/chartEx" Target="../charts/%s"/>' % (n, chart_partname.rsplit("/", 1)[1])
        for chart_partname, _, _, _, n in CHARTS
    )

    out = BytesIO()
    with zipfile.ZipFile(base) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = (
                src.read(item.filename).decode("utf-8")
                if item.filename.endswith((".xml", ".rels"))
                else None
            )
            if item.filename == "[Content_Types].xml":
                data = data.replace("</Types>", overrides + "</Types>")
            elif item.filename == "ppt/slides/_rels/slide1.xml.rels":
                data = data.replace("</Relationships>", slide_rels + "</Relationships>")
            elif item.filename == "ppt/slides/slide1.xml":
                frames = "".join(_frame(n) for _, _, _, _, n in CHARTS)
                data = data.replace("</p:spTree>", frames + "</p:spTree>")
            if data is None:
                dst.writestr(item, src.read(item.filename))
            else:
                dst.writestr(item, data.encode("utf-8"))
        for chart_partname, chart_xml, xlsx_partname, xlsx_blob, _ in CHARTS:
            dst.writestr(chart_partname, chart_xml.encode("utf-8"))
            dst.writestr(
                "ppt/charts/_rels/%s.rels" % chart_partname.rsplit("/", 1)[1],
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                'relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships/package" Target="../embeddings/%s"/>'
                "</Relationships>" % xlsx_partname.rsplit("/", 1)[1],
            )
            dst.writestr(xlsx_partname, xlsx_blob)
    out.seek(0)
    return out


def xlsx_column(xlsx_blob: bytes, col: str, first_row: int = 2) -> list[object]:
    """Cell values in column `col` of the first worksheet of `xlsx_blob`, from `first_row` down.

    Numbers come back as float and shared strings as str; a blank cell is |None|. Enough of the
    SpreadsheetML format to check a chart's cache against its workbook, without a spreadsheet
    library.
    """
    with zipfile.ZipFile(io.BytesIO(xlsx_blob)) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        names = z.namelist()
        shared = (
            re.findall(r"<t[^>]*>([^<]*)</t>", z.read("xl/sharedStrings.xml").decode("utf-8"))
            if "xl/sharedStrings.xml" in names
            else []
        )
    cells: dict[int, object] = {}
    for ref, attrs, value in re.findall(
        r'<c r="%s(\d+)"([^>]*)>(?:<f>[^<]*</f>)?<v>([^<]*)</v></c>' % col, sheet
    ):
        cells[int(ref)] = shared[int(value)] if 't="s"' in attrs else float(value)
    last_row = max(cells) if cells else first_row - 1
    return [cells.get(row) for row in range(first_row, last_row + 1)]
