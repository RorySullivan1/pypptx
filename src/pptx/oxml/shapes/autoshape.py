# pyright: reportPrivateUsage=false

"""lxml custom element classes for shape-related XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.shapes.shared import BaseShapeElement
from pptx.oxml.simpletypes import (
    ST_Angle,
    ST_Coordinate,
    ST_PositiveCoordinate,
    XsdBoolean,
    XsdString,
)
from pptx.oxml.text import CT_TextBody
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from pptx.oxml.shapes.shared import (
        CT_ApplicationNonVisualDrawingProps,
        CT_NonVisualDrawingProps,
        CT_ShapeProperties,
    )
    from pptx.util import Length


class CT_AdjPoint2D(BaseOxmlElement):
    """`a:pt` custom element class."""

    x: Length = RequiredAttribute("x", ST_Coordinate)  # pyright: ignore[reportAssignmentType]
    y: Length = RequiredAttribute("y", ST_Coordinate)  # pyright: ignore[reportAssignmentType]


class CT_CustomGeometry2D(BaseOxmlElement):
    """`a:custGeom` custom element class."""

    get_or_add_pathLst: Callable[[], CT_Path2DList]

    _tag_seq = ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst", "a:rect", "a:pathLst")
    pathLst: CT_Path2DList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:pathLst", successors=_tag_seq[6:]
    )


class CT_GeomGuide(BaseOxmlElement):
    """`a:gd` custom element class.

    Defines a "guide", corresponding to a yellow diamond-shaped handle on an autoshape.
    """

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    fmla: str = RequiredAttribute("fmla", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_GeomGuideList(BaseOxmlElement):
    """`a:avLst` custom element class."""

    _add_gd: Callable[[], CT_GeomGuide]

    gd_lst: list[CT_GeomGuide]

    gd = ZeroOrMore("a:gd")


class CT_NonVisualDrawingShapeProps(BaseShapeElement):
    """`p:cNvSpPr` custom element class."""

    spLocks = ZeroOrOne("a:spLocks")
    txBox: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "txBox", XsdBoolean
    )


class CT_Path2D(BaseOxmlElement):
    """`a:path` custom element class."""

    _add_arcTo: Callable[[], CT_Path2DArcTo]
    _add_close: Callable[[], CT_Path2DClose]
    _add_cubicBezTo: Callable[[], CT_Path2DCubicBezierTo]
    _add_lnTo: Callable[[], CT_Path2DLineTo]
    _add_moveTo: Callable[[], CT_Path2DMoveTo]
    _add_quadBezTo: Callable[[], CT_Path2DQuadBezierTo]

    arcTo = ZeroOrMore("a:arcTo", successors=())
    close = ZeroOrMore("a:close", successors=())
    cubicBezTo = ZeroOrMore("a:cubicBezTo", successors=())
    lnTo = ZeroOrMore("a:lnTo", successors=())
    moveTo = ZeroOrMore("a:moveTo", successors=())
    quadBezTo = ZeroOrMore("a:quadBezTo", successors=())
    w: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w", ST_PositiveCoordinate
    )
    h: Length | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "h", ST_PositiveCoordinate
    )

    def add_arcTo(self, wR: Length, hR: Length, stAng: float, swAng: float) -> CT_Path2DArcTo:
        """Return a newly created `a:arcTo` element for an elliptical arc segment.

        `wR` and `hR` are the horizontal and vertical radii of the arc's ellipse, in the
        path's local coordinates. `stAng` and `swAng` are the start angle and swing (sweep)
        angle of the arc, in degrees; they are written to XML in 60000ths of a degree per the
        `ST_Angle` schema type. The new `a:arcTo` element is appended to this `a:path`
        element.
        """
        arcTo = self._add_arcTo()
        arcTo.wR, arcTo.hR, arcTo.stAng, arcTo.swAng = wR, hR, stAng, swAng
        return arcTo

    def add_close(self) -> CT_Path2DClose:
        """Return a newly created `a:close` element.

        The new `a:close` element is appended to this `a:path` element.
        """
        return self._add_close()

    def add_cubicBezTo(
        self, x1: Length, y1: Length, x2: Length, y2: Length, x3: Length, y3: Length
    ) -> CT_Path2DCubicBezierTo:
        """Return a newly created `a:cubicBezTo` subtree for a cubic Bezier curve segment.

        `(x1, y1)` and `(x2, y2)` are the curve's first and second control points and
        `(x3, y3)` is its end point. The new `a:cubicBezTo` element is appended to this
        `a:path` element.
        """
        cubicBezTo = self._add_cubicBezTo()
        for x, y in ((x1, y1), (x2, y2), (x3, y3)):
            pt = cubicBezTo._add_pt()
            pt.x, pt.y = x, y
        return cubicBezTo

    def add_lnTo(self, x: Length, y: Length) -> CT_Path2DLineTo:
        """Return a newly created `a:lnTo` subtree with end point *(x, y)*.

        The new `a:lnTo` element is appended to this `a:path` element.
        """
        lnTo = self._add_lnTo()
        pt = lnTo._add_pt()
        pt.x, pt.y = x, y
        return lnTo

    def add_moveTo(self, x: Length, y: Length):
        """Return a newly created `a:moveTo` subtree with point `(x, y)`.

        The new `a:moveTo` element is appended to this `a:path` element.
        """
        moveTo = self._add_moveTo()
        pt = moveTo._add_pt()
        pt.x, pt.y = x, y
        return moveTo

    def add_quadBezTo(self, x1: Length, y1: Length, x2: Length, y2: Length) -> CT_Path2DQuadBezierTo:
        """Return a newly created `a:quadBezTo` subtree for a quadratic Bezier curve segment.

        `(x1, y1)` is the curve's control point and `(x2, y2)` is its end point. The new
        `a:quadBezTo` element is appended to this `a:path` element.
        """
        quadBezTo = self._add_quadBezTo()
        for x, y in ((x1, y1), (x2, y2)):
            pt = quadBezTo._add_pt()
            pt.x, pt.y = x, y
        return quadBezTo


class CT_Path2DArcTo(BaseOxmlElement):
    """`a:arcTo` custom element class."""

    wR: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "wR", ST_PositiveCoordinate
    )
    hR: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "hR", ST_PositiveCoordinate
    )
    stAng: float = RequiredAttribute("stAng", ST_Angle)  # pyright: ignore[reportAssignmentType]
    swAng: float = RequiredAttribute("swAng", ST_Angle)  # pyright: ignore[reportAssignmentType]


class CT_Path2DClose(BaseOxmlElement):
    """`a:close` custom element class."""


class CT_Path2DCubicBezierTo(BaseOxmlElement):
    """`a:cubicBezTo` custom element class."""

    _add_pt: Callable[[], CT_AdjPoint2D]

    pt = ZeroOrMore("a:pt", successors=())


class CT_Path2DLineTo(BaseOxmlElement):
    """`a:lnTo` custom element class."""

    _add_pt: Callable[[], CT_AdjPoint2D]

    pt = ZeroOrOne("a:pt", successors=())


class CT_Path2DList(BaseOxmlElement):
    """`a:pathLst` custom element class."""

    _add_path: Callable[[], CT_Path2D]

    path = ZeroOrMore("a:path", successors=())

    def add_path(self, w: Length, h: Length):
        """Return a newly created `a:path` child element."""
        path = self._add_path()
        path.w, path.h = w, h
        return path


class CT_Path2DMoveTo(BaseOxmlElement):
    """`a:moveTo` custom element class."""

    _add_pt: Callable[[], CT_AdjPoint2D]

    pt = ZeroOrOne("a:pt", successors=())


class CT_Path2DQuadBezierTo(BaseOxmlElement):
    """`a:quadBezTo` custom element class."""

    _add_pt: Callable[[], CT_AdjPoint2D]

    pt = ZeroOrMore("a:pt", successors=())


class CT_PresetGeometry2D(BaseOxmlElement):
    """`a:prstGeom` custom element class."""

    _add_avLst: Callable[[], CT_GeomGuideList]
    _remove_avLst: Callable[[], None]

    avLst: CT_GeomGuideList | None = ZeroOrOne("a:avLst")  # pyright: ignore[reportAssignmentType]
    prst: MSO_AUTO_SHAPE_TYPE = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "prst", MSO_AUTO_SHAPE_TYPE
    )

    @property
    def gd_lst(self) -> list[CT_GeomGuide]:
        """Sequence of `a:gd` element children of `a:avLst`. Empty if none are present."""
        avLst = self.avLst
        if avLst is None:
            return []
        return avLst.gd_lst

    def rewrite_guides(self, guides: list[tuple[str, int]]):
        """Replace any `a:gd` element children of `a:avLst` with ones forme from `guides`."""
        self._remove_avLst()
        avLst = self._add_avLst()
        for name, val in guides:
            gd = avLst._add_gd()
            gd.name = name
            gd.fmla = "val %d" % val


class CT_Shape(BaseShapeElement):
    """`p:sp` custom element class."""

    get_or_add_txBody: Callable[[], CT_TextBody]

    nvSpPr: CT_ShapeNonVisual = OneAndOnlyOne("p:nvSpPr")  # pyright: ignore[reportAssignmentType]
    spPr: CT_ShapeProperties = OneAndOnlyOne("p:spPr")  # pyright: ignore[reportAssignmentType]
    style = ZeroOrOne("p:style", successors=("p:txBody", "p:extLst"))
    txBody: CT_TextBody | None = ZeroOrOne("p:txBody", successors=("p:extLst",))  # pyright: ignore

    def add_path(self, w: Length, h: Length) -> CT_Path2D:
        custGeom = self.spPr.custGeom
        if custGeom is None:
            raise ValueError("shape must be freeform")
        pathLst = custGeom.get_or_add_pathLst()
        return pathLst.add_path(w=w, h=h)

    def get_or_add_ln(self):
        """Return the `a:ln` grandchild element, newly added if not present."""
        return self.spPr.get_or_add_ln()

    @property
    def has_custom_geometry(self):
        """True if this shape has custom geometry, i.e. is a freeform shape.

        A shape has custom geometry if it has a `p:spPr/a:custGeom`
        descendant (instead of `p:spPr/a:prstGeom`).
        """
        return self.spPr.custGeom is not None

    @property
    def is_autoshape(self):
        """True if this shape is an auto shape.

        A shape is an auto shape if it has a `a:prstGeom` element and does not have a txBox="1"
        attribute on cNvSpPr.
        """
        prstGeom = self.prstGeom
        if prstGeom is None:
            return False
        return self.nvSpPr.cNvSpPr.txBox is not True

    @property
    def is_textbox(self):
        """True if this shape is a text box.

        A shape is a text box if it has a `txBox` attribute on cNvSpPr that resolves to |True|.
        The default when the txBox attribute is missing is |False|.
        """
        return self.nvSpPr.cNvSpPr.txBox is True

    @property
    def ln(self):
        """`a:ln` grand-child element or |None| if not present."""
        return self.spPr.ln

    @staticmethod
    def new_autoshape_sp(
        id_: int, name: str, prst: str, left: int, top: int, width: int, height: int
    ) -> CT_Shape:
        """Return a new `p:sp` element tree configured as a base auto shape."""
        xml = (
            "<p:sp %s>\n"
            "  <p:nvSpPr>\n"
            '    <p:cNvPr id="%s" name="%s"/>\n'
            "    <p:cNvSpPr/>\n"
            "    <p:nvPr/>\n"
            "  </p:nvSpPr>\n"
            "  <p:spPr>\n"
            "    <a:xfrm>\n"
            '      <a:off x="%s" y="%s"/>\n'
            '      <a:ext cx="%s" cy="%s"/>\n'
            "    </a:xfrm>\n"
            '    <a:prstGeom prst="%s">\n'
            "      <a:avLst/>\n"
            "    </a:prstGeom>\n"
            "  </p:spPr>\n"
            "  <p:style>\n"
            '    <a:lnRef idx="1">\n'
            '      <a:schemeClr val="accent1"/>\n'
            "    </a:lnRef>\n"
            '    <a:fillRef idx="3">\n'
            '      <a:schemeClr val="accent1"/>\n'
            "    </a:fillRef>\n"
            '    <a:effectRef idx="2">\n'
            '      <a:schemeClr val="accent1"/>\n'
            "    </a:effectRef>\n"
            '    <a:fontRef idx="minor">\n'
            '      <a:schemeClr val="lt1"/>\n'
            "    </a:fontRef>\n"
            "  </p:style>\n"
            "  <p:txBody>\n"
            '    <a:bodyPr rtlCol="0" anchor="ctr"/>\n'
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            '      <a:pPr algn="ctr"/>\n'
            "    </a:p>\n"
            "  </p:txBody>\n"
            "</p:sp>" % (nsdecls("a", "p"), "%d", "%s", "%d", "%d", "%d", "%d", "%s")
        ) % (id_, name, left, top, width, height, prst)
        return cast(CT_Shape, parse_xml(xml))

    @staticmethod
    def new_freeform_sp(shape_id: int, name: str, x: int, y: int, cx: int, cy: int):
        """Return new `p:sp` element tree configured as freeform shape.

        The returned shape has a `a:custGeom` subtree but no paths in its
        path list.
        """
        xml = (
            "<p:sp %s>\n"
            "  <p:nvSpPr>\n"
            '    <p:cNvPr id="%s" name="%s"/>\n'
            "    <p:cNvSpPr/>\n"
            "    <p:nvPr/>\n"
            "  </p:nvSpPr>\n"
            "  <p:spPr>\n"
            "    <a:xfrm>\n"
            '      <a:off x="%s" y="%s"/>\n'
            '      <a:ext cx="%s" cy="%s"/>\n'
            "    </a:xfrm>\n"
            "    <a:custGeom>\n"
            "      <a:avLst/>\n"
            "      <a:gdLst/>\n"
            "      <a:ahLst/>\n"
            "      <a:cxnLst/>\n"
            '      <a:rect l="l" t="t" r="r" b="b"/>\n'
            "      <a:pathLst/>\n"
            "    </a:custGeom>\n"
            "  </p:spPr>\n"
            "  <p:style>\n"
            '    <a:lnRef idx="1">\n'
            '      <a:schemeClr val="accent1"/>\n'
            "    </a:lnRef>\n"
            '    <a:fillRef idx="3">\n'
            '      <a:schemeClr val="accent1"/>\n'
            "    </a:fillRef>\n"
            '    <a:effectRef idx="2">\n'
            '      <a:schemeClr val="accent1"/>\n'
            "    </a:effectRef>\n"
            '    <a:fontRef idx="minor">\n'
            '      <a:schemeClr val="lt1"/>\n'
            "    </a:fontRef>\n"
            "  </p:style>\n"
            "  <p:txBody>\n"
            '    <a:bodyPr rtlCol="0" anchor="ctr"/>\n'
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            '      <a:pPr algn="ctr"/>\n'
            "    </a:p>\n"
            "  </p:txBody>\n"
            "</p:sp>" % (nsdecls("a", "p"), "%d", "%s", "%d", "%d", "%d", "%d")
        ) % (shape_id, name, x, y, cx, cy)
        return cast(CT_Shape, parse_xml(xml))

    @staticmethod
    def new_placeholder_sp(
        id_: int, name: str, ph_type: PP_PLACEHOLDER, orient: str, sz, idx
    ) -> CT_Shape:
        """Return a new `p:sp` element tree configured as a placeholder shape."""
        sp = cast(
            CT_Shape,
            parse_xml(
                f"<p:sp {nsdecls('a', 'p')}>\n"
                f"  <p:nvSpPr>\n"
                f'    <p:cNvPr id="{id_}" name="{name}"/>\n'
                f"    <p:cNvSpPr>\n"
                f'      <a:spLocks noGrp="1"/>\n'
                f"    </p:cNvSpPr>\n"
                f"    <p:nvPr/>\n"
                f"  </p:nvSpPr>\n"
                f"  <p:spPr/>\n"
                f"</p:sp>"
            ),
        )

        ph = sp.nvSpPr.nvPr.get_or_add_ph()
        ph.type = ph_type
        ph.idx = idx
        ph.orient = orient
        ph.sz = sz

        placeholder_types_that_have_a_text_frame = (
            PP_PLACEHOLDER.TITLE,
            PP_PLACEHOLDER.CENTER_TITLE,
            PP_PLACEHOLDER.SUBTITLE,
            PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.OBJECT,
        )

        if ph_type in placeholder_types_that_have_a_text_frame:
            sp.append(CT_TextBody.new())

        return sp

    @staticmethod
    def new_textbox_sp(id_, name, left, top, width, height):
        """Return a new `p:sp` element tree configured as a base textbox shape."""
        tmpl = CT_Shape._textbox_sp_tmpl()
        xml = tmpl % (id_, name, left, top, width, height)
        sp = parse_xml(xml)
        return sp

    @property
    def prst(self):
        """Value of `prst` attribute of `a:prstGeom` element or |None| if not present."""
        prstGeom = self.prstGeom
        if prstGeom is None:
            return None
        return prstGeom.prst

    @property
    def prstGeom(self) -> CT_PresetGeometry2D:
        """Reference to `a:prstGeom` child element.

        |None| if this shape doesn't have one, for example, if it's a placeholder shape.
        """
        return self.spPr.prstGeom

    def _new_txBody(self):
        return CT_TextBody.new_p_txBody()

    @staticmethod
    def _textbox_sp_tmpl():
        return (
            "<p:sp %s>\n"
            "  <p:nvSpPr>\n"
            '    <p:cNvPr id="%s" name="%s"/>\n'
            '    <p:cNvSpPr txBox="1"/>\n'
            "    <p:nvPr/>\n"
            "  </p:nvSpPr>\n"
            "  <p:spPr>\n"
            "    <a:xfrm>\n"
            '      <a:off x="%s" y="%s"/>\n'
            '      <a:ext cx="%s" cy="%s"/>\n'
            "    </a:xfrm>\n"
            '    <a:prstGeom prst="rect">\n'
            "      <a:avLst/>\n"
            "    </a:prstGeom>\n"
            "    <a:noFill/>\n"
            "  </p:spPr>\n"
            "  <p:txBody>\n"
            '    <a:bodyPr wrap="none">\n'
            "      <a:spAutoFit/>\n"
            "    </a:bodyPr>\n"
            "    <a:lstStyle/>\n"
            "    <a:p/>\n"
            "  </p:txBody>\n"
            "</p:sp>" % (nsdecls("a", "p"), "%d", "%s", "%d", "%d", "%d", "%d")
        )


class CT_ShapeNonVisual(BaseShapeElement):
    """`p:nvSpPr` custom element class."""

    cNvPr: CT_NonVisualDrawingProps = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cNvPr"
    )
    cNvSpPr: CT_NonVisualDrawingShapeProps = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p:cNvSpPr"
    )
    nvPr: CT_ApplicationNonVisualDrawingProps = (  # pyright: ignore[reportAssignmentType]
        OneAndOnlyOne("p:nvPr")
    )
