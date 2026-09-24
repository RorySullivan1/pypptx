"""lxml custom element classes for DrawingML-related XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

from pptx.enum.dml import MSO_PATTERN_TYPE, MSO_RECT_ALIGNMENT
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import (
    ST_Coordinate,
    ST_Percentage,
    ST_PositiveFixedAngle,
    ST_PositiveFixedPercentage,
    ST_RelationshipId,
    ST_TileFlipMode,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OneOrMore,
    OptionalAttribute,
    OxmlElement,
    RequiredAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
)
from pptx.util import Length

if TYPE_CHECKING:
    from pptx.oxml.shapes.autoshape import CT_Path2D


class CT_Blip(BaseOxmlElement):
    """`a:blip` element, specifying an image resource and optional effects."""

    get_or_add_clrChange: Callable[[], CT_ColorChangeEffect]
    get_or_add_grayscl: Callable[[], BaseOxmlElement]
    get_or_add_lum: Callable[[], BaseOxmlElement]
    _remove_clrChange: Callable[[], None]
    _remove_grayscl: Callable[[], None]
    _remove_lum: Callable[[], None]

    _tag_seq = (
        "a:alphaBiLevel",
        "a:alphaCeiling",
        "a:alphaFloor",
        "a:alphaInv",
        "a:alphaMod",
        "a:alphaModFix",
        "a:alphaRepl",
        "a:biLevel",
        "a:blur",
        "a:clrChange",
        "a:clrRepl",
        "a:duotone",
        "a:fillOverlay",
        "a:grayscl",
        "a:hsl",
        "a:lum",
        "a:tint",
        "a:extLst",
    )
    clrChange: CT_ColorChangeEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:clrChange", successors=_tag_seq[10:]
    )
    duotone = ZeroOrOne("a:duotone", successors=_tag_seq[12:])
    grayscl = ZeroOrOne("a:grayscl", successors=_tag_seq[14:])
    lum = ZeroOrOne("a:lum", successors=_tag_seq[16:])
    del _tag_seq

    rEmbed = OptionalAttribute("r:embed", ST_RelationshipId)
    rLink = OptionalAttribute("r:link", ST_RelationshipId)


class CT_ColorChangeEffect(BaseOxmlElement):
    """`a:clrChange` element, specifying a color-to-transparent replacement on a blip.

    Contains `a:clrFrom` (color to replace) and `a:clrTo` (replacement, typically transparent).
    """

    _tag_seq = ("a:clrFrom", "a:clrTo")
    clrFrom = ZeroOrOne("a:clrFrom", successors=_tag_seq[1:])
    clrTo = ZeroOrOne("a:clrTo", successors=())
    del _tag_seq


class CT_BlipFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:blipFill> element.
    """

    _tag_seq = ("a:blip", "a:srcRect", "a:tile", "a:stretch")
    blip = ZeroOrOne("a:blip", successors=_tag_seq[1:])
    srcRect = ZeroOrOne("a:srcRect", successors=_tag_seq[2:])
    # -- `a:tile` and `a:stretch` are mutually exclusive per the schema --
    eg_tileStretch = ZeroOrOneChoice(
        (Choice("a:tile"), Choice("a:stretch")),
        successors=(),
    )
    del _tag_seq

    def crop(self, cropping):
        """
        Set `a:srcRect` child to crop according to *cropping* values.
        """
        srcRect = self._add_srcRect()
        srcRect.l, srcRect.t, srcRect.r, srcRect.b = cropping


class CT_StretchInfoProperties(BaseOxmlElement):
    """`a:stretch` custom element class.

    Specifies a picture fill (or `p:pic`) is scaled to fill its bounding box, optionally
    cropped to a sub-rectangle via `a:fillRect`.
    """

    fillRect = ZeroOrOne("a:fillRect", successors=())


class CT_TileInfoProperties(BaseOxmlElement):
    """`a:tile` custom element class.

    Specifies a picture fill is tiled (repeated) to fill its bounding box, with an optional
    offset, scale, flip, and alignment for the first tile.
    """

    tx: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "tx", ST_Coordinate, default=Length(0)
    )
    ty: Length = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "ty", ST_Coordinate, default=Length(0)
    )
    sx: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "sx", ST_Percentage, default=1.0
    )
    sy: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "sy", ST_Percentage, default=1.0
    )
    flip: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "flip", ST_TileFlipMode, default=ST_TileFlipMode.NONE
    )
    algn: MSO_RECT_ALIGNMENT = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "algn", MSO_RECT_ALIGNMENT, default=MSO_RECT_ALIGNMENT.TOP_LEFT
    )


class CT_GradientFillProperties(BaseOxmlElement):
    """`a:gradFill` custom element class."""

    get_or_add_path: Callable[[], CT_Path2D]

    _tag_seq = ("a:gsLst", "a:lin", "a:path", "a:tileRect")
    gsLst = ZeroOrOne("a:gsLst", successors=_tag_seq[1:])
    lin = ZeroOrOne("a:lin", successors=_tag_seq[2:])
    # -- NOTE: the `a:path` child here shares its tag with the unrelated `a:path` element used
    # -- in `a:custGeom/a:pathLst` (CT_Path2D), which is the class registered for this tag. Only
    # -- generic (untyped) attribute/child access is used on it below, via `path_val` and
    # -- `fillToRect`/`get_or_add_fillToRect`, rather than any CT_Path2D-specific API.
    path: CT_Path2D | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:path", successors=_tag_seq[3:]
    )
    del _tag_seq

    @property
    def path_val(self) -> str | None:
        """Value of the `path` attribute on the `a:path` child, or |None| if not present."""
        path = self.path
        if path is None:
            return None
        return path.get("path")

    @path_val.setter
    def path_val(self, value: str) -> None:
        path = self.get_or_add_path()
        path.set("path", value)

    @property
    def fillToRect(self) -> CT_RelativeRect | None:
        """The `a:fillToRect` grandchild (child of `a:path`), or |None| if not present."""
        path = self.path
        if path is None:
            return None
        return path.find(qn("a:fillToRect"))

    def get_or_add_fillToRect(self) -> CT_RelativeRect:
        """Return the `a:fillToRect` grandchild, adding `a:path` and/or itself if needed."""
        path = self.get_or_add_path()
        fillToRect = path.find(qn("a:fillToRect"))
        if fillToRect is None:
            fillToRect = OxmlElement("a:fillToRect")
            path.insert(0, fillToRect)
        return fillToRect

    def _remove_fillToRect(self) -> None:
        """Remove the `a:fillToRect` grandchild, if present. No-op if `a:path` is absent."""
        path = self.path
        if path is None:
            return
        fillToRect = path.find(qn("a:fillToRect"))
        if fillToRect is not None:
            path.remove(fillToRect)

    @classmethod
    def new_gradFill(cls):
        """Return newly-created "loose" default gradient subtree."""
        return parse_xml(
            '<a:gradFill %s rotWithShape="1">\n'
            "  <a:gsLst>\n"
            '    <a:gs pos="0">\n'
            '      <a:schemeClr val="accent1">\n'
            '        <a:tint val="100000"/>\n'
            '        <a:shade val="100000"/>\n'
            '        <a:satMod val="130000"/>\n'
            "      </a:schemeClr>\n"
            "    </a:gs>\n"
            '    <a:gs pos="100000">\n'
            '      <a:schemeClr val="accent1">\n'
            '        <a:tint val="50000"/>\n'
            '        <a:shade val="100000"/>\n'
            '        <a:satMod val="350000"/>\n'
            "      </a:schemeClr>\n"
            "    </a:gs>\n"
            "  </a:gsLst>\n"
            '  <a:lin scaled="0"/>\n'
            "</a:gradFill>\n" % nsdecls("a")
        )

    def _new_gsLst(self):
        """Override default to add minimum subtree."""
        return CT_GradientStopList.new_gsLst()


class CT_GradientStop(BaseOxmlElement):
    """`a:gs` custom element class."""

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    pos = RequiredAttribute("pos", ST_PositiveFixedPercentage)


class CT_GradientStopList(BaseOxmlElement):
    """`a:gsLst` custom element class."""

    gs = OneOrMore("a:gs")

    @classmethod
    def new_gsLst(cls):
        """Return newly-created "loose" default stop-list subtree.

        An `a:gsLst` element must have at least two `a:gs` children. These
        are the default from the PowerPoint built-in "White" template.
        """
        return parse_xml(
            "<a:gsLst %s>\n"
            '  <a:gs pos="0">\n'
            '    <a:schemeClr val="accent1">\n'
            '      <a:tint val="100000"/>\n'
            '      <a:shade val="100000"/>\n'
            '      <a:satMod val="130000"/>\n'
            "    </a:schemeClr>\n"
            "  </a:gs>\n"
            '  <a:gs pos="100000">\n'
            '    <a:schemeClr val="accent1">\n'
            '      <a:tint val="50000"/>\n'
            '      <a:shade val="100000"/>\n'
            '      <a:satMod val="350000"/>\n'
            "    </a:schemeClr>\n"
            "  </a:gs>\n"
            "</a:gsLst>\n" % nsdecls("a")
        )


class CT_GroupFillProperties(BaseOxmlElement):
    """`a:grpFill` custom element class"""


class CT_LinearShadeProperties(BaseOxmlElement):
    """`a:lin` custom element class"""

    ang = OptionalAttribute("ang", ST_PositiveFixedAngle)


class CT_NoFillProperties(BaseOxmlElement):
    """`a:noFill` custom element class"""


class CT_PatternFillProperties(BaseOxmlElement):
    """`a:pattFill` custom element class"""

    _tag_seq = ("a:fgClr", "a:bgClr")
    fgClr = ZeroOrOne("a:fgClr", successors=_tag_seq[1:])
    bgClr = ZeroOrOne("a:bgClr", successors=_tag_seq[2:])
    del _tag_seq
    prst = OptionalAttribute("prst", MSO_PATTERN_TYPE)

    def _new_bgClr(self):
        """Override default to add minimum subtree."""
        xml = ("<a:bgClr %s>\n" ' <a:srgbClr val="FFFFFF"/>\n' "</a:bgClr>\n") % nsdecls("a")
        bgClr = parse_xml(xml)
        return bgClr

    def _new_fgClr(self):
        """Override default to add minimum subtree."""
        xml = ("<a:fgClr %s>\n" ' <a:srgbClr val="000000"/>\n' "</a:fgClr>\n") % nsdecls("a")
        fgClr = parse_xml(xml)
        return fgClr


class CT_RelativeRect(BaseOxmlElement):
    """`a:srcRect` element and perhaps others."""

    l: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]  # noqa: E741
        "l", ST_Percentage, default=0.0
    )
    t: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "t", ST_Percentage, default=0.0
    )
    r: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r", ST_Percentage, default=0.0
    )
    b: float = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "b", ST_Percentage, default=0.0
    )


class CT_SolidColorFillProperties(BaseOxmlElement):
    """`a:solidFill` custom element class."""

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
