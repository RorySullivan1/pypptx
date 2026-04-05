"""lxml custom element classes for DrawingML-related XML elements."""

from __future__ import annotations

from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.simpletypes import (
    ST_HexColorRGB,
    ST_Percentage,
    ST_PositiveFixedAngle,
    XsdString,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
)


class _BaseColorElement(BaseOxmlElement):
    """
    Base class for <a:srgbClr> and <a:schemeClr> elements.
    """

    tint = ZeroOrOne("a:tint")
    shade = ZeroOrOne("a:shade")
    satMod = ZeroOrOne("a:satMod")
    satOff = ZeroOrOne("a:satOff")
    lumMod = ZeroOrOne("a:lumMod")
    lumOff = ZeroOrOne("a:lumOff")
    alpha = ZeroOrOne("a:alpha")

    def add_lumMod(self, value):
        """
        Return a newly added <a:lumMod> child element.
        """
        lumMod = self._add_lumMod()
        lumMod.val = value
        return lumMod

    def add_lumOff(self, value):
        """
        Return a newly added <a:lumOff> child element.
        """
        lumOff = self._add_lumOff()
        lumOff.val = value
        return lumOff

    def clear_lum(self):
        """
        Return self after removing any <a:lumMod> and <a:lumOff> child
        elements.
        """
        self._remove_lumMod()
        self._remove_lumOff()
        return self


class CT_Color(BaseOxmlElement):
    """Custom element class for `a:fgClr`, `a:bgClr` and perhaps others."""

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


class CT_HslColor(_BaseColorElement):
    """Custom element class for `a:hslClr` element."""

    hue: float = RequiredAttribute("hue", ST_PositiveFixedAngle)  # pyright: ignore[reportAssignmentType]
    sat: float = RequiredAttribute("sat", ST_Percentage)  # pyright: ignore[reportAssignmentType]
    lum: float = RequiredAttribute("lum", ST_Percentage)  # pyright: ignore[reportAssignmentType]


class CT_Percentage(BaseOxmlElement):
    """
    Custom element class for <a:lumMod> and <a:lumOff> elements.
    """

    val = RequiredAttribute("val", ST_Percentage)


class CT_PresetColor(_BaseColorElement):
    """Custom element class for `a:prstClr` element."""

    val: str = RequiredAttribute("val", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SchemeColor(_BaseColorElement):
    """
    Custom element class for <a:schemeClr> element.
    """

    val = RequiredAttribute("val", MSO_THEME_COLOR)


class CT_ScRgbColor(_BaseColorElement):
    """Custom element class for `a:scrgbClr` element."""

    r: float = RequiredAttribute("r", ST_Percentage)  # pyright: ignore[reportAssignmentType]
    g: float = RequiredAttribute("g", ST_Percentage)  # pyright: ignore[reportAssignmentType]
    b: float = RequiredAttribute("b", ST_Percentage)  # pyright: ignore[reportAssignmentType]


class CT_SRgbColor(_BaseColorElement):
    """
    Custom element class for <a:srgbClr> element.
    """

    val = RequiredAttribute("val", ST_HexColorRGB)


class CT_SystemColor(_BaseColorElement):
    """Custom element class for `a:sysClr` element."""

    val: str = RequiredAttribute("val", XsdString)  # pyright: ignore[reportAssignmentType]
    lastClr: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "lastClr", ST_HexColorRGB
    )
