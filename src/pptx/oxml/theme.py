"""lxml custom element classes for theme-related XML elements."""

from __future__ import annotations

from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne

from . import parse_from_template


class CT_Color(BaseOxmlElement):
    """Color element within a color scheme (e.g. `a:dk1`, `a:lt1`, `a:accent1`)."""


class CT_ColorScheme(BaseOxmlElement):
    """`a:clrScheme` element, defining the color scheme of a theme."""

    _tag_seq = (
        "a:dk1",
        "a:lt1",
        "a:dk2",
        "a:lt2",
        "a:accent1",
        "a:accent2",
        "a:accent3",
        "a:accent4",
        "a:accent5",
        "a:accent6",
        "a:hlink",
        "a:folHlink",
        "a:extLst",
    )
    dk1: CT_Color | None = ZeroOrOne("a:dk1", successors=_tag_seq[1:])  # pyright: ignore[reportAssignmentType]
    lt1: CT_Color | None = ZeroOrOne("a:lt1", successors=_tag_seq[2:])  # pyright: ignore[reportAssignmentType]
    dk2: CT_Color | None = ZeroOrOne("a:dk2", successors=_tag_seq[3:])  # pyright: ignore[reportAssignmentType]
    lt2: CT_Color | None = ZeroOrOne("a:lt2", successors=_tag_seq[4:])  # pyright: ignore[reportAssignmentType]
    accent1: CT_Color | None = ZeroOrOne("a:accent1", successors=_tag_seq[5:])  # pyright: ignore[reportAssignmentType]
    accent2: CT_Color | None = ZeroOrOne("a:accent2", successors=_tag_seq[6:])  # pyright: ignore[reportAssignmentType]
    accent3: CT_Color | None = ZeroOrOne("a:accent3", successors=_tag_seq[7:])  # pyright: ignore[reportAssignmentType]
    accent4: CT_Color | None = ZeroOrOne("a:accent4", successors=_tag_seq[8:])  # pyright: ignore[reportAssignmentType]
    accent5: CT_Color | None = ZeroOrOne("a:accent5", successors=_tag_seq[9:])  # pyright: ignore[reportAssignmentType]
    accent6: CT_Color | None = ZeroOrOne("a:accent6", successors=_tag_seq[10:])  # pyright: ignore[reportAssignmentType]
    hlink: CT_Color | None = ZeroOrOne("a:hlink", successors=_tag_seq[11:])  # pyright: ignore[reportAssignmentType]
    folHlink: CT_Color | None = ZeroOrOne("a:folHlink", successors=_tag_seq[12:])  # pyright: ignore[reportAssignmentType]
    del _tag_seq

    name: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString
    )


class CT_FontScheme(BaseOxmlElement):
    """`a:fontScheme` element, defining the font scheme of a theme."""

    _tag_seq = ("a:majorFont", "a:minorFont", "a:extLst")
    majorFont = ZeroOrOne("a:majorFont", successors=_tag_seq[1:])
    minorFont = ZeroOrOne("a:minorFont", successors=_tag_seq[2:])
    del _tag_seq

    name: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString
    )


class CT_FontCollection(BaseOxmlElement):
    """`a:majorFont` or `a:minorFont` element, specifying a font collection."""

    _tag_seq = ("a:latin", "a:ea", "a:cs", "a:font", "a:extLst")
    latin = ZeroOrOne("a:latin", successors=_tag_seq[1:])
    ea = ZeroOrOne("a:ea", successors=_tag_seq[2:])
    cs = ZeroOrOne("a:cs", successors=_tag_seq[3:])
    del _tag_seq


class CT_BaseStyles(BaseOxmlElement):
    """`a:themeElements` element containing the core theme definitions."""

    _tag_seq = ("a:clrScheme", "a:fontScheme", "a:fmtScheme", "a:extLst")
    clrScheme: CT_ColorScheme | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:clrScheme", successors=_tag_seq[1:]
    )
    fontScheme: CT_FontScheme | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:fontScheme", successors=_tag_seq[2:]
    )
    del _tag_seq


class CT_OfficeStyleSheet(BaseOxmlElement):
    """``<a:theme>`` element, root of a theme part."""

    _tag_seq = (
        "a:themeElements",
        "a:objectDefaults",
        "a:extraClrSchemeLst",
        "a:custClrLst",
        "a:extLst",
    )
    themeElements: CT_BaseStyles | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:themeElements", successors=_tag_seq[1:]
    )
    del _tag_seq

    @classmethod
    def new_default(cls):
        """Return a new ``<a:theme>`` element containing default settings
        suitable for use with a notes master.
        """
        return parse_from_template("theme")
