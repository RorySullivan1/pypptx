"""Custom element classes for the presentation properties (presProps.xml) part.

Covers `p:presentationPr` (the root element) and `p:showPr` (slide-show settings),
per ECMA-376 §19.2.1.29 (CT_PresentationProperties) and §19.2.1.31
(CT_ShowProperties). View properties, MRU colours, and web publishing properties
are out of scope.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

from pptx.oxml.simpletypes import XsdBoolean, XsdInt, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
)

if TYPE_CHECKING:
    from pptx.oxml.dml.color import CT_Color


class CT_PresentationProperties(BaseOxmlElement):
    """`p:presentationPr` element, root of the PresPropsPart, stored as ``ppt/presProps.xml``."""

    get_or_add_showPr: Callable[[], CT_ShowProperties]

    _tag_seq = ("p:showPr", "p:clrMru", "p:extLst")
    showPr: CT_ShowProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:showPr", successors=_tag_seq[1:]
    )
    del _tag_seq


class CT_ShowProperties(BaseOxmlElement):
    """`p:showPr` element, specifying slide-show settings for the presentation."""

    get_or_add_penClr: Callable[[], CT_Color]

    _tag_seq = (
        "p:present",
        "p:browse",
        "p:kiosk",
        "p:sldAll",
        "p:sldRg",
        "p:custShow",
        "p:penClr",
        "p:extLst",
    )
    showTypeChoice = ZeroOrOneChoice(
        (Choice("p:present"), Choice("p:browse"), Choice("p:kiosk")),
        successors=_tag_seq[3:],
    )
    slideRangeChoice = ZeroOrOneChoice(
        (Choice("p:sldAll"), Choice("p:sldRg"), Choice("p:custShow")),
        successors=_tag_seq[6:],
    )
    penClr: CT_Color | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:penClr", successors=_tag_seq[7:]
    )
    del _tag_seq

    loop: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "loop", XsdBoolean, default=False
    )
    showNarration: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "showNarration", XsdBoolean, default=False
    )
    showAnimation: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "showAnimation", XsdBoolean, default=True
    )
    useTimings: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "useTimings", XsdBoolean, default=True
    )


class CT_ShowInfoBrowse(BaseOxmlElement):
    """`p:browse` element, browsed-by-individual show-type settings."""

    showScrollbar: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "showScrollbar", XsdBoolean, default=True
    )


class CT_ShowInfoKiosk(BaseOxmlElement):
    """`p:kiosk` element, kiosk-mode show-type settings."""

    restart: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "restart", XsdUnsignedInt, default=5
    )


class CT_IndexRange(BaseOxmlElement):
    """`p:sldRg` element, a range of slides (1-based, inclusive) to show."""

    st: int = RequiredAttribute("st", XsdInt)  # pyright: ignore[reportAssignmentType]
    end: int = RequiredAttribute("end", XsdInt)  # pyright: ignore[reportAssignmentType]


class CT_CustomShowId(BaseOxmlElement):
    """`p:custShow` element, a reference to a custom show by id."""

    id: int = RequiredAttribute("id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
