"""Custom element classes for presentation section elements."""

from __future__ import annotations

from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrMore


class CT_SectionSlideIdListEntry(BaseOxmlElement):
    """`p14:sldId` element, a reference to a slide within a section."""

    id: str = OptionalAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_Section(BaseOxmlElement):
    """`p14:section` element, a named group of slides."""

    sldId_lst: list[CT_SectionSlideIdListEntry]

    sldId = ZeroOrMore("p14:sldId")
    name: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString
    )


class CT_SectionList(BaseOxmlElement):
    """`p14:sectionLst` element, container for presentation sections."""

    section_lst: list[CT_Section]

    section = ZeroOrMore("p14:section")
