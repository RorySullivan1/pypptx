"""Custom element classes for tag-related XML elements."""

from __future__ import annotations

from typing import Callable

from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, RequiredAttribute, ZeroOrMore


class CT_StringTag(BaseOxmlElement):
    """`p:tag` element, a single name-value string pair."""

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    val: str = RequiredAttribute("val", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_TagList(BaseOxmlElement):
    """`p:tagLst` element, container for string tags on a shape or slide."""

    _add_tag: Callable[..., CT_StringTag]
    tag_lst: list[CT_StringTag]

    tag = ZeroOrMore("p:tag")
