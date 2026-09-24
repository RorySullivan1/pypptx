"""Custom element classes for the table styles (tableStyles.xml) part.

Covers `a:tblStyleLst` (the root element) and `a:tblStyle`, per ECMA-376
§20.1.4.2.22 (CT_TableStyleList) and §20.1.4.2.23 (CT_TableStyle). Only the id
and name of each defined style are modeled; authoring/editing the style
definitions themselves (fills, borders, fonts, banding, etc.) is out of scope.
"""

from __future__ import annotations

from typing import Callable

from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore


class CT_TableStyle(BaseOxmlElement):
    """`a:tblStyle` element, a single named table-style definition."""

    styleId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "styleId", XsdString
    )
    styleName: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "styleName", XsdString
    )


class CT_TableStyleList(BaseOxmlElement):
    """`a:tblStyleLst` element, root of the TableStylesPart, ``ppt/tableStyles.xml``."""

    _add_tblStyle: Callable[..., CT_TableStyle]
    tblStyle_lst: list[CT_TableStyle]

    tblStyle = ZeroOrMore("a:tblStyle")

    def_: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "def", XsdString
    )
