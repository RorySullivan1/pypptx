"""Special chart line and bar oxml objects."""

from __future__ import annotations

from pptx.oxml.simpletypes import ST_GapAmount
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne


class CT_UpDownBars(BaseOxmlElement):
    """`c:upDownBars` element, specifying up/down bars on a line or stock chart."""

    _tag_seq = ("c:gapWidth", "c:upBars", "c:downBars", "c:extLst")
    gapWidth = ZeroOrOne("c:gapWidth", successors=_tag_seq[1:])
    upBars = ZeroOrOne("c:upBars", successors=_tag_seq[2:])
    downBars = ZeroOrOne("c:downBars", successors=_tag_seq[3:])
    del _tag_seq


class CT_UpDownBar(BaseOxmlElement):
    """`c:upBars` and `c:downBars` elements — containers for spPr formatting."""

    spPr = ZeroOrOne("c:spPr", successors=())
