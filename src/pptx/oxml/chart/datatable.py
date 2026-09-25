"""lxml custom element class for the chart data-table (`c:dTable`) element."""

from __future__ import annotations

from typing import cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.xmlchemy import BaseOxmlElement, ZeroOrOne


class CT_DTable(BaseOxmlElement):
    """`c:dTable` custom element class.

    `c:dTable` is an optional child of `c:plotArea`, appearing after any axis elements
    and before `c:spPr` and `c:extLst`. It causes a data table to be displayed below the
    plot area, showing the chart's underlying values.
    """

    _tag_seq = (
        "c:showHorzBorder",
        "c:showVertBorder",
        "c:showOutline",
        "c:showKeys",
        "c:spPr",
        "c:txPr",
        "c:extLst",
    )
    showHorzBorder = ZeroOrOne("c:showHorzBorder", successors=_tag_seq[1:])
    showVertBorder = ZeroOrOne("c:showVertBorder", successors=_tag_seq[2:])
    showOutline = ZeroOrOne("c:showOutline", successors=_tag_seq[3:])
    showKeys = ZeroOrOne("c:showKeys", successors=_tag_seq[4:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[5:])
    del _tag_seq

    @staticmethod
    def new_dTable() -> "CT_DTable":
        """Return a "loose" `c:dTable` element with PowerPoint's plain default settings.

        The schema default for each `showXxx` child, when omitted, is |True|. However,
        PowerPoint always writes all four boolean children explicitly when a data table
        is added via its UI ("Data Table" option, not "Data Table with Legend Keys"). That
        default has borders and an outline but no legend keys, i.e. `showHorzBorder=1`,
        `showVertBorder=1`, `showOutline=1`, `showKeys=0`. This method reproduces that
        behavior so a newly-added data table looks the same as one added by PowerPoint.
        """
        return cast(
            "CT_DTable",
            parse_xml(
                "<c:dTable %s>"
                '  <c:showHorzBorder val="1"/>'
                '  <c:showVertBorder val="1"/>'
                '  <c:showOutline val="1"/>'
                '  <c:showKeys val="0"/>'
                "</c:dTable>" % nsdecls("c")
            ),
        )
