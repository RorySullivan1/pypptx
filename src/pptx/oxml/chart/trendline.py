"""Trendline-related oxml objects."""

from __future__ import annotations

from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    OxmlElement,
    ZeroOrOne,
)


class CT_Trendline(BaseOxmlElement):
    """``<c:trendline>`` element specifying a trendline on a chart series."""

    _tag_seq = (
        "c:name",
        "c:spPr",
        "c:trendlineType",
        "c:order",
        "c:period",
        "c:forward",
        "c:backward",
        "c:intercept",
        "c:dispRSqr",
        "c:dispEq",
        "c:trendlineLbl",
        "c:extLst",
    )
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[2:])
    trendlineType = ZeroOrOne("c:trendlineType", successors=_tag_seq[3:])
    order = ZeroOrOne("c:order", successors=_tag_seq[4:])
    period = ZeroOrOne("c:period", successors=_tag_seq[5:])
    forward = ZeroOrOne("c:forward", successors=_tag_seq[6:])
    backward = ZeroOrOne("c:backward", successors=_tag_seq[7:])
    intercept = ZeroOrOne("c:intercept", successors=_tag_seq[8:])
    dispRSqr = ZeroOrOne("c:dispRSqr", successors=_tag_seq[9:])
    dispEq = ZeroOrOne("c:dispEq", successors=_tag_seq[10:])
    trendlineLbl = ZeroOrOne("c:trendlineLbl", successors=_tag_seq[11:])
    del _tag_seq

    @property
    def trendline_name(self) -> str | None:
        """Return the text content of the ``c:name`` child element, or None."""
        names = self.xpath("c:name")
        if not names:
            return None
        return names[0].text

    @trendline_name.setter
    def trendline_name(self, value: str | None):
        names = self.xpath("c:name")
        if names:
            self.remove(names[0])
        if value is not None:
            name_elm = OxmlElement("c:name")
            name_elm.text = value
            self.insert(0, name_elm)


class CT_TrendlineType(BaseOxmlElement):
    """``<c:trendlineType>`` element specifying the regression type.

    Valid values: exp, linear, log, movingAvg, poly, power.
    """

    val = OptionalAttribute("val", XsdString, default="linear")


class CT_TrendlineLabel(BaseOxmlElement):
    """``<c:trendlineLbl>`` element for trendline label formatting."""

    _tag_seq = (
        "c:layout",
        "c:tx",
        "c:numFmt",
        "c:spPr",
        "c:txPr",
        "c:extLst",
    )
    layout = ZeroOrOne("c:layout", successors=_tag_seq[1:])
    tx = ZeroOrOne("c:tx", successors=_tag_seq[2:])
    numFmt = ZeroOrOne("c:numFmt", successors=_tag_seq[3:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[4:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[5:])
    del _tag_seq
