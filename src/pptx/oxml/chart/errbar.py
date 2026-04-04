"""Error bar-related oxml objects."""

from __future__ import annotations

from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrOne,
)


class CT_ErrBars(BaseOxmlElement):
    """``<c:errBars>`` element specifying error bars on a chart series.

    NOTE: The ``c:val`` child is not declared as ``ZeroOrOne`` because the tag
    ``c:val`` is already registered globally as ``CT_NumDataSource`` (for series
    value data). Error bar value access uses xpath in the API layer instead.
    """

    _tag_seq = (
        "c:errDir",
        "c:errBarType",
        "c:errValType",
        "c:noEndCap",
        "c:plus",
        "c:minus",
        "c:val",
        "c:spPr",
        "c:extLst",
    )
    errDir = ZeroOrOne("c:errDir", successors=_tag_seq[1:])
    errBarType = ZeroOrOne("c:errBarType", successors=_tag_seq[2:])
    errValType = ZeroOrOne("c:errValType", successors=_tag_seq[3:])
    noEndCap = ZeroOrOne("c:noEndCap", successors=_tag_seq[4:])
    plus = ZeroOrOne("c:plus", successors=_tag_seq[5:])
    minus = ZeroOrOne("c:minus", successors=_tag_seq[6:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[8:])
    del _tag_seq

    @property
    def val_val(self) -> float | None:
        """Return the float from the ``c:val/@val`` attribute, or None if not present."""
        vals = self.xpath("c:val/@val")
        if not vals:
            return None
        return float(vals[0])

    @val_val.setter
    def val_val(self, value: float | None):
        from pptx.oxml.xmlchemy import OxmlElement

        val_elms = self.xpath("c:val")
        if val_elms:
            self.remove(val_elms[0])
        if value is not None:
            val_elm = OxmlElement("c:val")
            val_elm.set("val", str(value))
            # insert before c:spPr if present, otherwise append
            spPr = self.spPr
            if spPr is not None:
                spPr.addprevious(val_elm)
            else:
                self.append(val_elm)


class CT_ErrDir(BaseOxmlElement):
    """``<c:errDir>`` element — error bar direction (x or y)."""

    val = OptionalAttribute("val", XsdString, default="y")


class CT_ErrBarType(BaseOxmlElement):
    """``<c:errBarType>`` element — which bars to show (both, plus, minus)."""

    val = OptionalAttribute("val", XsdString, default="both")


class CT_ErrValType(BaseOxmlElement):
    """``<c:errValType>`` element — how the error amount is determined."""

    val = OptionalAttribute("val", XsdString, default="fixedVal")
