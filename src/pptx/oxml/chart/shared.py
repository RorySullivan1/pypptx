"""Shared oxml objects for charts."""

from __future__ import annotations

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import (
    ST_LayoutMode,
    XsdBoolean,
    XsdDouble,
    XsdString,
    XsdUnsignedInt,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)


class CT_Boolean(BaseOxmlElement):
    """
    Common complex type used for elements having a True/False value.
    """

    val = OptionalAttribute("val", XsdBoolean, default=True)


class CT_Boolean_Explicit(BaseOxmlElement):
    """Always spells out the `val` attribute, e.g. `val=1`.

    At least one boolean element is improperly interpreted by one or more
    versions of PowerPoint. The `c:overlay` element is interpreted as |False|
    when no `val` attribute is present, contrary to the behavior described in
    the schema. A remedy for this is to interpret a missing `val` attribute
    as |True| (consistent with the spec), but always write the attribute
    whenever there is occasion for changing the element.
    """

    _val = OptionalAttribute("val", XsdBoolean, default=True)

    @property
    def val(self):
        return self._val

    @val.setter
    def val(self, value):
        val_str = "1" if bool(value) is True else "0"
        self.set("val", val_str)


class CT_Double(BaseOxmlElement):
    """
    Used for floating point values.
    """

    val = RequiredAttribute("val", XsdDouble)


class CT_Layout(BaseOxmlElement):
    """
    ``<c:layout>`` custom element class
    """

    manualLayout = ZeroOrOne("c:manualLayout", successors=("c:extLst",))

    @property
    def horz_offset(self):
        """
        The float value in ./c:manualLayout/c:x when
        c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that XPath
        expression finds no match.
        """
        manualLayout = self.manualLayout
        if manualLayout is None:
            return 0.0
        return manualLayout.horz_offset

    @horz_offset.setter
    def horz_offset(self, offset):
        """
        Set the value of ./c:manualLayout/c:x@val to *offset* and
        ./c:manualLayout/c:xMode@val to "factor". Remove ./c:manualLayout if
        *offset* == 0.
        """
        if offset == 0.0:
            self._remove_manualLayout()
            return
        manualLayout = self.get_or_add_manualLayout()
        manualLayout.horz_offset = offset


class CT_LayoutMode(BaseOxmlElement):
    """
    Used for ``<c:xMode>``, ``<c:yMode>``, ``<c:wMode>``, and ``<c:hMode>``
    child elements of CT_ManualLayout.
    """

    val = OptionalAttribute("val", ST_LayoutMode, default=ST_LayoutMode.FACTOR)


class CT_ManualLayout(BaseOxmlElement):
    """
    ``<c:manualLayout>`` custom element class
    """

    _tag_seq = (
        "c:layoutTarget",
        "c:xMode",
        "c:yMode",
        "c:wMode",
        "c:hMode",
        "c:x",
        "c:y",
        "c:w",
        "c:h",
        "c:extLst",
    )
    xMode = ZeroOrOne("c:xMode", successors=_tag_seq[2:])
    yMode = ZeroOrOne("c:yMode", successors=_tag_seq[3:])
    wMode = ZeroOrOne("c:wMode", successors=_tag_seq[4:])
    hMode = ZeroOrOne("c:hMode", successors=_tag_seq[5:])
    x = ZeroOrOne("c:x", successors=_tag_seq[6:])
    y = ZeroOrOne("c:y", successors=_tag_seq[7:])
    w = ZeroOrOne("c:w", successors=_tag_seq[8:])
    h = ZeroOrOne("c:h", successors=_tag_seq[9:])
    del _tag_seq

    @property
    def horz_offset(self):
        """
        The float value in ./c:x@val when ./c:xMode@val == "factor". 0.0 when
        ./c:x is not present or ./c:xMode@val != "factor".
        """
        x, xMode = self.x, self.xMode
        if x is None or xMode is None or xMode.val != ST_LayoutMode.FACTOR:
            return 0.0
        return x.val

    @horz_offset.setter
    def horz_offset(self, offset):
        """
        Set the value of ./c:x@val to *offset* and ./c:xMode@val to "factor".
        """
        self.get_or_add_xMode().val = ST_LayoutMode.FACTOR
        self.get_or_add_x().val = offset

    def _get_layout_val(self, val_attr, mode_attr):
        """Return float value of a layout dimension, or None."""
        val_elm = getattr(self, val_attr)
        if val_elm is None:
            return None
        mode_elm = getattr(self, mode_attr)
        if mode_elm is not None and mode_elm.val != ST_LayoutMode.FACTOR:
            return None
        return val_elm.val

    def _set_layout_val(self, val_attr, mode_attr, value):
        """Set a layout dimension value, or remove it if None."""
        remover = getattr(self, f"_remove_{val_attr}")
        mode_remover = getattr(self, f"_remove_{mode_attr}")
        if value is None:
            remover()
            mode_remover()
            return
        adder = getattr(self, f"get_or_add_{mode_attr}")
        adder().val = ST_LayoutMode.FACTOR
        val_adder = getattr(self, f"get_or_add_{val_attr}")
        val_adder().val = value

    @property
    def left(self):
        """Float x position (0.0-1.0) as fraction of chart width, or None."""
        return self._get_layout_val("x", "xMode")

    @left.setter
    def left(self, value):
        self._set_layout_val("x", "xMode", value)

    @property
    def top(self):
        """Float y position (0.0-1.0) as fraction of chart height, or None."""
        return self._get_layout_val("y", "yMode")

    @top.setter
    def top(self, value):
        self._set_layout_val("y", "yMode", value)

    @property
    def width(self):
        """Float width (0.0-1.0) as fraction of chart width, or None."""
        return self._get_layout_val("w", "wMode")

    @width.setter
    def width(self, value):
        self._set_layout_val("w", "wMode", value)

    @property
    def height(self):
        """Float height (0.0-1.0) as fraction of chart height, or None."""
        return self._get_layout_val("h", "hMode")

    @height.setter
    def height(self, value):
        self._set_layout_val("h", "hMode", value)


class CT_NumFmt(BaseOxmlElement):
    """
    ``<c:numFmt>`` element specifying the formatting for number labels on a
    tick mark or data point.
    """

    formatCode = RequiredAttribute("formatCode", XsdString)
    sourceLinked = OptionalAttribute("sourceLinked", XsdBoolean)


class CT_Title(BaseOxmlElement):
    """`c:title` custom element class."""

    _tag_seq = ("c:tx", "c:layout", "c:overlay", "c:spPr", "c:txPr", "c:extLst")
    tx = ZeroOrOne("c:tx", successors=_tag_seq[1:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[4:])
    del _tag_seq

    def get_or_add_tx_rich(self):
        """Return `c:tx/c:rich`, newly created if not present.

        Return the `c:rich` grandchild at `c:tx/c:rich`. Both the `c:tx` and
        `c:rich` elements are created if not already present. Any
        `c:tx/c:strRef` element is removed. (Such an element would contain
        a cell reference for the axis title text in the chart's Excel
        worksheet.)
        """
        tx = self.get_or_add_tx()
        tx._remove_strRef()
        return tx.get_or_add_rich()

    @property
    def tx_rich(self):
        """Return `c:tx/c:rich` or |None| if not present."""
        richs = self.xpath("c:tx/c:rich")
        if not richs:
            return None
        return richs[0]

    @staticmethod
    def new_title():
        """Return "loose" `c:title` element containing default children."""
        return parse_xml(
            "<c:title %s>" "  <c:layout/>" '  <c:overlay val="0"/>' "</c:title>" % nsdecls("c")
        )


class CT_Tx(BaseOxmlElement):
    """
    ``<c:tx>`` element containing the text for a label on a data point or
    other chart item.
    """

    strRef = ZeroOrOne("c:strRef")
    rich = ZeroOrOne("c:rich")

    def _new_rich(self):
        return parse_xml(
            "<c:rich %s>"
            "  <a:bodyPr/>"
            "  <a:lstStyle/>"
            "  <a:p>"
            "    <a:pPr>"
            "      <a:defRPr/>"
            "    </a:pPr>"
            "  </a:p>"
            "</c:rich>" % nsdecls("c", "a")
        )


class CT_UnsignedInt(BaseOxmlElement):
    """
    ``<c:idx>`` element and others.
    """

    val = RequiredAttribute("val", XsdUnsignedInt)
