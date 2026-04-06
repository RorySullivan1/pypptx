"""Custom element classes for custom document properties."""

from __future__ import annotations

from typing import Callable

from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import XsdInt, XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore


class CT_CustomProperty(BaseOxmlElement):
    """`cust:property` element, a single custom property."""

    fmtid: str = RequiredAttribute("fmtid", XsdString)  # pyright: ignore[reportAssignmentType]
    pid: int = RequiredAttribute("pid", XsdInt)  # pyright: ignore[reportAssignmentType]
    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]

    @property
    def value(self) -> str | int | float | bool | None:
        """Return the typed value of this property."""
        for child in self:
            tag = child.tag
            if tag == qn("vt:lpwstr") or tag == qn("vt:lpstr"):
                return child.text or ""
            if tag == qn("vt:i4"):
                return int(child.text) if child.text else 0
            if tag == qn("vt:r8"):
                return float(child.text) if child.text else 0.0
            if tag == qn("vt:bool"):
                return child.text in ("true", "1") if child.text else False
        return None

    @value.setter
    def value(self, val: str | int | float | bool) -> None:
        """Set the value, replacing existing type child."""
        from lxml import etree

        # Remove existing value child
        for child in list(self):
            self.remove(child)
        # Add new typed child
        if isinstance(val, bool):
            child = etree.SubElement(self, qn("vt:bool"))
            child.text = "true" if val else "false"
        elif isinstance(val, int):
            child = etree.SubElement(self, qn("vt:i4"))
            child.text = str(val)
        elif isinstance(val, float):
            child = etree.SubElement(self, qn("vt:r8"))
            child.text = str(val)
        else:
            child = etree.SubElement(self, qn("vt:lpwstr"))
            child.text = str(val)


class CT_CustomProperties(BaseOxmlElement):
    """`cust:Properties` element, container for custom document properties."""

    _add_custprop: Callable[..., CT_CustomProperty]
    custprop_lst: list[CT_CustomProperty]

    custprop = ZeroOrMore("cust:property")

    _FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

    def get_property(self, name: str) -> CT_CustomProperty | None:
        """Return the property element with the given name, or None."""
        for prop in self.custprop_lst:
            if prop.name == name:
                return prop
        return None

    @property
    def _next_pid(self) -> int:
        """Return the next available property ID (starting at 2)."""
        used = [prop.pid for prop in self.custprop_lst]
        return max([1] + used) + 1
