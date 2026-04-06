# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.custprops` module."""

from __future__ import annotations

import pytest

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.parts.custprops import CustomPropertiesPart


class DescribeCustomPropertiesPart:
    """Unit-test suite for `pptx.parts.custprops.CustomPropertiesPart`."""

    _FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

    def _part(self, props: dict[str, str] | None = None):
        xml_parts = ['<cust:Properties %s %s>' % (nsdecls("cust"), nsdecls("vt"))]
        pid = 2
        if props:
            for name, val in props.items():
                xml_parts.append(
                    f'<cust:property fmtid="{self._FMTID}" pid="{pid}" name="{name}">'
                    f"<vt:lpwstr>{val}</vt:lpwstr></cust:property>"
                )
                pid += 1
        xml_parts.append("</cust:Properties>")
        element = parse_xml("".join(xml_parts))
        return CustomPropertiesPart(
            "/docProps/custom.xml", "application/test", None, element
        )

    def it_supports_getitem(self):
        props = self._part({"Author": "Alice", "Version": "1.0"})
        assert props["Author"] == "Alice"
        assert props["Version"] == "1.0"

    def it_raises_KeyError_on_missing(self):
        props = self._part()
        with pytest.raises(KeyError):
            props["missing"]

    def it_supports_setitem_new(self):
        props = self._part()
        props["Author"] = "Bob"
        assert props["Author"] == "Bob"
        assert len(props) == 1

    def it_supports_setitem_update(self):
        props = self._part({"Author": "Alice"})
        props["Author"] = "Bob"
        assert props["Author"] == "Bob"
        assert len(props) == 1

    def it_supports_delitem(self):
        props = self._part({"Author": "Alice", "Version": "1.0"})
        del props["Author"]
        assert len(props) == 1
        assert "Author" not in props

    def it_supports_contains(self):
        props = self._part({"Author": "Alice"})
        assert "Author" in props
        assert "missing" not in props

    def it_supports_len(self):
        props = self._part({"a": "1", "b": "2"})
        assert len(props) == 2

    def it_supports_iteration(self):
        props = self._part({"a": "1", "b": "2"})
        assert sorted(props) == ["a", "b"]

    def it_provides_items(self):
        props = self._part({"Author": "Alice"})
        assert props.items() == [("Author", "Alice")]

    def it_supports_get_with_default(self):
        props = self._part({"Author": "Alice"})
        assert props.get("Author") == "Alice"
        assert props.get("missing") is None
        assert props.get("missing", "default") == "default"

    def it_handles_typed_values(self):
        xml = (
            '<cust:Properties %s %s>'
            '<cust:property fmtid="%s" pid="2" name="count"><vt:i4>42</vt:i4></cust:property>'
            '<cust:property fmtid="%s" pid="3" name="rate"><vt:r8>3.14</vt:r8></cust:property>'
            '<cust:property fmtid="%s" pid="4" name="flag"><vt:bool>true</vt:bool></cust:property>'
            "</cust:Properties>"
        ) % (nsdecls("cust"), nsdecls("vt"), self._FMTID, self._FMTID, self._FMTID)
        element = parse_xml(xml)
        props = CustomPropertiesPart("/docProps/custom.xml", "application/test", None, element)
        assert props["count"] == 42
        assert props["rate"] == 3.14
        assert props["flag"] is True

    def it_can_set_typed_values(self):
        props = self._part()
        props["count"] = 42
        props["flag"] = True
        props["rate"] = 3.14
        assert props["count"] == 42
        assert props["flag"] is True
        assert props["rate"] == 3.14
