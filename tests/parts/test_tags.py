# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.tags` module."""

from __future__ import annotations

import pytest

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.parts.tags import TagsPart


class DescribeTagsPart:
    """Unit-test suite for `pptx.parts.tags.TagsPart`."""

    def _tags_part(self, tags: dict[str, str] | None = None):
        tag_xml_parts = ['<p:tagLst %s>' % nsdecls("p")]
        if tags:
            for name, val in tags.items():
                tag_xml_parts.append(f'  <p:tag name="{name}" val="{val}"/>')
        tag_xml_parts.append('</p:tagLst>')
        tagLst = parse_xml("".join(tag_xml_parts))
        return TagsPart("/ppt/tags/tags1.xml", "application/test", None, tagLst)

    def it_supports_getitem(self):
        tags = self._tags_part({"color": "red", "size": "large"})
        assert tags["color"] == "red"
        assert tags["size"] == "large"

    def it_raises_KeyError_on_missing_key(self):
        tags = self._tags_part()
        with pytest.raises(KeyError):
            tags["missing"]

    def it_supports_setitem_for_new_key(self):
        tags = self._tags_part()
        tags["color"] = "blue"
        assert tags["color"] == "blue"
        assert len(tags) == 1

    def it_supports_setitem_for_existing_key(self):
        tags = self._tags_part({"color": "red"})
        tags["color"] = "blue"
        assert tags["color"] == "blue"
        assert len(tags) == 1

    def it_supports_delitem(self):
        tags = self._tags_part({"color": "red", "size": "large"})
        del tags["color"]
        assert len(tags) == 1
        assert "color" not in tags

    def it_raises_KeyError_on_delitem_missing(self):
        tags = self._tags_part()
        with pytest.raises(KeyError):
            del tags["missing"]

    def it_supports_contains(self):
        tags = self._tags_part({"color": "red"})
        assert "color" in tags
        assert "missing" not in tags

    def it_supports_len(self):
        tags = self._tags_part({"a": "1", "b": "2"})
        assert len(tags) == 2

    def it_supports_iteration(self):
        tags = self._tags_part({"a": "1", "b": "2"})
        assert sorted(tags) == ["a", "b"]

    def it_provides_items(self):
        tags = self._tags_part({"color": "red"})
        assert tags.items() == [("color", "red")]

    def it_supports_get_with_default(self):
        tags = self._tags_part({"color": "red"})
        assert tags.get("color") == "red"
        assert tags.get("missing") is None
        assert tags.get("missing", "default") == "default"
