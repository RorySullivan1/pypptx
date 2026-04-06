"""Tags part, corresponds to ``ppt/tags/tags[N].xml`` parts in package."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.tags import CT_TagList

if TYPE_CHECKING:
    from pptx.package import Package


class TagsPart(XmlPart):
    """Corresponds to a ``ppt/tags/tags[N].xml`` part.

    Contains key-value string pairs associated with a shape or slide.
    """

    _element: CT_TagList

    partname_template = "/ppt/tags/tags%d.xml"

    @classmethod
    def new(cls, package: Package) -> TagsPart:
        """Return a new empty |TagsPart|."""
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        partname = package.next_partname(cls.partname_template)
        tagLst_xml = '<p:tagLst %s/>' % nsdecls("p", "r")
        tagLst = parse_xml(tagLst_xml)
        return cls(partname, CT.PML_TAGS, package, tagLst)

    def __getitem__(self, key: str) -> str:
        """Return value for `key`. Raises KeyError if not found."""
        for tag in self._element.tag_lst:
            if tag.name == key:
                return tag.val
        raise KeyError(key)

    def __setitem__(self, key: str, value: str) -> None:
        """Set `key` to `value`, overwriting if it already exists."""
        for tag in self._element.tag_lst:
            if tag.name == key:
                tag.val = value
                return
        new_tag = self._element._add_tag()  # pyright: ignore[reportPrivateUsage]
        new_tag.name = key
        new_tag.val = value

    def __delitem__(self, key: str) -> None:
        """Remove tag with `key`. Raises KeyError if not found."""
        for tag in self._element.tag_lst:
            if tag.name == key:
                self._element.remove(tag)
                return
        raise KeyError(key)

    def __contains__(self, key: object) -> bool:
        """Return True if `key` is present."""
        if not isinstance(key, str):
            return False
        return any(tag.name == key for tag in self._element.tag_lst)

    def __len__(self) -> int:
        return len(self._element.tag_lst)

    def __iter__(self):
        """Iterate over tag names."""
        for tag in self._element.tag_lst:
            yield tag.name

    def items(self) -> list[tuple[str, str]]:
        """Return list of (name, value) pairs."""
        return [(tag.name, tag.val) for tag in self._element.tag_lst]

    def get(self, key: str, default: str | None = None) -> str | None:
        """Return value for `key`, or `default` if not found."""
        for tag in self._element.tag_lst:
            if tag.name == key:
                return tag.val
        return default
