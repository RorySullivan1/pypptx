"""Custom properties part, corresponds to ``/docProps/custom.xml`` part in package."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.custprops import CT_CustomProperties

if TYPE_CHECKING:
    from pptx.package import Package


class CustomPropertiesPart(XmlPart):
    """Corresponds to part named ``/docProps/custom.xml``.

    Contains custom document properties (key-value metadata) for this package.
    """

    _element: CT_CustomProperties

    @classmethod
    def default(cls, package: Package) -> CustomPropertiesPart:
        """Return a new empty |CustomPropertiesPart|."""
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        xml = '<cust:Properties %s %s/>' % (nsdecls("cust"), nsdecls("vt"))
        element = parse_xml(xml)
        return cls(PackURI("/docProps/custom.xml"), CT.OFC_CUSTOM_PROPERTIES, package, element)

    def __getitem__(self, name: str) -> str | int | float | bool | None:
        """Return value for property `name`. Raises KeyError if not found."""
        prop = self._element.get_property(name)
        if prop is None:
            raise KeyError(name)
        return prop.value

    def __setitem__(self, name: str, value: str | int | float | bool) -> None:
        """Set property `name` to `value`, creating or updating as needed."""
        prop = self._element.get_property(name)
        if prop is not None:
            prop.value = value
        else:
            next_pid = self._element._next_pid
            new_prop = self._element._add_custprop()  # pyright: ignore[reportPrivateUsage]
            new_prop.fmtid = CT_CustomProperties._FMTID
            new_prop.pid = next_pid
            new_prop.name = name
            new_prop.value = value

    def __delitem__(self, name: str) -> None:
        """Remove property `name`. Raises KeyError if not found."""
        prop = self._element.get_property(name)
        if prop is None:
            raise KeyError(name)
        self._element.remove(prop)

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        return self._element.get_property(name) is not None

    def __len__(self) -> int:
        return len(self._element.custprop_lst)

    def __iter__(self) -> Iterator[str]:
        """Iterate over property names."""
        for prop in self._element.custprop_lst:
            yield prop.name

    def items(self) -> list[tuple[str, str | int | float | bool | None]]:
        """Return list of (name, value) pairs."""
        return [(prop.name, prop.value) for prop in self._element.custprop_lst]

    def get(
        self, name: str, default: str | int | float | bool | None = None
    ) -> str | int | float | bool | None:
        """Return value for `name`, or `default` if not found."""
        prop = self._element.get_property(name)
        if prop is None:
            return default
        return prop.value
