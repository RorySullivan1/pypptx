"""Table styles part, corresponds to ``ppt/tableStyles.xml`` in package."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import LazyXmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.tablestyles import CT_TableStyleList

if TYPE_CHECKING:
    from pptx.package import Package


class TableStylesPart(LazyXmlPart):
    """Corresponds to part named ``ppt/tableStyles.xml``.

    Lists the table styles (GUID id + display name) available for use in this
    presentation via `Table.table_style_id`, plus the id of the default style.
    Authoring/editing style definitions is out of scope.

    Loaded lazily -- see |LazyXmlPart| -- so a package that is opened and re-saved
    without ever reading `Presentation.table_styles` writes this part's bytes back
    unchanged.
    """

    _element: CT_TableStyleList  # pyright: ignore[reportIncompatibleMethodOverride]

    partname_template = "/ppt/tableStyles.xml"

    @classmethod
    def new(cls, package: Package) -> TableStylesPart:
        """Return a new, empty |TableStylesPart|, with no default style defined."""
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        tblStyleLst_xml = "<a:tblStyleLst %s/>" % nsdecls("a")
        element = parse_xml(tblStyleLst_xml)
        return cls(
            PackURI(cls.partname_template), CT.PML_TABLE_STYLES, package, element=element
        )
