"""Presentation properties part, corresponds to ``ppt/presProps.xml`` in package."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import LazyXmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.presprops import CT_PresentationProperties

if TYPE_CHECKING:
    from pptx.package import Package


class PresPropsPart(LazyXmlPart):
    """Corresponds to part named ``ppt/presProps.xml``.

    Contains slide-show settings for the presentation: loop until Esc, show type
    (present/browse/kiosk), narration/animation/timings toggles, pen colour, and the
    slide range shown (all slides, a from-to range, or a custom show).

    Loaded lazily -- see |LazyXmlPart| -- so a package that is opened and re-saved
    without ever reading `Presentation.slide_show_settings` writes this part's bytes
    back unchanged.
    """

    _element: CT_PresentationProperties  # pyright: ignore[reportIncompatibleMethodOverride]

    partname_template = "/ppt/presProps.xml"

    @classmethod
    def new(cls, package: Package) -> PresPropsPart:
        """Return a new, minimal |PresPropsPart|, containing no `p:showPr`."""
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        presentationPr_xml = "<p:presentationPr %s/>" % nsdecls("p", "a", "r")
        element = parse_xml(presentationPr_xml)
        return cls(
            PackURI(cls.partname_template), CT.PML_PRES_PROPS, package, element=element
        )
