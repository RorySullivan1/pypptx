"""Custom element classes for picture format effect elements on blips."""

from __future__ import annotations

from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import ST_Percentage
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute


class CT_LuminanceEffect(BaseOxmlElement):
    """`a:lum` element, specifying brightness and contrast adjustments on a blip."""

    bright: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "bright", ST_Percentage
    )
    contrast: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "contrast", ST_Percentage
    )


class CT_GrayscaleEffect(BaseOxmlElement):
    """`a:grayscl` element, converting a blip to grayscale."""


class CT_DuotoneEffect(BaseOxmlElement):
    """`a:duotone` element, applying a duotone color effect to a blip.

    Contains exactly two color child elements defining the duotone mapping.
    The first color replaces the darkest tones, the second replaces the lightest.
    """

    _COLOR_TAGS = frozenset(
        qn(t)
        for t in ("a:scrgbClr", "a:srgbClr", "a:hslClr", "a:sysClr", "a:schemeClr", "a:prstClr")
    )

    @property
    def color_elms(self) -> list[BaseOxmlElement]:
        """Return the (up to 2) color child elements in document order."""
        return [child for child in self if child.tag in self._COLOR_TAGS]
