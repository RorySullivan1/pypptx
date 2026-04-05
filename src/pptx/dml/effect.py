"""Visual effects on a shape such as shadow, glow, and reflection."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx.enum.dml import MSO_RECT_ALIGNMENT
    from pptx.util import Length


class ShadowFormat:
    """Provides access to shadow effect on a shape."""

    def __init__(self, spPr):
        # ---spPr may also be a grpSpPr; both have a:effectLst child---
        self._element = spPr

    @property
    def alignment(self) -> MSO_RECT_ALIGNMENT | None:
        """Shadow alignment (anchor point) for outer shadow.

        Read/write. A member of :ref:`MsoRectAlignment` or |None|.
        Only applicable to outer shadow; ignored for inner shadow.
        """
        outerShdw = self._outerShdw
        if outerShdw is None:
            return None
        return outerShdw.algn

    @alignment.setter
    def alignment(self, value: MSO_RECT_ALIGNMENT | None):
        outerShdw = self._get_or_add_outerShdw()
        outerShdw.algn = value

    @property
    def blur_radius(self) -> Length | None:
        """Blur radius of the shadow in EMU.

        Read/write. |None| indicates no blur radius is set.
        """
        shdw = self._outerShdw if self._outerShdw is not None else self._innerShdw
        if shdw is None:
            return None
        return shdw.blurRad

    @blur_radius.setter
    def blur_radius(self, value: Length | None):
        outerShdw = self._get_or_add_outerShdw()
        outerShdw.blurRad = value

    @property
    def direction(self) -> float | None:
        """Direction of the shadow offset in degrees (0-360).

        Read/write. 0 degrees is to the right, 90 is below, etc.
        |None| indicates the direction is not explicitly set.
        """
        shdw = self._outerShdw if self._outerShdw is not None else self._innerShdw
        if shdw is None:
            return None
        return shdw.dir

    @direction.setter
    def direction(self, value: float | None):
        outerShdw = self._get_or_add_outerShdw()
        outerShdw.dir = value

    @property
    def distance(self) -> Length | None:
        """Distance of the shadow from the shape edge in EMU.

        Read/write. |None| indicates no distance is set.
        """
        shdw = self._outerShdw if self._outerShdw is not None else self._innerShdw
        if shdw is None:
            return None
        return shdw.dist

    @distance.setter
    def distance(self, value: Length | None):
        outerShdw = self._get_or_add_outerShdw()
        outerShdw.dist = value

    @property
    def inherit(self):
        """True if shape inherits shadow settings.

        Read/write. An explicitly-defined shadow setting on a shape causes
        this property to return |False|. A shape with no explicitly-defined
        shadow setting inherits its shadow settings from the style hierarchy
        (and so returns |True|).

        Assigning |True| causes any explicitly-defined shadow setting to be
        removed and inheritance is restored. Note this has the side-effect of
        removing **all** explicitly-defined effects, such as glow and
        reflection, and restoring inheritance for all effects on the shape.
        Assigning |False| causes the inheritance link to be broken and **no**
        effects to appear on the shape.
        """
        if self._element.effectLst is None:
            return True
        return False

    @inherit.setter
    def inherit(self, value):
        inherit = bool(value)
        if inherit:
            # ---remove any explicitly-defined effects
            self._element._remove_effectLst()
        else:
            # ---ensure at least the effectLst element is present
            self._element.get_or_add_effectLst()

    @property
    def rotate_with_shape(self) -> bool | None:
        """Whether the shadow rotates when the shape is rotated.

        Read/write. Only applicable to outer shadow.
        """
        outerShdw = self._outerShdw
        if outerShdw is None:
            return None
        return outerShdw.rotWithShape

    @rotate_with_shape.setter
    def rotate_with_shape(self, value: bool | None):
        outerShdw = self._get_or_add_outerShdw()
        outerShdw.rotWithShape = value

    @property
    def shadow_type(self) -> str | None:
        """Type of shadow effect: ``"outer"``, ``"inner"``, or |None|.

        Read-only. |None| when no shadow is explicitly defined.
        """
        if self._outerShdw is not None:
            return "outer"
        if self._innerShdw is not None:
            return "inner"
        return None

    def _get_or_add_outerShdw(self):
        """Return the `a:outerShdw` element, creating effectLst and outerShdw if needed."""
        effectLst = self._element.get_or_add_effectLst()
        return effectLst.get_or_add_outerShdw()

    @property
    def _innerShdw(self):
        """Return `a:innerShdw` element or None."""
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.innerShdw

    @property
    def _outerShdw(self):
        """Return `a:outerShdw` element or None."""
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.outerShdw
