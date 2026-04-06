"""Visual effects on a shape such as shadow, glow, and reflection."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx.enum.dml import MSO_RECT_ALIGNMENT
    from pptx.oxml.dml.effect import (
        CT_GlowEffect,
        CT_InnerShadowEffect,
        CT_OuterShadowEffect,
        CT_ReflectionEffect,
    )
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.util import Length


class ShadowFormat:
    """Provides access to shadow effect on a shape."""

    def __init__(self, spPr: BaseOxmlElement) -> None:
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
    def alignment(self, value: MSO_RECT_ALIGNMENT | None) -> None:
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
    def blur_radius(self, value: Length | None) -> None:
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
    def direction(self, value: float | None) -> None:
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
    def distance(self, value: Length | None) -> None:
        outerShdw = self._get_or_add_outerShdw()
        outerShdw.dist = value

    @property
    def inherit(self) -> bool:
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
    def inherit(self, value: bool) -> None:
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
    def rotate_with_shape(self, value: bool | None) -> None:
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

    def _get_or_add_outerShdw(self) -> CT_OuterShadowEffect:
        """Return the `a:outerShdw` element, creating effectLst and outerShdw if needed."""
        effectLst = self._element.get_or_add_effectLst()
        return effectLst.get_or_add_outerShdw()

    @property
    def _innerShdw(self) -> CT_InnerShadowEffect | None:
        """Return `a:innerShdw` element or None."""
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.innerShdw

    @property
    def _outerShdw(self) -> CT_OuterShadowEffect | None:
        """Return `a:outerShdw` element or None."""
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.outerShdw


class GlowFormat:
    """Provides access to glow effect on a shape."""

    def __init__(self, spPr: BaseOxmlElement) -> None:
        self._element = spPr

    @property
    def enabled(self) -> bool:
        """True when a glow effect is explicitly defined."""
        effectLst = self._element.effectLst
        return effectLst is not None and effectLst.glow is not None

    @property
    def radius(self) -> Length | None:
        """Radius of the glow in EMU.

        Read/write. |None| when no glow is defined.
        """
        effectLst = self._element.effectLst
        if effectLst is None or effectLst.glow is None:
            return None
        return effectLst.glow.rad

    @radius.setter
    def radius(self, value: Length | None) -> None:
        effectLst = self._element.get_or_add_effectLst()
        if value is None:
            effectLst._remove_glow()
            return
        glow = effectLst.get_or_add_glow()
        glow.rad = value

    def clear(self) -> None:
        """Remove any explicitly-defined glow effect."""
        effectLst = self._element.effectLst
        if effectLst is not None:
            effectLst._remove_glow()


class ReflectionFormat:
    """Provides access to reflection effect on a shape."""

    def __init__(self, spPr: BaseOxmlElement) -> None:
        self._element = spPr

    @property
    def enabled(self) -> bool:
        """True when a reflection effect is explicitly defined."""
        effectLst = self._element.effectLst
        return effectLst is not None and effectLst.reflection is not None

    @property
    def blur_radius(self) -> Length | None:
        """Blur radius of the reflection in EMU."""
        reflection = self._reflection
        if reflection is None:
            return None
        return reflection.blurRad

    @blur_radius.setter
    def blur_radius(self, value: Length | None) -> None:
        reflection = self._get_or_add_reflection()
        reflection.blurRad = value

    @property
    def start_opacity(self) -> float | None:
        """Starting opacity of the reflection (0.0 to 1.0)."""
        reflection = self._reflection
        if reflection is None:
            return None
        return reflection.stA

    @start_opacity.setter
    def start_opacity(self, value: float | None) -> None:
        reflection = self._get_or_add_reflection()
        reflection.stA = value

    @property
    def end_opacity(self) -> float | None:
        """Ending opacity of the reflection (0.0 to 1.0)."""
        reflection = self._reflection
        if reflection is None:
            return None
        return reflection.endA

    @end_opacity.setter
    def end_opacity(self, value: float | None) -> None:
        reflection = self._get_or_add_reflection()
        reflection.endA = value

    @property
    def distance(self) -> Length | None:
        """Distance of the reflection from the shape in EMU."""
        reflection = self._reflection
        if reflection is None:
            return None
        return reflection.dist

    @distance.setter
    def distance(self, value: Length | None) -> None:
        reflection = self._get_or_add_reflection()
        reflection.dist = value

    def clear(self) -> None:
        """Remove any explicitly-defined reflection effect."""
        effectLst = self._element.effectLst
        if effectLst is not None:
            effectLst._remove_reflection()

    def _get_or_add_reflection(self) -> CT_ReflectionEffect:
        effectLst = self._element.get_or_add_effectLst()
        return effectLst.get_or_add_reflection()

    @property
    def _reflection(self) -> CT_ReflectionEffect | None:
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.reflection


class SoftEdgeFormat:
    """Provides access to soft edge effect on a shape."""

    def __init__(self, spPr: BaseOxmlElement) -> None:
        self._element = spPr

    @property
    def enabled(self) -> bool:
        """True when a soft edge effect is explicitly defined."""
        effectLst = self._element.effectLst
        return effectLst is not None and effectLst.softEdge is not None

    @property
    def radius(self) -> Length | None:
        """Radius of the soft edge in EMU.

        Read/write. |None| when no soft edge is defined.
        """
        effectLst = self._element.effectLst
        if effectLst is None or effectLst.softEdge is None:
            return None
        return effectLst.softEdge.rad

    @radius.setter
    def radius(self, value: Length | None) -> None:
        effectLst = self._element.get_or_add_effectLst()
        if value is None:
            effectLst._remove_softEdge()
            return
        softEdge = effectLst.get_or_add_softEdge()
        softEdge.rad = value

    def clear(self) -> None:
        """Remove any explicitly-defined soft edge effect."""
        effectLst = self._element.effectLst
        if effectLst is not None:
            effectLst._remove_softEdge()
