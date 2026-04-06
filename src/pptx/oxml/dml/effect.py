"""Custom element classes for DrawingML effect-related XML elements."""

from __future__ import annotations

from typing import Callable

from pptx.enum.dml import MSO_RECT_ALIGNMENT
from pptx.oxml.simpletypes import (
    ST_Angle,
    ST_PositiveCoordinate,
    ST_PositiveFixedPercentage,
    XsdBoolean,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OptionalAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
)


class CT_OuterShadowEffect(BaseOxmlElement):
    """`a:outerShdw` element, specifying an outer shadow effect."""

    _tag_seq = (
        "a:scrgbClr",
        "a:srgbClr",
        "a:hslClr",
        "a:sysClr",
        "a:schemeClr",
        "a:prstClr",
    )
    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    del _tag_seq

    blurRad: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "blurRad", ST_PositiveCoordinate
    )
    dist: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dist", ST_PositiveCoordinate
    )
    dir: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dir", ST_Angle
    )
    algn: MSO_RECT_ALIGNMENT | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "algn", MSO_RECT_ALIGNMENT
    )
    rotWithShape: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rotWithShape", XsdBoolean
    )


class CT_InnerShadowEffect(BaseOxmlElement):
    """`a:innerShdw` element, specifying an inner shadow effect."""

    _tag_seq = (
        "a:scrgbClr",
        "a:srgbClr",
        "a:hslClr",
        "a:sysClr",
        "a:schemeClr",
        "a:prstClr",
    )
    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    del _tag_seq

    blurRad: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "blurRad", ST_PositiveCoordinate
    )
    dist: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dist", ST_PositiveCoordinate
    )
    dir: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dir", ST_Angle
    )


class CT_ReflectionEffect(BaseOxmlElement):
    """`a:reflection` element, specifying a reflection effect."""

    blurRad: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "blurRad", ST_PositiveCoordinate
    )
    stA: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "stA", ST_PositiveFixedPercentage
    )
    endA: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "endA", ST_PositiveFixedPercentage
    )
    dist: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dist", ST_PositiveCoordinate
    )
    dir: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dir", ST_Angle
    )
    rotWithShape: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rotWithShape", XsdBoolean
    )


class CT_GlowEffect(BaseOxmlElement):
    """`a:glow` element, specifying a glow effect.

    Contains a single color child element (e.g. `a:srgbClr`, `a:schemeClr`).
    """

    _tag_seq = (
        "a:scrgbClr",
        "a:srgbClr",
        "a:hslClr",
        "a:sysClr",
        "a:schemeClr",
        "a:prstClr",
    )
    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    del _tag_seq

    rad: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rad", ST_PositiveCoordinate
    )


class CT_SoftEdgesEffect(BaseOxmlElement):
    """`a:softEdge` element, specifying a soft edge effect."""

    rad: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rad", ST_PositiveCoordinate
    )


class CT_EffectList(BaseOxmlElement):
    """`a:effectLst` element, container for visual effects on a shape."""

    get_or_add_glow: Callable[[], CT_GlowEffect]
    get_or_add_innerShdw: Callable[[], CT_InnerShadowEffect]
    get_or_add_outerShdw: Callable[[], CT_OuterShadowEffect]
    get_or_add_reflection: Callable[[], CT_ReflectionEffect]
    get_or_add_softEdge: Callable[[], CT_SoftEdgesEffect]
    _remove_glow: Callable[[], None]
    _remove_innerShdw: Callable[[], None]
    _remove_outerShdw: Callable[[], None]
    _remove_reflection: Callable[[], None]
    _remove_softEdge: Callable[[], None]

    _tag_seq = (
        "a:blur",
        "a:fillOverlay",
        "a:glow",
        "a:innerShdw",
        "a:outerShdw",
        "a:prstShdw",
        "a:reflection",
        "a:softEdge",
    )
    glow: CT_GlowEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:glow", successors=_tag_seq[3:]
    )
    innerShdw: CT_InnerShadowEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:innerShdw", successors=_tag_seq[4:]
    )
    outerShdw: CT_OuterShadowEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:outerShdw", successors=_tag_seq[5:]
    )
    reflection: CT_ReflectionEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:reflection", successors=_tag_seq[7:]
    )
    softEdge: CT_SoftEdgesEffect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:softEdge", successors=()
    )
    del _tag_seq
