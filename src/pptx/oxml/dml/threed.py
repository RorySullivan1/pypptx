"""Custom element classes for 3D formatting XML elements."""

from __future__ import annotations

from pptx.oxml.simpletypes import ST_Angle, ST_PositiveCoordinate, XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne


class CT_Bevel(BaseOxmlElement):
    """`a:bevelT` or `a:bevelB` element, specifying a bevel profile."""

    w: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w", ST_PositiveCoordinate
    )
    h: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "h", ST_PositiveCoordinate
    )
    prst: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "prst", XsdString
    )


class CT_Shape3D(BaseOxmlElement):
    """`a:sp3d` element, specifying 3D properties of a shape."""

    _tag_seq = ("a:bevelT", "a:bevelB", "a:extrusionClr", "a:contourClr", "a:extLst")
    bevelT: CT_Bevel | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:bevelT", successors=_tag_seq[1:]
    )
    bevelB: CT_Bevel | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:bevelB", successors=_tag_seq[2:]
    )
    del _tag_seq

    extrusionH: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "extrusionH", ST_PositiveCoordinate
    )
    contourW: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "contourW", ST_PositiveCoordinate
    )
    prstMaterial: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "prstMaterial", XsdString
    )


class CT_Camera(BaseOxmlElement):
    """`a:camera` element within `a:scene3d`, specifying camera properties."""

    prst: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "prst", XsdString
    )
    fov: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "fov", ST_Angle
    )


class CT_LightRig(BaseOxmlElement):
    """`a:lightRig` element within `a:scene3d`, specifying lighting."""

    rig: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rig", XsdString
    )
    dir: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dir", XsdString
    )


class CT_Scene3D(BaseOxmlElement):
    """`a:scene3d` element, specifying 3D scene properties."""

    _tag_seq = ("a:camera", "a:lightRig", "a:backdrop", "a:extLst")
    camera: CT_Camera | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:camera", successors=_tag_seq[1:]
    )
    lightRig: CT_LightRig | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:lightRig", successors=_tag_seq[2:]
    )
    del _tag_seq
