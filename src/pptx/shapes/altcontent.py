"""Shapes PowerPoint wraps in `mc:AlternateContent`: 3D models, zooms, equations and others.

Newer shape kinds are written as an `mc:AlternateContent` element: an `mc:Choice` holding the
shape for a reader that understands it, and an `mc:Fallback` holding a stand-in (usually a
picture) for any other reader. pypptx lists each such shape in `slide.shapes` as an
|AlternateContentShape|, read-only, identified by `content_kind` and placed where its fallback is.
Chartex charts, also written this way, are graphic frames instead (see `GraphicFrame.chartex`).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Literal, cast

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.exc import ShapeError
from pptx.oxml.shapes.shared import alternate_content_of
from pptx.shapes.base import BaseShape

if TYPE_CHECKING:
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.util import Length

ContentKind = Literal["model3d", "zoom", "equation", "unknown"]

_MODEL3D_URI = "http://schemas.microsoft.com/office/drawing/2017/model3d"
_ZOOM_URIS = (
    "http://schemas.microsoft.com/office/powerpoint/2016/slidezoom",
    "http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom",
    "http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom",
)
_MATH_URI = "http://schemas.openxmlformats.org/officeDocument/2006/math"


class AlternateContentShape(BaseShape):
    """A shape wrapped in `mc:AlternateContent`, such as a 3D model, zoom or equation. Read-only.

    Its name, id, position and size are those of the fallback shape (what a reader without
    support for the shape shows), or of the choice shape when there is no fallback. Moving,
    removing and duplicating it through the shape collection act on the whole wrapper; changing
    its geometry or name raises |ShapeError|.
    """

    @property
    def content_kind(self) -> ContentKind:
        """What the wrapped shape is: "model3d", "zoom" (slide, section or summary zoom),
        "equation", or "unknown" for any other alternate content."""
        alternateContent = self.alternate_content_element
        namespaces = set(_required_namespace_uris(alternateContent))
        uris = set(
            cast("list[str]", alternateContent.xpath("./mc:Choice//a:graphicData/@uri"))
        ) | namespaces
        if _MODEL3D_URI in uris:
            return "model3d"
        if uris.intersection(_ZOOM_URIS):
            return "zoom"
        if alternateContent.xpath("./mc:Choice//*[namespace-uri()='%s']" % _MATH_URI):
            return "equation"
        return "unknown"

    @property
    def alternate_content_element(self) -> BaseOxmlElement:
        """The `mc:AlternateContent` element holding this shape, for direct XML access."""
        alternateContent = alternate_content_of(self._element)
        if alternateContent is None:  # pragma: no cover -- the shape factory ensures a wrapper
            raise ShapeError("shape is not wrapped in mc:AlternateContent")
        return alternateContent

    @property
    def has_fallback(self) -> bool:
        """|True| if the wrapper has an `mc:Fallback` shape, which then gives this shape's
        position and size."""
        return bool(self.alternate_content_element.xpath("./mc:Fallback/*"))

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """`MSO_SHAPE_TYPE.MODEL_3D` for a 3D model, |None| for other alternate content."""
        if self.content_kind == "model3d":
            return MSO_SHAPE_TYPE.MODEL_3D
        return None  # pyright: ignore[reportReturnType]

    # -- geometry and name come from the fallback; changing them would leave the choice (what
    # -- PowerPoint shows) behind, so they are read-only --

    @property
    def height(self) -> Length:
        """Height of the represented shape (the fallback, else the choice), in EMU. Read-only.

        Assigning raises |ShapeError|.
        """
        return super().height

    @height.setter
    def height(self, value: Length):
        raise ShapeError("alternate-content shape is read-only")

    @property
    def left(self) -> Length:
        """Distance from the slide's left edge to the represented shape, in EMU. Read-only.

        Assigning raises |ShapeError|.
        """
        return super().left

    @left.setter
    def left(self, value: Length):
        raise ShapeError("alternate-content shape is read-only")

    @property
    def name(self) -> str:
        """Name of the represented shape, as shown in the selection pane. Read-only.

        Assigning raises |ShapeError|.
        """
        return super().name

    @name.setter
    def name(self, value: str):
        raise ShapeError("alternate-content shape is read-only")

    @property
    def rotation(self) -> float:
        """Clockwise rotation of the represented shape, in degrees. Read-only.

        Assigning raises |ShapeError|.
        """
        return super().rotation

    @rotation.setter
    def rotation(self, value: float):
        raise ShapeError("alternate-content shape is read-only")

    @property
    def top(self) -> Length:
        """Distance from the slide's top edge to the represented shape, in EMU. Read-only.

        Assigning raises |ShapeError|.
        """
        return super().top

    @top.setter
    def top(self, value: Length):
        raise ShapeError("alternate-content shape is read-only")

    @property
    def width(self) -> Length:
        """Width of the represented shape, in EMU. Read-only.

        Assigning raises |ShapeError|.
        """
        return super().width

    @width.setter
    def width(self, value: Length):
        raise ShapeError("alternate-content shape is read-only")


def _required_namespace_uris(alternateContent: BaseOxmlElement):
    """Namespace URI of each prefix named in an `mc:Choice/@Requires` of `alternateContent`."""
    for choice in alternateContent.xpath("./mc:Choice"):
        for prefix in (choice.get("Requires") or "").split():
            uri = choice.nsmap.get(prefix)
            if uri is not None:
                yield uri
