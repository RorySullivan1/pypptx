"""SmartArt (DrawingML diagram) parts.

A SmartArt graphic frame relates its slide part to four diagram parts -- data, layout,
quick-style and colors -- and, in files saved by PowerPoint, a fifth *drawing* part caching the
shapes PowerPoint last laid out. All are loaded lazily (see |LazyXmlPart|) so SmartArt that is
opened and re-saved without being read writes its bytes back unchanged.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.package import LazyXmlPart

if TYPE_CHECKING:
    from pptx.oxml.diagram import CT_DataModel


class DiagramDataPart(LazyXmlPart):
    """Diagram-data part, e.g. ``/ppt/diagrams/data1.xml``, holding a `dgm:dataModel`.

    The content of the SmartArt: its points (nodes, with their text) and the connections that
    arrange them into a tree.
    """

    @property
    def data_model(self) -> CT_DataModel:
        """The `dgm:dataModel` root element of this part (parsed on first access)."""
        return self._element  # pyright: ignore[reportReturnType]


class DiagramLayoutPart(LazyXmlPart):
    """Diagram-layout definition part, e.g. ``/ppt/diagrams/layout1.xml``; held as-is."""


class DiagramStylePart(LazyXmlPart):
    """Diagram quick-style definition part, e.g. ``/ppt/diagrams/quickStyle1.xml``; held as-is."""


class DiagramColorsPart(LazyXmlPart):
    """Diagram colors definition part, e.g. ``/ppt/diagrams/colors1.xml``; held as-is."""


class DiagramDrawingPart(LazyXmlPart):
    """PowerPoint's cached drawing of a diagram, e.g. ``/ppt/diagrams/drawing1.xml``.

    A display cache: PowerPoint rebuilds it from the data and layout parts when it is absent.
    pypptx drops it when node text is edited so a stale drawing is never shown.
    """
