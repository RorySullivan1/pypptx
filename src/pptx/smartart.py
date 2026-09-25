"""SmartArt (DrawingML diagram) objects: read the node tree and edit node text.

A |SmartArt| is reached from a graphic frame, `graphic_frame.smartart`, when
`graphic_frame.has_smartart` is |True|. Creating SmartArt, adding or removing nodes, and
changing its layout are not supported.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

if TYPE_CHECKING:
    from pptx.oxml.diagram import CT_DataModel, CT_Pt
    from pptx.parts.diagram import DiagramDataPart
    from pptx.parts.slide import BaseSlidePart


class SmartArt:
    """The content of a SmartArt graphic: a tree of |SmartArtNode| objects.

    Not intended to be constructed directly; use `GraphicFrame.smartart`.
    """

    def __init__(self, data_part: DiagramDataPart, host_part: BaseSlidePart):
        self._data_part = data_part
        self._host_part = host_part

    def iter_nodes(self) -> Iterator[SmartArtNode]:
        """Generate every node of the tree, depth-first in document order.

        A parent is generated before its children, so this is the order the nodes' text appears
        in PowerPoint's SmartArt text pane.
        """

        def iter_subtree(node: SmartArtNode) -> Iterator[SmartArtNode]:
            yield node
            for child in node.children:
                yield from iter_subtree(child)

        for node in self.nodes:
            yield from iter_subtree(node)

    @property
    def nodes(self) -> tuple[SmartArtNode, ...]:
        """The top-level nodes of this SmartArt, in order.

        Each node's descendants are reached through its `.children`. Assistant nodes (as in an
        organization chart) are included with the ordinary nodes.
        """
        doc_pt = self._data_model.doc_pt
        if doc_pt is None:
            return ()
        return tuple(
            SmartArtNode(pt, self, 0) for pt in self._data_model.child_pts(doc_pt.modelId)
        )

    @property
    def _data_model(self) -> CT_DataModel:
        return self._data_part.data_model

    def _child_nodes(self, node: SmartArtNode) -> tuple[SmartArtNode, ...]:
        return tuple(
            SmartArtNode(pt, self, node.level + 1)
            for pt in self._data_model.child_pts(node.model_id)
        )

    def _invalidate_drawing(self) -> None:
        """Drop PowerPoint's cached drawing of this diagram, which no longer matches the data.

        PowerPoint rebuilds the drawing from the data and layout parts when it opens the file.
        """
        rId = self._data_model.drawing_rId
        if rId is None:
            return
        self._data_model.remove_drawing_ref()
        self._host_part.drop_rel(rId)


class SmartArtNode:
    """One node of a SmartArt tree: a bullet in PowerPoint's SmartArt text pane.

    Not intended to be constructed directly; use `SmartArt.nodes` and `SmartArtNode.children`.
    """

    def __init__(self, pt: CT_Pt, smartart: SmartArt, level: int):
        self._pt = pt
        self._smartart = smartart
        self._level = level

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, SmartArtNode):
            return False
        return self._pt is other._pt

    def __ne__(self, other: object) -> bool:
        return not self == other

    def __hash__(self) -> int:
        return hash(self._pt)

    @property
    def children(self) -> tuple[SmartArtNode, ...]:
        """The nodes directly below this one, in order; an empty tuple for a leaf node."""
        return self._smartart._child_nodes(self)  # pyright: ignore[reportPrivateUsage]

    @property
    def level(self) -> int:
        """Zero-based depth of this node: 0 for a top-level node, 1 for its children, etc."""
        return self._level

    @property
    def model_id(self) -> str:
        """Identifier of this node within its SmartArt, e.g. "{3F2B...}" or "1"."""
        return self._pt.modelId

    @property
    def text(self) -> str:
        """The text of this node, "" when it has none.

        Read/write. As with `TextFrame.text`, a line-feed (`"\\n"`) separates paragraphs and a
        vertical-tab (`"\\v"`) stands for a line break.

        Assignment replaces all text of the node with a single run per paragraph, clears the
        node's placeholder (prompt-text) flag when the new text is not empty, and drops the
        drawing PowerPoint caches for the diagram so PowerPoint lays the SmartArt out again, with
        the new text, when it opens the file.
        """
        t = self._pt.t
        if t is None:
            return ""
        return "\n".join(p.text for p in t.p_lst)

    @text.setter
    def text(self, text: str) -> None:
        t = self._pt.get_or_add_t()
        t.clear_content()
        for p_text in text.split("\n"):
            t.add_p().append_text(p_text)
        prSet = self._pt.prSet
        if text and prSet is not None:
            prSet.phldr = None
        self._smartart._invalidate_drawing()  # pyright: ignore[reportPrivateUsage]
