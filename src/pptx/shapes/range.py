# pyright: reportPrivateUsage=false

"""ShapeRange — bulk operations on a collection of shapes."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, overload

if TYPE_CHECKING:
    from collections.abc import Iterable

    from pptx.shapes.base import BaseShape
    from pptx.util import Length


class ShapeRange:
    """Enables bulk alignment, distribution, and property setting on multiple shapes.

    Wraps an ordered tuple of |BaseShape| instances and provides PowerPoint-style
    alignment and distribution operations that act on the bounding box of the collection.
    """

    def __init__(self, shapes: Iterable[BaseShape]) -> None:
        self._shapes = tuple(shapes)
        if len(self._shapes) < 1:
            raise ValueError("ShapeRange requires at least one shape")

    def __len__(self) -> int:
        return len(self._shapes)

    def __iter__(self) -> Iterator[BaseShape]:
        return iter(self._shapes)

    @overload
    def __getitem__(self, idx: int) -> BaseShape: ...

    @overload
    def __getitem__(self, idx: slice) -> tuple[BaseShape, ...]: ...

    def __getitem__(self, idx: int | slice) -> BaseShape | tuple[BaseShape, ...]:
        if isinstance(idx, slice):
            return self._shapes[idx]
        return self._shapes[idx]

    # --- bounding box (read-only) ---

    @property
    def bbox_left(self) -> Length:
        """Leftmost edge of the bounding box enclosing all shapes, in EMU."""
        return min(s._element.x for s in self._shapes)

    @property
    def bbox_top(self) -> Length:
        """Topmost edge of the bounding box enclosing all shapes, in EMU."""
        return min(s._element.y for s in self._shapes)

    @property
    def bbox_right(self) -> int:
        """Rightmost edge of the bounding box enclosing all shapes, in EMU."""
        return max(s._element.x + s._element.cx for s in self._shapes)

    @property
    def bbox_bottom(self) -> int:
        """Bottommost edge of the bounding box enclosing all shapes, in EMU."""
        return max(s._element.y + s._element.cy for s in self._shapes)

    @property
    def bbox_width(self) -> int:
        """Width of the bounding box enclosing all shapes, in EMU."""
        return self.bbox_right - self.bbox_left

    @property
    def bbox_height(self) -> int:
        """Height of the bounding box enclosing all shapes, in EMU."""
        return self.bbox_bottom - self.bbox_top

    # --- alignment ---

    def align_left(self) -> None:
        """Align the left edges of all shapes to the leftmost shape's left edge."""
        target = self.bbox_left
        for s in self._shapes:
            s._element.x = target

    def align_center(self) -> None:
        """Center all shapes horizontally within the bounding box."""
        center = self.bbox_left + self.bbox_width // 2
        for s in self._shapes:
            s._element.x = center - s._element.cx // 2

    def align_right(self) -> None:
        """Align the right edges of all shapes to the rightmost shape's right edge."""
        target = self.bbox_right
        for s in self._shapes:
            s._element.x = target - s._element.cx

    def align_top(self) -> None:
        """Align the top edges of all shapes to the topmost shape's top edge."""
        target = self.bbox_top
        for s in self._shapes:
            s._element.y = target

    def align_middle(self) -> None:
        """Center all shapes vertically within the bounding box."""
        middle = self.bbox_top + self.bbox_height // 2
        for s in self._shapes:
            s._element.y = middle - s._element.cy // 2

    def align_bottom(self) -> None:
        """Align the bottom edges of all shapes to the bottommost shape's bottom edge."""
        target = self.bbox_bottom
        for s in self._shapes:
            s._element.y = target - s._element.cy

    # --- distribution ---

    def distribute_horizontal(self) -> None:
        """Distribute shapes with equal horizontal spacing between them.

        Shapes are sorted by their left edge. The first and last shapes (by position)
        stay fixed; interior shapes are repositioned to create equal gaps.

        Raises ``ValueError`` if fewer than 3 shapes.
        """
        if len(self._shapes) < 3:
            raise ValueError("distribute_horizontal requires at least 3 shapes")
        sorted_shapes = sorted(self._shapes, key=lambda s: s._element.x)
        total_width = sum(s._element.cx for s in sorted_shapes)
        bbox_extent = sorted_shapes[-1]._element.x + sorted_shapes[-1]._element.cx - sorted_shapes[0]._element.x
        total_gap = bbox_extent - total_width
        gap = total_gap / (len(sorted_shapes) - 1)
        pos = float(sorted_shapes[0]._element.x)
        for s in sorted_shapes:
            s._element.x = round(pos)
            pos += s._element.cx + gap

    def distribute_vertical(self) -> None:
        """Distribute shapes with equal vertical spacing between them.

        Shapes are sorted by their top edge. The first and last shapes (by position)
        stay fixed; interior shapes are repositioned to create equal gaps.

        Raises ``ValueError`` if fewer than 3 shapes.
        """
        if len(self._shapes) < 3:
            raise ValueError("distribute_vertical requires at least 3 shapes")
        sorted_shapes = sorted(self._shapes, key=lambda s: s._element.y)
        total_height = sum(s._element.cy for s in sorted_shapes)
        bbox_extent = sorted_shapes[-1]._element.y + sorted_shapes[-1]._element.cy - sorted_shapes[0]._element.y
        total_gap = bbox_extent - total_height
        gap = total_gap / (len(sorted_shapes) - 1)
        pos = float(sorted_shapes[0]._element.y)
        for s in sorted_shapes:
            s._element.y = round(pos)
            pos += s._element.cy + gap

    # --- batch property setters ---

    def set_left(self, value: int | Length) -> None:
        """Set the left edge of every shape to *value* (EMU)."""
        for s in self._shapes:
            s._element.x = value

    def set_top(self, value: int | Length) -> None:
        """Set the top edge of every shape to *value* (EMU)."""
        for s in self._shapes:
            s._element.y = value

    def set_width(self, value: int | Length) -> None:
        """Set the width of every shape to *value* (EMU)."""
        for s in self._shapes:
            s._element.cx = value

    def set_height(self, value: int | Length) -> None:
        """Set the height of every shape to *value* (EMU)."""
        for s in self._shapes:
            s._element.cy = value

    def set_rotation(self, value: float) -> None:
        """Set the rotation of every shape to *value* degrees."""
        for s in self._shapes:
            s.rotation = value

    def set_hidden(self, value: bool) -> None:
        """Set the hidden state of every shape to *value*."""
        for s in self._shapes:
            s.hidden = value
