"""GroupShape and related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.exc import ShapeError, UnsupportedEffectError
from pptx.shapes.base import BaseShape
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.action import ActionSetting
    from pptx.oxml.shapes.groupshape import CT_GroupShape, CT_GroupShapeProperties
    from pptx.shapes.shapetree import GroupShapes
    from pptx.types import ProvidesPart


class GroupShape(BaseShape):
    """A shape that acts as a container for other shapes."""

    def __init__(self, grpSp: CT_GroupShape, parent: ProvidesPart):
        super().__init__(grpSp, parent)
        self._grpSp = grpSp

    @lazyproperty
    def click_action(self) -> ActionSetting:
        """Unconditionally raises `TypeError`.

        A group shape cannot have a click action or hover action.
        """
        raise ShapeError("a group shape cannot have a click action")

    @property
    def has_text_frame(self) -> bool:
        """Unconditionally |False|.

        A group shape does not have a textframe and cannot itself contain text. This does not
        impact the ability of shapes contained by the group to each have their own text.
        """
        return False

    @property
    def _effect_properties(self) -> CT_GroupShapeProperties:
        """A group's effects are on its `p:grpSpPr`."""
        return self._grpSp.grpSpPr

    @property
    def _three_d_properties(self) -> CT_GroupShapeProperties:
        """A group has no 3D shape properties (`p:grpSpPr` has no `a:sp3d`)."""
        raise UnsupportedEffectError("3D formatting is not supported on a group shape")

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """Member of :ref:`MsoShapeType` identifying the type of this shape.

        Unconditionally `MSO_SHAPE_TYPE.GROUP` in this case
        """
        return MSO_SHAPE_TYPE.GROUP

    @lazyproperty
    def shapes(self) -> GroupShapes:
        """|GroupShapes| object for this group.

        The |GroupShapes| object provides access to the group's member shapes and provides methods
        for adding new ones.
        """
        from pptx.shapes.shapetree import GroupShapes

        return GroupShapes(self._element, self)
