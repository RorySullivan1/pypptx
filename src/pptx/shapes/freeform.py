"""Objects related to construction of freeform shapes."""

from __future__ import annotations

import math
from typing import TYPE_CHECKING, Iterable, Iterator, Sequence

from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from typing_extensions import TypeAlias

    from pptx.oxml.shapes.autoshape import (
        CT_Path2D,
        CT_Path2DArcTo,
        CT_Path2DClose,
        CT_Path2DCubicBezierTo,
        CT_Path2DLineTo,
        CT_Path2DMoveTo,
        CT_Path2DQuadBezierTo,
        CT_Shape,
    )
    from pptx.shapes.shapetree import _BaseGroupShapes  # pyright: ignore[reportPrivateUsage]
    from pptx.util import Length

CT_DrawingOperation: TypeAlias = "CT_Path2DArcTo | CT_Path2DClose | CT_Path2DCubicBezierTo | CT_Path2DLineTo | CT_Path2DMoveTo | CT_Path2DQuadBezierTo"  # noqa: E501
DrawingOperation: TypeAlias = "_Arc | _Close | _CubicBezier | _LineSegment | _MoveTo | _QuadraticBezier"


class FreeformBuilder(Sequence[DrawingOperation]):
    """Allows a freeform shape to be specified and created.

    The initial pen position is provided on construction. From there, drawing proceeds using
    successive calls to draw line segments. The freeform shape may be closed by calling the
    :meth:`close` method.

    A shape may have more than one contour, in which case overlapping areas are "subtracted". A
    contour is a sequence of line segments beginning with a "move-to" operation. A move-to
    operation is automatically inserted in each new freeform; additional move-to ops can be
    inserted with the `.move_to()` method.
    """

    def __init__(
        self,
        shapes: _BaseGroupShapes,
        start_x: Length,
        start_y: Length,
        x_scale: float,
        y_scale: float,
    ):
        super(FreeformBuilder, self).__init__()
        self._shapes = shapes
        self._start_x = start_x
        self._start_y = start_y
        self._x_scale = x_scale
        self._y_scale = y_scale

    def __getitem__(  # pyright: ignore[reportIncompatibleMethodOverride]
        self, idx: int
    ) -> DrawingOperation:
        return self._drawing_operations.__getitem__(idx)

    def __iter__(self) -> Iterator[DrawingOperation]:
        return self._drawing_operations.__iter__()

    def __len__(self):
        return self._drawing_operations.__len__()

    @classmethod
    def new(
        cls,
        shapes: _BaseGroupShapes,
        start_x: float,
        start_y: float,
        x_scale: float,
        y_scale: float,
    ):
        """Return a new |FreeformBuilder| object.

        The initial pen location is specified (in local coordinates) by
        (`start_x`, `start_y`).
        """
        return cls(shapes, Emu(int(round(start_x))), Emu(int(round(start_y))), x_scale, y_scale)

    def add_arc(
        self, x_radius: float, y_radius: float, start_angle: float, swing_angle: float
    ) -> FreeformBuilder:
        """Add an elliptical arc segment starting from the current pen location.

        `x_radius` and `y_radius` are the horizontal and vertical radii, in local
        coordinates, of the ellipse the arc is swept on. Each is rounded to the nearest
        integer before use.

        `start_angle` locates the start of the arc on that ellipse and `swing_angle`
        specifies how far the arc sweeps from there, both in degrees. Angles increase
        *clockwise* starting from the ellipse's 3 o'clock position (its local +x axis); this
        matches the angle convention of the underlying `a:arcTo` XML element (and of the
        `rot` attribute used elsewhere in DrawingML). A positive `swing_angle` therefore
        sweeps clockwise and a negative one counter-clockwise; `swing_angle` may exceed +/-
        360 to sweep around the ellipse more than once.

        The pen ends at the point on the ellipse located at `start_angle + swing_angle`, which
        becomes the starting point for the next drawing operation.

        Returns this |FreeformBuilder| object so it can be used in chained calls.
        """
        self._drawing_operations.append(_Arc.new(self, x_radius, y_radius, start_angle, swing_angle))
        return self

    def add_cubic_bezier(
        self,
        control_point_1: tuple[float, float],
        control_point_2: tuple[float, float],
        end_point: tuple[float, float],
    ) -> FreeformBuilder:
        """Add a cubic Bezier curve segment ending at `end_point`.

        `control_point_1` and `control_point_2` are (x, y) pairs (2-tuples), in local
        coordinates, specifying the curve's two control points and `end_point` is an (x, y)
        pair specifying where the curve ends. Each coordinate is rounded to the nearest
        integer before use.

        Returns this |FreeformBuilder| object so it can be used in chained calls.
        """
        self._drawing_operations.append(
            _CubicBezier.new(self, control_point_1, control_point_2, end_point)
        )
        return self

    def add_quadratic_bezier(
        self, control_point: tuple[float, float], end_point: tuple[float, float]
    ) -> FreeformBuilder:
        """Add a quadratic Bezier curve segment ending at `end_point`.

        `control_point` is an (x, y) pair (2-tuple), in local coordinates, specifying the
        curve's control point and `end_point` is an (x, y) pair specifying where the curve
        ends. Each coordinate is rounded to the nearest integer before use.

        Returns this |FreeformBuilder| object so it can be used in chained calls.
        """
        self._drawing_operations.append(_QuadraticBezier.new(self, control_point, end_point))
        return self

    def add_line_segments(self, vertices: Iterable[tuple[float, float]], close: bool = True):
        """Add a straight line segment to each point in `vertices`.

        `vertices` must be an iterable of (x, y) pairs (2-tuples). Each x and y value is rounded
        to the nearest integer before use. The optional `close` parameter determines whether the
        resulting contour is `closed` or left `open`.

        Returns this |FreeformBuilder| object so it can be used in chained calls.
        """
        for x, y in vertices:
            self._add_line_segment(x, y)
        if close:
            self._add_close()
        return self

    def convert_to_shape(self, origin_x: Length = Emu(0), origin_y: Length = Emu(0)):
        """Return new freeform shape positioned relative to specified offset.

        `origin_x` and `origin_y` locate the origin of the local coordinate system in slide
        coordinates (EMU), perhaps most conveniently by use of a |Length| object.

        Note that this method may be called more than once to add multiple shapes of the same
        geometry in different locations on the slide.
        """
        sp = self._add_freeform_sp(origin_x, origin_y)
        path = self._start_path(sp)
        for drawing_operation in self:
            drawing_operation.apply_operation_to(path)
        self._shapes._invalidate_shape_cache()  # pyright: ignore[reportPrivateUsage]
        return self._shapes._shape_factory(sp)  # pyright: ignore[reportPrivateUsage]

    def move_to(self, x: float, y: float):
        """Move pen to (x, y) (local coordinates) without drawing line.

        Returns this |FreeformBuilder| object so it can be used in chained calls.
        """
        self._drawing_operations.append(_MoveTo.new(self, x, y))
        return self

    @property
    def shape_offset_x(self) -> Length:
        """Return x distance of shape origin from local coordinate origin.

        The returned integer represents the leftmost extent of the freeform shape, in local
        coordinates. Note that the bounding box of the shape need not start at the local origin.
        """
        min_x = self._start_x
        for drawing_operation in self:
            for x, _ in drawing_operation.points:
                min_x = min(min_x, x)
        return Emu(min_x)

    @property
    def shape_offset_y(self) -> Length:
        """Return y distance of shape origin from local coordinate origin.

        The returned integer represents the topmost extent of the freeform shape, in local
        coordinates. Note that the bounding box of the shape need not start at the local origin.
        """
        min_y = self._start_y
        for drawing_operation in self:
            for _, y in drawing_operation.points:
                min_y = min(min_y, y)
        return Emu(min_y)

    def _add_close(self):
        """Add a close |_Close| operation to the drawing sequence."""
        self._drawing_operations.append(_Close.new())

    def _add_freeform_sp(self, origin_x: Length, origin_y: Length):
        """Add a freeform `p:sp` element having no drawing elements.

        `origin_x` and `origin_y` are specified in slide coordinates, and represent the location
        of the local coordinates origin on the slide.
        """
        spTree = self._shapes._spTree  # pyright: ignore[reportPrivateUsage]
        return spTree.add_freeform_sp(
            origin_x + self._left, origin_y + self._top, self._width, self._height
        )

    def _add_line_segment(self, x: float, y: float) -> None:
        """Add a |_LineSegment| operation to the drawing sequence."""
        self._drawing_operations.append(_LineSegment.new(self, x, y))

    @property
    def _current_point(self) -> tuple[Length, Length]:
        """(x, y) location of the pen in local coordinates.

        This is the endpoint of the most-recently added drawing operation (`.close()`
        operations are skipped since they do not change the pen location), or the freeform's
        starting location if no drawing operations have been added yet.
        """
        for drawing_operation in reversed(self._drawing_operations):
            if isinstance(drawing_operation, _Close):
                continue
            return drawing_operation.x, drawing_operation.y
        return self._start_x, self._start_y

    @lazyproperty
    def _drawing_operations(self) -> list[DrawingOperation]:
        """Return the sequence of drawing operation objects for freeform."""
        return []

    @property
    def _dx(self) -> Length:
        """Return width of this shape's path in local units."""
        min_x = max_x = self._start_x
        for drawing_operation in self:
            for x, _ in drawing_operation.points:
                min_x = min(min_x, x)
                max_x = max(max_x, x)
        return Emu(max_x - min_x)

    @property
    def _dy(self) -> Length:
        """Return integer height of this shape's path in local units."""
        min_y = max_y = self._start_y
        for drawing_operation in self:
            for _, y in drawing_operation.points:
                min_y = min(min_y, y)
                max_y = max(max_y, y)
        return Emu(max_y - min_y)

    @property
    def _height(self):
        """Return vertical size of this shape's path in slide coordinates.

        This value is based on the actual extents of the shape and does not include any
        positioning offset.
        """
        return int(round(self._dy * self._y_scale))

    @property
    def _left(self):
        """Return leftmost extent of this shape's path in slide coordinates.

        Note that this value does not include any positioning offset; it assumes the drawing
        (local) coordinate origin is at (0, 0) on the slide.
        """
        return int(round(self.shape_offset_x * self._x_scale))

    def _local_to_shape(self, local_x: Length, local_y: Length) -> tuple[Length, Length]:
        """Translate local coordinates point to shape coordinates.

        Shape coordinates have the same unit as local coordinates, but are offset such that the
        origin of the shape coordinate system (0, 0) is located at the top-left corner of the
        shape bounding box.
        """
        return Emu(local_x - self.shape_offset_x), Emu(local_y - self.shape_offset_y)

    def _start_path(self, sp: CT_Shape) -> CT_Path2D:
        """Return a newly created `a:path` element added to `sp`.

        The returned `a:path` element has an `a:moveTo` element representing the shape starting
        point as its only child.
        """
        path = sp.add_path(w=self._dx, h=self._dy)
        path.add_moveTo(*self._local_to_shape(self._start_x, self._start_y))
        return path

    @property
    def _top(self):
        """Return topmost extent of this shape's path in slide coordinates.

        Note that this value does not include any positioning offset; it assumes the drawing
        (local) coordinate origin is located at slide coordinates (0, 0) (top-left corner of
        slide).
        """
        return int(round(self.shape_offset_y * self._y_scale))

    @property
    def _width(self):
        """Return width of this shape's path in slide coordinates.

        This value is based on the actual extents of the shape path and does not include any
        positioning offset.
        """
        return int(round(self._dx * self._x_scale))


class _BaseDrawingOperation:
    """Base class for freeform drawing operations.

    A drawing operation has at least one location (x, y) in local coordinates.
    """

    def __init__(self, freeform_builder: FreeformBuilder, x: Length, y: Length):
        super(_BaseDrawingOperation, self).__init__()
        self._freeform_builder = freeform_builder
        self._x = x
        self._y = y

    def apply_operation_to(self, path: CT_Path2D) -> CT_DrawingOperation:
        """Add the XML element(s) implementing this operation to `path`.

        Must be implemented by each subclass.
        """
        raise NotImplementedError("must be implemented by each subclass")

    @property
    def points(self) -> Sequence[tuple[Length, Length]]:
        """Sequence of (x, y) points, in local coordinates, that count toward this operation's
        contribution to the freeform shape's bounding extents.

        The default implementation contributes only the operation's target point. Subclasses
        with additional points relevant to the shape's extents (such as Bezier control points)
        override this to include them.
        """
        return ((self._x, self._y),)

    @property
    def x(self) -> Length:
        """Return the horizontal (x) target location of this operation.

        The returned value is an integer in local coordinates.
        """
        return self._x

    @property
    def y(self) -> Length:
        """Return the vertical (y) target location of this operation.

        The returned value is an integer in local coordinates.
        """
        return self._y


class _Arc(_BaseDrawingOperation):
    """Specifies an elliptical arc segment swept from the current pen position.

    The arc lies on an ellipse having the specified horizontal and vertical radii. Angles are
    measured in degrees and increase *clockwise* starting from the ellipse's 3 o'clock
    position (its local +x axis), matching the angle convention of the underlying `a:arcTo`
    XML element.
    """

    def __init__(
        self,
        freeform_builder: FreeformBuilder,
        x_radius: Length,
        y_radius: Length,
        start_angle: float,
        swing_angle: float,
        center_x: float,
        center_y: float,
        x: Length,
        y: Length,
    ):
        super(_Arc, self).__init__(freeform_builder, x, y)
        self._x_radius = x_radius
        self._y_radius = y_radius
        self._start_angle = start_angle
        self._swing_angle = swing_angle
        self._center_x = center_x
        self._center_y = center_y

    @classmethod
    def new(
        cls,
        freeform_builder: FreeformBuilder,
        x_radius: float,
        y_radius: float,
        start_angle: float,
        swing_angle: float,
    ) -> _Arc:
        """Return a new _Arc object swept from `freeform_builder`'s current pen position.

        `x_radius` and `y_radius` are rounded to the nearest integer before use. The center of
        the arc's ellipse and its end-point are computed from the builder's current pen
        location, the radii, and the angles.
        """
        x_radius_emu, y_radius_emu = Emu(int(round(x_radius))), Emu(int(round(y_radius)))
        start_x, start_y = freeform_builder._current_point  # pyright: ignore[reportPrivateUsage]

        start_theta = math.radians(start_angle)
        center_x = start_x - x_radius_emu * math.cos(start_theta)
        center_y = start_y - y_radius_emu * math.sin(start_theta)

        end_theta = math.radians(start_angle + swing_angle)
        end_x = center_x + x_radius_emu * math.cos(end_theta)
        end_y = center_y + y_radius_emu * math.sin(end_theta)

        return cls(
            freeform_builder,
            x_radius_emu,
            y_radius_emu,
            start_angle,
            swing_angle,
            center_x,
            center_y,
            Emu(int(round(end_x))),
            Emu(int(round(end_y))),
        )

    def apply_operation_to(self, path: CT_Path2D) -> CT_Path2DArcTo:
        """Add `a:arcTo` element to `path` for this arc segment.

        Returns the `a:arcTo` element newly added to the path. The radii are written
        unchanged; local-coordinate path points are scaled per-axis (independently in x and
        y) when the shape's overall bounding box is established, so the ellipse's aspect
        ratio is preserved without further adjustment here.
        """
        return path.add_arcTo(
            self._x_radius, self._y_radius, self._start_angle, self._swing_angle
        )

    @property
    def points(self) -> Sequence[tuple[Length, Length]]:
        """(x, y) points, in local coordinates, contributing to the shape's bounding extents.

        Includes the arc's end point along with any of the ellipse's axis-aligned extrema (at
        0, 90, 180, and 270 degrees) that fall within the swept angle range, so the computed
        bounding box captures the arc's true extent even where it bulges beyond a straight
        line drawn between its start and end points.
        """
        extrema = (0.0, 90.0, 180.0, 270.0)
        extrema_points = [
            self._point_on_ellipse(angle) for angle in extrema if self._sweeps_through(angle)
        ]
        return (*extrema_points, (self._x, self._y))

    def _point_on_ellipse(self, angle_degrees: float) -> tuple[Length, Length]:
        """(x, y) point, in local coordinates, at `angle_degrees` on this arc's ellipse."""
        theta = math.radians(angle_degrees)
        x = self._center_x + self._x_radius * math.cos(theta)
        y = self._center_y + self._y_radius * math.sin(theta)
        return Emu(int(round(x))), Emu(int(round(y)))

    def _sweeps_through(self, angle_degrees: float) -> bool:
        """True when `angle_degrees` (mod 360) falls within the arc's swept angle range."""
        lo = min(self._start_angle, self._start_angle + self._swing_angle)
        hi = max(self._start_angle, self._start_angle + self._swing_angle)
        if hi - lo >= 360:
            return True
        offset = (angle_degrees - lo) % 360
        return lo + offset <= hi


class _Close:
    """Specifies adding a `<a:close/>` element to the current contour."""

    @classmethod
    def new(cls) -> _Close:
        """Return a new _Close object."""
        return cls()

    def apply_operation_to(self, path: CT_Path2D) -> CT_Path2DClose:
        """Add `a:close` element to `path`."""
        return path.add_close()

    @property
    def points(self) -> Sequence[tuple[Length, Length]]:
        """A `.close()` operation does not move the pen and adds no new points."""
        return ()


class _CubicBezier(_BaseDrawingOperation):
    """Specifies a cubic Bezier curve segment ending at the specified point."""

    def __init__(
        self,
        freeform_builder: FreeformBuilder,
        control_point_1: tuple[Length, Length],
        control_point_2: tuple[Length, Length],
        x: Length,
        y: Length,
    ):
        super(_CubicBezier, self).__init__(freeform_builder, x, y)
        self._control_point_1 = control_point_1
        self._control_point_2 = control_point_2

    @classmethod
    def new(
        cls,
        freeform_builder: FreeformBuilder,
        control_point_1: tuple[float, float],
        control_point_2: tuple[float, float],
        end_point: tuple[float, float],
    ) -> _CubicBezier:
        """Return a new _CubicBezier object ending at `end_point`.

        `control_point_1`, `control_point_2`, and `end_point` are each (x, y) pairs whose
        coordinates are rounded to the nearest integer before use.
        """
        x1, y1 = control_point_1
        x2, y2 = control_point_2
        x3, y3 = end_point
        return cls(
            freeform_builder,
            (Emu(int(round(x1))), Emu(int(round(y1)))),
            (Emu(int(round(x2))), Emu(int(round(y2)))),
            Emu(int(round(x3))),
            Emu(int(round(y3))),
        )

    def apply_operation_to(self, path: CT_Path2D) -> CT_Path2DCubicBezierTo:
        """Add `a:cubicBezTo` element to `path` for this curve segment.

        Returns the `a:cubicBezTo` element newly added to the path.
        """
        offset_x = self._freeform_builder.shape_offset_x
        offset_y = self._freeform_builder.shape_offset_y
        x1, y1 = self._control_point_1
        x2, y2 = self._control_point_2
        return path.add_cubicBezTo(
            Emu(x1 - offset_x),
            Emu(y1 - offset_y),
            Emu(x2 - offset_x),
            Emu(y2 - offset_y),
            Emu(self._x - offset_x),
            Emu(self._y - offset_y),
        )

    @property
    def points(self) -> Sequence[tuple[Length, Length]]:
        """Both control points and the curve's end point, in local coordinates."""
        return (self._control_point_1, self._control_point_2, (self._x, self._y))


class _LineSegment(_BaseDrawingOperation):
    """Specifies a straight line segment ending at the specified point."""

    @classmethod
    def new(cls, freeform_builder: FreeformBuilder, x: float, y: float) -> _LineSegment:
        """Return a new _LineSegment object ending at point *(x, y)*.

        Both `x` and `y` are rounded to the nearest integer before use.
        """
        return cls(freeform_builder, Emu(int(round(x))), Emu(int(round(y))))

    def apply_operation_to(self, path: CT_Path2D) -> CT_Path2DLineTo:
        """Add `a:lnTo` element to `path` for this line segment.

        Returns the `a:lnTo` element newly added to the path.
        """
        return path.add_lnTo(
            Emu(self._x - self._freeform_builder.shape_offset_x),
            Emu(self._y - self._freeform_builder.shape_offset_y),
        )


class _MoveTo(_BaseDrawingOperation):
    """Specifies a new pen position."""

    @classmethod
    def new(cls, freeform_builder: FreeformBuilder, x: float, y: float) -> _MoveTo:
        """Return a new _MoveTo object for move to point `(x, y)`.

        Both `x` and `y` are rounded to the nearest integer before use.
        """
        return cls(freeform_builder, Emu(int(round(x))), Emu(int(round(y))))

    def apply_operation_to(self, path: CT_Path2D) -> CT_Path2DMoveTo:
        """Add `a:moveTo` element to `path` for this line segment."""
        return path.add_moveTo(
            Emu(self._x - self._freeform_builder.shape_offset_x),
            Emu(self._y - self._freeform_builder.shape_offset_y),
        )


class _QuadraticBezier(_BaseDrawingOperation):
    """Specifies a quadratic Bezier curve segment ending at the specified point."""

    def __init__(
        self,
        freeform_builder: FreeformBuilder,
        control_point: tuple[Length, Length],
        x: Length,
        y: Length,
    ):
        super(_QuadraticBezier, self).__init__(freeform_builder, x, y)
        self._control_point = control_point

    @classmethod
    def new(
        cls,
        freeform_builder: FreeformBuilder,
        control_point: tuple[float, float],
        end_point: tuple[float, float],
    ) -> _QuadraticBezier:
        """Return a new _QuadraticBezier object ending at `end_point`.

        `control_point` and `end_point` are each (x, y) pairs whose coordinates are rounded to
        the nearest integer before use.
        """
        x1, y1 = control_point
        x2, y2 = end_point
        return cls(
            freeform_builder,
            (Emu(int(round(x1))), Emu(int(round(y1)))),
            Emu(int(round(x2))),
            Emu(int(round(y2))),
        )

    def apply_operation_to(self, path: CT_Path2D) -> CT_Path2DQuadBezierTo:
        """Add `a:quadBezTo` element to `path` for this curve segment.

        Returns the `a:quadBezTo` element newly added to the path.
        """
        offset_x = self._freeform_builder.shape_offset_x
        offset_y = self._freeform_builder.shape_offset_y
        x1, y1 = self._control_point
        return path.add_quadBezTo(
            Emu(x1 - offset_x),
            Emu(y1 - offset_y),
            Emu(self._x - offset_x),
            Emu(self._y - offset_y),
        )

    @property
    def points(self) -> Sequence[tuple[Length, Length]]:
        """The control point and the curve's end point, in local coordinates."""
        return (self._control_point, (self._x, self._y))
