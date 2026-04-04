"""3D view and surface-related chart API objects."""

from __future__ import annotations

from pptx.dml.chtfmt import ChartFormat
from pptx.shared import ElementProxy
from pptx.util import lazyproperty


class View3D(ElementProxy):
    """Provides access to 3D view properties of a chart.

    Accessed via ``Chart.view_3d``.
    """

    def __init__(self, view3D):
        super().__init__(view3D)
        self._view3D = view3D

    @property
    def rot_x(self):
        """Read/write integer rotation about the X axis (-90 to 90 degrees).

        |None| indicates the default value should be used.
        """
        return self._view3D.rot_x_val

    @rot_x.setter
    def rot_x(self, value):
        self._view3D.rot_x_val = value

    @property
    def rot_y(self):
        """Read/write integer rotation about the Y axis (0 to 360 degrees).

        |None| indicates the default value should be used.
        """
        return self._view3D.rot_y_val

    @rot_y.setter
    def rot_y(self, value):
        self._view3D.rot_y_val = value

    @property
    def right_angle_axes(self):
        """Read/write boolean specifying whether axes are at right angles.

        |None| indicates the default value should be used.
        """
        return self._view3D.r_ang_ax_val

    @right_angle_axes.setter
    def right_angle_axes(self, value):
        self._view3D.r_ang_ax_val = value

    @property
    def perspective(self):
        """Read/write integer perspective angle (0 to 240).

        |None| indicates the default value should be used. Only effective
        when :attr:`right_angle_axes` is False.
        """
        return self._view3D.perspective_val

    @perspective.setter
    def perspective(self, value):
        self._view3D.perspective_val = value

    @property
    def depth_percent(self):
        """Read/write integer depth as percentage of chart width (20-2000).

        |None| indicates the default (100) should be used.
        """
        return self._view3D.depth_percent_val

    @depth_percent.setter
    def depth_percent(self, value):
        self._view3D.depth_percent_val = value

    @property
    def height_percent(self):
        """Read/write integer height as percentage of chart width (5-500).

        |None| indicates the default (100) should be used.
        """
        return self._view3D.h_percent_val

    @height_percent.setter
    def height_percent(self, value):
        self._view3D.h_percent_val = value


class ChartSurface(ElementProxy):
    """Provides access to floor or wall properties of a 3D chart.

    Accessed via ``Chart.floor``, ``Chart.back_wall``, or ``Chart.side_wall``.
    """

    def __init__(self, surface):
        super().__init__(surface)
        self._surface = surface

    @lazyproperty
    def format(self):
        """|ChartFormat| object providing access to shape formatting.

        Provides access to line and fill formatting for this surface.
        """
        return ChartFormat(self._surface)

    @property
    def thickness(self):
        """Read/write integer thickness of this surface.

        |None| indicates the default (0) should be used.
        """
        thickness = self._surface.thickness
        if thickness is None:
            return None
        return thickness.val

    @thickness.setter
    def thickness(self, value):
        self._surface._remove_thickness()
        if value is not None:
            self._surface._add_thickness(val=value)
