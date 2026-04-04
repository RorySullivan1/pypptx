"""Plot area-related chart API objects."""

from __future__ import annotations

from pptx.dml.chtfmt import ChartFormat
from pptx.shared import ElementProxy
from pptx.util import lazyproperty


class PlotArea(ElementProxy):
    """Provides access to plot area properties.

    Accessed via ``Chart.plot_area``.
    """

    def __init__(self, plotArea):
        super().__init__(plotArea)
        self._plotArea = plotArea

    @lazyproperty
    def format(self):
        """|ChartFormat| object providing access to shape formatting.

        Provides access to line and fill formatting for the plot area.
        """
        return ChartFormat(self._plotArea)

    @property
    def left(self):
        """Read/write float x position as fraction of chart width (0.0-1.0).

        |None| if no manual layout is defined.
        """
        layout = self._plotArea.layout
        if layout is None:
            return None
        manualLayout = layout.manualLayout
        if manualLayout is None:
            return None
        return manualLayout.left

    @left.setter
    def left(self, value):
        manualLayout = self._plotArea.get_or_add_layout().get_or_add_manualLayout()
        manualLayout.left = value

    @property
    def top(self):
        """Read/write float y position as fraction of chart height (0.0-1.0).

        |None| if no manual layout is defined.
        """
        layout = self._plotArea.layout
        if layout is None:
            return None
        manualLayout = layout.manualLayout
        if manualLayout is None:
            return None
        return manualLayout.top

    @top.setter
    def top(self, value):
        manualLayout = self._plotArea.get_or_add_layout().get_or_add_manualLayout()
        manualLayout.top = value

    @property
    def width(self):
        """Read/write float width as fraction of chart width (0.0-1.0).

        |None| if no manual layout is defined.
        """
        layout = self._plotArea.layout
        if layout is None:
            return None
        manualLayout = layout.manualLayout
        if manualLayout is None:
            return None
        return manualLayout.width

    @width.setter
    def width(self, value):
        manualLayout = self._plotArea.get_or_add_layout().get_or_add_manualLayout()
        manualLayout.width = value

    @property
    def height(self):
        """Read/write float height as fraction of chart height (0.0-1.0).

        |None| if no manual layout is defined.
        """
        layout = self._plotArea.layout
        if layout is None:
            return None
        manualLayout = layout.manualLayout
        if manualLayout is None:
            return None
        return manualLayout.height

    @height.setter
    def height(self, value):
        manualLayout = self._plotArea.get_or_add_layout().get_or_add_manualLayout()
        manualLayout.height = value
