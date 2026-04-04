"""Special chart line and bar API objects."""

from __future__ import annotations

from pptx.dml.chtfmt import ChartFormat
from pptx.shared import ElementProxy
from pptx.util import lazyproperty


class DropLines(ElementProxy):
    """Provides access to drop line formatting on a line or area chart."""

    @lazyproperty
    def format(self):
        """|ChartFormat| object providing access to line formatting."""
        return ChartFormat(self._element)


class HiLowLines(ElementProxy):
    """Provides access to high-low line formatting on a line or stock chart."""

    @lazyproperty
    def format(self):
        """|ChartFormat| object providing access to line formatting."""
        return ChartFormat(self._element)


class SeriesLines(ElementProxy):
    """Provides access to series connector line formatting on a bar or pie chart."""

    @lazyproperty
    def format(self):
        """|ChartFormat| object providing access to line formatting."""
        return ChartFormat(self._element)


class UpDownBars(ElementProxy):
    """Provides access to up/down bars on a line or stock chart."""

    def __init__(self, upDownBars):
        super().__init__(upDownBars)
        self._upDownBars = upDownBars

    @property
    def gap_width(self):
        """Read/write integer gap width between up/down bars as percentage.

        Default is 150 when not specified.
        """
        gapWidth = self._upDownBars.gapWidth
        if gapWidth is None:
            return 150
        return gapWidth.val

    @gap_width.setter
    def gap_width(self, value):
        gapWidth = self._upDownBars.get_or_add_gapWidth()
        gapWidth.val = value

    @lazyproperty
    def up_bars_format(self):
        """|ChartFormat| for the up bars (positive change)."""
        return ChartFormat(self._upDownBars.get_or_add_upBars())

    @lazyproperty
    def down_bars_format(self):
        """|ChartFormat| for the down bars (negative change)."""
        return ChartFormat(self._upDownBars.get_or_add_downBars())
