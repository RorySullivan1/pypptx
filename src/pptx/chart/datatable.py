"""Data table of a chart."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx.oxml.chart.datatable import CT_DTable


class DataTable:
    """Represents the data table of a chart.

    A data table is a grid, appearing directly below the plot area, that displays the
    numeric values underlying the chart. A chart can have at most one data table; use
    |Chart.has_data_table| to add or remove it.
    """

    def __init__(self, dTable: CT_DTable) -> None:
        super(DataTable, self).__init__()
        self._element = dTable

    @property
    def horizontal_border(self) -> bool:
        """Read/write boolean specifying whether horizontal borders are drawn on the data
        table.

        Defaults to |True|, matching the setting PowerPoint uses when a data table is
        first added to a chart.
        """
        showHorzBorder = self._element.showHorzBorder
        if showHorzBorder is None:
            return True
        return showHorzBorder.val

    @horizontal_border.setter
    def horizontal_border(self, value: bool) -> None:
        self._element.get_or_add_showHorzBorder().val = bool(value)

    @property
    def vertical_border(self) -> bool:
        """Read/write boolean specifying whether vertical borders are drawn on the data
        table.

        Defaults to |True|, matching the setting PowerPoint uses when a data table is
        first added to a chart.
        """
        showVertBorder = self._element.showVertBorder
        if showVertBorder is None:
            return True
        return showVertBorder.val

    @vertical_border.setter
    def vertical_border(self, value: bool) -> None:
        self._element.get_or_add_showVertBorder().val = bool(value)

    @property
    def outline(self) -> bool:
        """Read/write boolean specifying whether an outline border is drawn around the
        data table.

        Defaults to |True|, matching the setting PowerPoint uses when a data table is
        first added to a chart.
        """
        showOutline = self._element.showOutline
        if showOutline is None:
            return True
        return showOutline.val

    @outline.setter
    def outline(self, value: bool) -> None:
        self._element.get_or_add_showOutline().val = bool(value)

    @property
    def show_keys(self) -> bool:
        """Read/write boolean specifying whether legend keys are shown next to each series
        name in the data table.

        Defaults to |True| per the schema when the `c:showKeys` element is absent.
        Note that when a data table is *added* via |Chart.has_data_table|, this value is
        explicitly set to |False|, matching PowerPoint's plain "Data Table" default (as
        opposed to "Data Table with Legend Keys").
        """
        showKeys = self._element.showKeys
        if showKeys is None:
            return True
        return showKeys.val

    @show_keys.setter
    def show_keys(self, value: bool) -> None:
        self._element.get_or_add_showKeys().val = bool(value)
