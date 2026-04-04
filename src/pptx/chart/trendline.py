"""Trendline-related objects."""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING

from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import XL_TRENDLINE_TYPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.chart.series import CT_SeriesComposite
    from pptx.oxml.chart.trendline import CT_Trendline


class Trendline:
    """Represents a single trendline on a chart series.

    A trendline shows the trend or direction of data in a series, such as a linear
    regression line or a moving average.
    """

    def __init__(self, trendline: CT_Trendline):
        self._element = trendline

    @lazyproperty
    def format(self) -> ChartFormat:
        """The |ChartFormat| instance for this trendline, providing access to line formatting."""
        return ChartFormat(self._element)

    @property
    def trendline_type(self) -> XL_TRENDLINE_TYPE:
        """Read/write. The type of this trendline as a member of :ref:`XlTrendlineType`."""
        trendlineType = self._element.trendlineType
        if trendlineType is None:
            return XL_TRENDLINE_TYPE.LINEAR
        xml_val = trendlineType.val
        return XL_TRENDLINE_TYPE.from_xml(xml_val)

    @trendline_type.setter
    def trendline_type(self, value: XL_TRENDLINE_TYPE):
        trendlineType = self._element.get_or_add_trendlineType()
        trendlineType.val = XL_TRENDLINE_TYPE.to_xml(value)

    @property
    def order(self) -> int:
        """Read/write int. The polynomial order (2-6). Only meaningful when type is POLYNOMIAL."""
        order = self._element.order
        if order is None:
            return 2
        return order.val

    @order.setter
    def order(self, value: int):
        self._element.get_or_add_order().val = value

    @property
    def period(self) -> int:
        """Read/write int. The moving average period (2+). Only meaningful when type is MOVING_AVERAGE."""
        period = self._element.period
        if period is None:
            return 2
        return period.val

    @period.setter
    def period(self, value: int):
        self._element.get_or_add_period().val = value

    @property
    def forward(self) -> float:
        """Read/write float. Number of periods to forecast forward."""
        forward = self._element.forward
        if forward is None:
            return 0.0
        return forward.val

    @forward.setter
    def forward(self, value: float):
        if value == 0.0:
            self._element._remove_forward()
            return
        self._element.get_or_add_forward().val = value

    @property
    def backward(self) -> float:
        """Read/write float. Number of periods to forecast backward."""
        backward = self._element.backward
        if backward is None:
            return 0.0
        return backward.val

    @backward.setter
    def backward(self, value: float):
        if value == 0.0:
            self._element._remove_backward()
            return
        self._element.get_or_add_backward().val = value

    @property
    def intercept(self) -> float | None:
        """Read/write float or None. The point where the trendline crosses the value axis."""
        intercept = self._element.intercept
        if intercept is None:
            return None
        return intercept.val

    @intercept.setter
    def intercept(self, value: float | None):
        if value is None:
            self._element._remove_intercept()
            return
        self._element.get_or_add_intercept().val = value

    @property
    def display_r_squared(self) -> bool:
        """Read/write boolean. True if the R-squared value is displayed on the chart."""
        dispRSqr = self._element.dispRSqr
        if dispRSqr is None:
            return False
        return dispRSqr.val

    @display_r_squared.setter
    def display_r_squared(self, value: bool):
        self._element.get_or_add_dispRSqr().val = value

    @property
    def display_equation(self) -> bool:
        """Read/write boolean. True if the equation for the trendline is displayed on the chart."""
        dispEq = self._element.dispEq
        if dispEq is None:
            return False
        return dispEq.val

    @display_equation.setter
    def display_equation(self, value: bool):
        self._element.get_or_add_dispEq().val = value

    @property
    def name(self) -> str | None:
        """Read/write string or None. A custom name for the trendline."""
        return self._element.trendline_name

    @name.setter
    def name(self, value: str | None):
        self._element.trendline_name = value


class TrendlineCollection(Sequence):
    """A sequence of |Trendline| objects on a series."""

    def __init__(self, ser: CT_SeriesComposite):
        self._ser = ser

    def __getitem__(self, index):
        return Trendline(self._ser.trendline_lst[index])

    def __len__(self) -> int:
        return len(self._ser.trendline_lst)

    def add(self, trendline_type: XL_TRENDLINE_TYPE = XL_TRENDLINE_TYPE.LINEAR) -> Trendline:
        """Add a new trendline to the series and return it.

        Args:
            trendline_type: The type of trendline, defaults to LINEAR.

        Returns:
            The newly created |Trendline| object.
        """
        trendline_elm = self._ser._add_trendline()
        trendlineType = OxmlElement("c:trendlineType")
        trendlineType.set("val", XL_TRENDLINE_TYPE.to_xml(trendline_type))
        trendline_elm.insert(0, trendlineType)
        return Trendline(trendline_elm)

    def remove(self, trendline: Trendline) -> None:
        """Remove the specified trendline from the series."""
        self._ser.remove(trendline._element)
