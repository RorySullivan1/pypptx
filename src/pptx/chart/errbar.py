"""Error bar-related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import XL_ERROR_BAR_DIRECTION, XL_ERROR_BAR_INCLUDE, XL_ERROR_BAR_TYPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.chart.errbar import CT_ErrBars
    from pptx.oxml.chart.series import CT_SeriesComposite


class ErrorBars:
    """Represents error bars on a chart series.

    Provides access to error bar type, direction, include mode, value, and formatting.
    """

    def __init__(self, errBars: CT_ErrBars):
        self._element = errBars

    @lazyproperty
    def format(self) -> ChartFormat:
        """The |ChartFormat| instance for these error bars, providing access to line formatting."""
        return ChartFormat(self._element)

    @property
    def type(self) -> XL_ERROR_BAR_TYPE:
        """Read/write. How the error amount is determined.

        A member of :ref:`XlErrorBarType`, e.g. ``XL_ERROR_BAR_TYPE.FIXED_VALUE``.
        """
        errValType = self._element.errValType
        if errValType is None:
            return XL_ERROR_BAR_TYPE.FIXED_VALUE
        return XL_ERROR_BAR_TYPE.from_xml(errValType.val)

    @type.setter
    def type(self, value: XL_ERROR_BAR_TYPE):
        self._element.get_or_add_errValType().val = XL_ERROR_BAR_TYPE.to_xml(value)

    @property
    def direction(self) -> XL_ERROR_BAR_DIRECTION:
        """Read/write. The direction of the error bars (X or Y).

        A member of :ref:`XlErrorBarDirection`.
        """
        errDir = self._element.errDir
        if errDir is None:
            return XL_ERROR_BAR_DIRECTION.Y
        return XL_ERROR_BAR_DIRECTION.from_xml(errDir.val)

    @direction.setter
    def direction(self, value: XL_ERROR_BAR_DIRECTION):
        self._element.get_or_add_errDir().val = XL_ERROR_BAR_DIRECTION.to_xml(value)

    @property
    def include(self) -> XL_ERROR_BAR_INCLUDE:
        """Read/write. Which error bars to display (both, plus, or minus).

        A member of :ref:`XlErrorBarInclude`.
        """
        errBarType = self._element.errBarType
        if errBarType is None:
            return XL_ERROR_BAR_INCLUDE.BOTH
        return XL_ERROR_BAR_INCLUDE.from_xml(errBarType.val)

    @include.setter
    def include(self, value: XL_ERROR_BAR_INCLUDE):
        self._element.get_or_add_errBarType().val = XL_ERROR_BAR_INCLUDE.to_xml(value)

    @property
    def value(self) -> float:
        """Read/write float. The fixed value, percentage, or number of standard deviations.

        Only meaningful when type is FIXED_VALUE, PERCENT, or ST_DEV.
        """
        val = self._element.val_val
        if val is None:
            return 0.0
        return val

    @value.setter
    def value(self, amount: float):
        self._element.val_val = amount

    @property
    def has_end_cap(self) -> bool:
        """Read/write boolean. True if the error bars have end caps (T-shaped ends)."""
        noEndCap = self._element.noEndCap
        if noEndCap is None:
            return True
        return not noEndCap.val

    @has_end_cap.setter
    def has_end_cap(self, value: bool):
        self._element.get_or_add_noEndCap().val = not value


class ErrorBarsCollection:
    """Provides access to error bars on a series.

    A series can have up to two error bar sets (X and Y directions on XY charts).
    For most chart types, only one set (Y direction) is applicable.
    """

    def __init__(self, ser: CT_SeriesComposite):
        self._ser = ser

    @property
    def has_error_bars(self) -> bool:
        """True if this series has at least one set of error bars."""
        return len(self._ser.errBars_lst) > 0

    def __len__(self) -> int:
        return len(self._ser.errBars_lst)

    def __iter__(self):
        return (ErrorBars(eb) for eb in self._ser.errBars_lst)

    def __getitem__(self, index) -> ErrorBars:
        return ErrorBars(self._ser.errBars_lst[index])

    def add(
        self,
        bar_type: XL_ERROR_BAR_TYPE = XL_ERROR_BAR_TYPE.FIXED_VALUE,
        value: float = 0.0,
    ) -> ErrorBars:
        """Add error bars to the series and return the new ErrorBars object.

        Args:
            bar_type: How the error amount is determined, defaults to FIXED_VALUE.
            value: The error amount.
        """
        errBars_elm = self._ser._add_errBars()
        errBarType_elm = OxmlElement("c:errBarType")
        errBarType_elm.set("val", "both")
        errBars_elm.append(errBarType_elm)
        errValType_elm = OxmlElement("c:errValType")
        errValType_elm.set("val", XL_ERROR_BAR_TYPE.to_xml(bar_type))
        errBars_elm.append(errValType_elm)
        if bar_type not in (XL_ERROR_BAR_TYPE.ST_ERROR, XL_ERROR_BAR_TYPE.CUSTOM):
            val_elm = OxmlElement("c:val")
            val_elm.set("val", str(value))
            errBars_elm.append(val_elm)
        return ErrorBars(errBars_elm)

    def remove(self, error_bars: ErrorBars) -> None:
        """Remove the specified error bars from the series."""
        self._ser.remove(error_bars._element)
