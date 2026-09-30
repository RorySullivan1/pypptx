"""DrawingML objects related to line formatting."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.dml.fill import FillFormat
from pptx.enum.dml import MSO_FILL, MSO_LINE_JOIN_STYLE
from pptx.exc import InvalidValueError
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.dml.color import ColorFormat
    from pptx.enum.dml import (
        MSO_LINE_CAP_STYLE,
        MSO_LINE_COMPOUND_TYPE,
        MSO_LINE_DASH_STYLE,
        MSO_LINE_END_SIZE,
        MSO_LINE_END_TYPE,
    )
    from pptx.oxml.shapes.shared import CT_LineProperties
    from pptx.util import Length


class LineFormat:
    """Provides access to line properties such as color, style, and width.

    A LineFormat object is typically accessed via the ``.line`` property of
    a shape such as |Shape| or |Picture|.
    """

    def __init__(self, parent: object) -> None:
        super(LineFormat, self).__init__()
        self._parent = parent

    @property
    def begin_arrowhead_length(self) -> MSO_LINE_END_SIZE | None:
        """Size of the arrowhead at the beginning of the line.

        Read/write. A member of :ref:`MsoLineEndSize` or |None|.
        """
        ln = self._ln
        if ln is None or ln.headEnd is None:
            return None
        return ln.headEnd.len

    @begin_arrowhead_length.setter
    def begin_arrowhead_length(self, value: MSO_LINE_END_SIZE | None) -> None:
        if value is None:
            ln = self._ln
            if ln is not None and ln.headEnd is not None:
                ln.headEnd.len = None
            return
        headEnd = self._get_or_add_ln().get_or_add_headEnd()
        headEnd.len = value

    @property
    def begin_arrowhead_type(self) -> MSO_LINE_END_TYPE | None:
        """Type of arrowhead at the beginning of the line.

        Read/write. A member of :ref:`MsoLineEndType` or |None|.
        """
        ln = self._ln
        if ln is None or ln.headEnd is None:
            return None
        return ln.headEnd.type

    @begin_arrowhead_type.setter
    def begin_arrowhead_type(self, value: MSO_LINE_END_TYPE | None) -> None:
        if value is None:
            ln = self._ln
            if ln is not None and ln.headEnd is not None:
                ln.headEnd.type = None
            return
        headEnd = self._get_or_add_ln().get_or_add_headEnd()
        headEnd.type = value

    @property
    def begin_arrowhead_width(self) -> MSO_LINE_END_SIZE | None:
        """Width of the arrowhead at the beginning of the line.

        Read/write. A member of :ref:`MsoLineEndSize` or |None|.
        """
        ln = self._ln
        if ln is None or ln.headEnd is None:
            return None
        return ln.headEnd.w

    @begin_arrowhead_width.setter
    def begin_arrowhead_width(self, value: MSO_LINE_END_SIZE | None) -> None:
        if value is None:
            ln = self._ln
            if ln is not None and ln.headEnd is not None:
                ln.headEnd.w = None
            return
        headEnd = self._get_or_add_ln().get_or_add_headEnd()
        headEnd.w = value

    @property
    def cap_style(self) -> MSO_LINE_CAP_STYLE | None:
        """Style of the line's end caps (flat, round, or square).

        Read/write. A member of :ref:`MsoLineCapStyle` or |None|. |None| indicates
        no explicit setting is present and the effective value is inherited from
        the style hierarchy. Assigning |None| removes any existing explicit
        setting. Reflects the `a:ln/@cap` attribute.
        """
        ln = self._ln
        if ln is None:
            return None
        return ln.cap

    @cap_style.setter
    def cap_style(self, value: MSO_LINE_CAP_STYLE | None) -> None:
        if value is None:
            ln = self._ln
            if ln is not None:
                ln.cap = None
            return
        ln = self._get_or_add_ln()
        ln.cap = value

    @lazyproperty
    def color(self) -> ColorFormat:
        """
        The |ColorFormat| instance that provides access to the color settings
        for this line. Essentially a shortcut for ``line.fill.fore_color``.
        As a side-effect, accessing this property causes the line fill type
        to be set to ``MSO_FILL.SOLID``. If this sounds risky for your use
        case, use ``line.fill.type`` to non-destructively discover the
        existing fill type.
        """
        if self.fill.type != MSO_FILL.SOLID:
            self.fill.solid()
        return self.fill.fore_color

    @property
    def compound_type(self) -> MSO_LINE_COMPOUND_TYPE | None:
        """Compound line type (single, double, thick-thin, etc.).

        Read/write. A member of :ref:`MsoLineCompoundType` or |None|.
        |None| indicates the default (single line).
        """
        ln = self._ln
        if ln is None:
            return None
        return ln.cmpd

    @compound_type.setter
    def compound_type(self, value: MSO_LINE_COMPOUND_TYPE | None) -> None:
        ln = self._get_or_add_ln()
        ln.cmpd = value

    @property
    def dash_style(self) -> MSO_LINE_DASH_STYLE | None:
        """Return value indicating line style.

        Returns a member of :ref:`MsoLineDashStyle` indicating line style, or
        |None| if no explicit value has been set. When no explicit value has
        been set, the line dash style is inherited from the style hierarchy.

        Assigning |None| removes any existing explicitly-defined dash style.
        """
        ln = self._ln
        if ln is None:
            return None
        return ln.prstDash_val

    @dash_style.setter
    def dash_style(self, dash_style: MSO_LINE_DASH_STYLE | None) -> None:
        if dash_style is None:
            ln = self._ln
            if ln is None:
                return
            ln._remove_prstDash()
            ln._remove_custDash()
            return
        ln = self._get_or_add_ln()
        ln.prstDash_val = dash_style

    @property
    def end_arrowhead_length(self) -> MSO_LINE_END_SIZE | None:
        """Size of the arrowhead at the end of the line.

        Read/write. A member of :ref:`MsoLineEndSize` or |None|.
        """
        ln = self._ln
        if ln is None or ln.tailEnd is None:
            return None
        return ln.tailEnd.len

    @end_arrowhead_length.setter
    def end_arrowhead_length(self, value: MSO_LINE_END_SIZE | None) -> None:
        if value is None:
            ln = self._ln
            if ln is not None and ln.tailEnd is not None:
                ln.tailEnd.len = None
            return
        tailEnd = self._get_or_add_ln().get_or_add_tailEnd()
        tailEnd.len = value

    @property
    def end_arrowhead_type(self) -> MSO_LINE_END_TYPE | None:
        """Type of arrowhead at the end of the line.

        Read/write. A member of :ref:`MsoLineEndType` or |None|.
        """
        ln = self._ln
        if ln is None or ln.tailEnd is None:
            return None
        return ln.tailEnd.type

    @end_arrowhead_type.setter
    def end_arrowhead_type(self, value: MSO_LINE_END_TYPE | None) -> None:
        if value is None:
            ln = self._ln
            if ln is not None and ln.tailEnd is not None:
                ln.tailEnd.type = None
            return
        tailEnd = self._get_or_add_ln().get_or_add_tailEnd()
        tailEnd.type = value

    @property
    def end_arrowhead_width(self) -> MSO_LINE_END_SIZE | None:
        """Width of the arrowhead at the end of the line.

        Read/write. A member of :ref:`MsoLineEndSize` or |None|.
        """
        ln = self._ln
        if ln is None or ln.tailEnd is None:
            return None
        return ln.tailEnd.w

    @end_arrowhead_width.setter
    def end_arrowhead_width(self, value: MSO_LINE_END_SIZE | None) -> None:
        if value is None:
            ln = self._ln
            if ln is not None and ln.tailEnd is not None:
                ln.tailEnd.w = None
            return
        tailEnd = self._get_or_add_ln().get_or_add_tailEnd()
        tailEnd.w = value

    @lazyproperty
    def fill(self) -> FillFormat:
        """
        |FillFormat| instance for this line, providing access to fill
        properties such as foreground color.
        """
        ln = self._get_or_add_ln()
        return FillFormat.from_fill_parent(ln)

    @property
    def join_style(self) -> MSO_LINE_JOIN_STYLE | None:
        """Style used to join two line segments (round, bevel, or miter).

        Read/write. A member of :ref:`MsoLineJoinStyle` or |None|. |None|
        indicates no explicit join is present, in which case the effective
        value is inherited from the style hierarchy. Assigning |None| removes
        any existing explicit join (`a:round`, `a:bevel`, or `a:miter`).

        Use :attr:`miter_limit` to set the miter-limit ratio when this is
        `MSO_LINE_JOIN_STYLE.MITER`.
        """
        ln = self._ln
        if ln is None:
            return None
        if ln.round is not None:
            return MSO_LINE_JOIN_STYLE.ROUND
        if ln.bevel is not None:
            return MSO_LINE_JOIN_STYLE.BEVEL
        if ln.miter is not None:
            return MSO_LINE_JOIN_STYLE.MITER
        return None

    @join_style.setter
    def join_style(self, value: MSO_LINE_JOIN_STYLE | None) -> None:
        if value is None:
            ln = self._ln
            if ln is not None:
                ln._remove_eg_lineJoinProperties()
            return
        ln = self._get_or_add_ln()
        if value == MSO_LINE_JOIN_STYLE.ROUND:
            ln.get_or_change_to_round()
        elif value == MSO_LINE_JOIN_STYLE.BEVEL:
            ln.get_or_change_to_bevel()
        elif value == MSO_LINE_JOIN_STYLE.MITER:
            ln.get_or_change_to_miter()
        else:
            raise InvalidValueError(f"only a member of MSO_LINE_JOIN_STYLE or None can be assigned, got {value!r}")

    @property
    def miter_limit(self) -> float | None:
        """Miter-limit ratio for a mitered line join, as a float.

        Read/write. Applicable only when :attr:`join_style` is
        `MSO_LINE_JOIN_STYLE.MITER`. Reflects the `a:miter/@lim` attribute,
        which stores the value in 1000ths of a percent (e.g. `400000` for a
        4.0 ratio); this property exposes it as a plain float, e.g. `4.0`.
        Returns |None| when no `a:miter` element is present, or when it is
        present but `lim` is not explicitly set (in which case the OOXML
        default miter-limit ratio applies).

        Setting this value has no effect unless :attr:`join_style` is (or is
        also set to) `MSO_LINE_JOIN_STYLE.MITER`; assigning a value implicitly
        changes the join to miter. Assigning |None| removes the `lim`
        attribute but leaves the miter join in place.
        """
        ln = self._ln
        if ln is None or ln.miter is None:
            return None
        return ln.miter.lim

    @miter_limit.setter
    def miter_limit(self, value: float | None) -> None:
        if value is None:
            ln = self._ln
            if ln is not None and ln.miter is not None:
                ln.miter.lim = None
            return
        miter = self._get_or_add_ln().get_or_change_to_miter()
        miter.lim = value

    @property
    def no_fill(self) -> bool:
        """Whether the line has no fill (invisible).

        Read-only. |True| when the line has a ``noFill`` child element.
        Use ``line.fill.background()`` to make a line invisible.
        """
        ln = self._ln
        if ln is None:
            return False
        return ln.noFill is not None

    @property
    def width(self) -> Length:
        """
        The width of the line expressed as an integer number of :ref:`English
        Metric Units <EMU>`. The returned value is an instance of |Length|,
        a value class having properties such as `.inches`, `.cm`, and `.pt`
        for converting the value into convenient units.
        """
        ln = self._ln
        if ln is None:
            return Emu(0)
        return ln.w

    @width.setter
    def width(self, emu: Length | None) -> None:
        if emu is None:
            emu = 0
        ln = self._get_or_add_ln()
        ln.w = emu

    def _get_or_add_ln(self) -> CT_LineProperties:
        """
        Return the ``<a:ln>`` element containing the line format properties
        in the XML.
        """
        return self._parent.get_or_add_ln()

    @property
    def _ln(self) -> CT_LineProperties | None:
        return self._parent.ln
