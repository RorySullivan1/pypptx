"""DrawingML objects related to line formatting."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.dml.fill import FillFormat
from pptx.enum.dml import MSO_FILL
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.enum.dml import MSO_LINE_COMPOUND_TYPE, MSO_LINE_END_SIZE, MSO_LINE_END_TYPE


class LineFormat:
    """Provides access to line properties such as color, style, and width.

    A LineFormat object is typically accessed via the ``.line`` property of
    a shape such as |Shape| or |Picture|.
    """

    def __init__(self, parent):
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
    def begin_arrowhead_length(self, value: MSO_LINE_END_SIZE | None):
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
    def begin_arrowhead_type(self, value: MSO_LINE_END_TYPE | None):
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
    def begin_arrowhead_width(self, value: MSO_LINE_END_SIZE | None):
        if value is None:
            ln = self._ln
            if ln is not None and ln.headEnd is not None:
                ln.headEnd.w = None
            return
        headEnd = self._get_or_add_ln().get_or_add_headEnd()
        headEnd.w = value

    @lazyproperty
    def color(self):
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
    def compound_type(self, value: MSO_LINE_COMPOUND_TYPE | None):
        ln = self._get_or_add_ln()
        ln.cmpd = value

    @property
    def dash_style(self):
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
    def dash_style(self, dash_style):
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
    def end_arrowhead_length(self, value: MSO_LINE_END_SIZE | None):
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
    def end_arrowhead_type(self, value: MSO_LINE_END_TYPE | None):
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
    def end_arrowhead_width(self, value: MSO_LINE_END_SIZE | None):
        if value is None:
            ln = self._ln
            if ln is not None and ln.tailEnd is not None:
                ln.tailEnd.w = None
            return
        tailEnd = self._get_or_add_ln().get_or_add_tailEnd()
        tailEnd.w = value

    @lazyproperty
    def fill(self):
        """
        |FillFormat| instance for this line, providing access to fill
        properties such as foreground color.
        """
        ln = self._get_or_add_ln()
        return FillFormat.from_fill_parent(ln)

    @property
    def width(self):
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
    def width(self, emu):
        if emu is None:
            emu = 0
        ln = self._get_or_add_ln()
        ln.w = emu

    def _get_or_add_ln(self):
        """
        Return the ``<a:ln>`` element containing the line format properties
        in the XML.
        """
        return self._parent.get_or_add_ln()

    @property
    def _ln(self):
        return self._parent.ln
