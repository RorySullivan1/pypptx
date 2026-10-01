"""Deep read-only walk of a presentation, reporting every read that raises.

Used by the real-world corpus tests: a file saved by PowerPoint should be readable through the
public API without any property access raising. The walk visits the slide masters, their
layouts, the slides and their notes, recursing into group shapes, and reads each shape's
geometry, type, placeholder format, text (paragraphs, runs, basic font properties), table
cells, chart type/categories/series values and fill type.
"""

from __future__ import annotations

from typing import Callable, Iterable, List, Tuple

from pptx.presentation import Presentation

Failure = Tuple[str, str]  # (where, "ExceptionType: message")


def walk(prs: Presentation) -> List[Failure]:
    """Return one `(where, error)` pair for every read in `prs` that raised; empty when clean."""
    failures: List[Failure] = []

    def attempt(where: str, read: Callable[[], object]) -> None:
        try:
            read()
        except Exception as e:  # noqa: BLE001 -- the point is to report any exception
            failures.append((where, "%s: %s" % (type(e).__name__, e)))

    for m_idx, master in enumerate(prs.slide_masters):
        _walk_shapes(master.shapes, "master[%d]" % m_idx, attempt)
        for l_idx, layout in enumerate(master.slide_layouts):
            _walk_shapes(layout.shapes, "master[%d]/layout[%d]" % (m_idx, l_idx), attempt)

    for s_idx, slide in enumerate(prs.slides):
        where = "slide[%d]" % s_idx
        _walk_shapes(slide.shapes, where, attempt)
        attempt(where + "/notes", lambda slide=slide: _read_notes(slide))

    return failures


def _read_notes(slide) -> None:
    if slide.has_notes_slide:
        text_frame = slide.notes_slide.notes_text_frame
        if text_frame is not None:
            text_frame.text


def _walk_shapes(shapes: Iterable, where: str, attempt) -> None:
    for idx, shape in enumerate(shapes):
        here = "%s/shape[%d]" % (where, idx)
        for name in ("shape_id", "name", "shape_type", "left", "top", "width", "height"):
            attempt("%s.%s" % (here, name), lambda shape=shape, name=name: getattr(shape, name))
        attempt(here + ".rotation", lambda shape=shape: getattr(shape, "rotation", None))
        attempt(here + ".placeholder_format", lambda shape=shape: _read_placeholder(shape))
        attempt(here + ".fill", lambda shape=shape: _read_fill(shape))
        attempt(here + ".text_frame", lambda shape=shape: _read_text(shape))
        attempt(here + ".table", lambda shape=shape: _read_table(shape))
        attempt(here + ".chart", lambda shape=shape: _read_chart(shape))
        if hasattr(shape, "shapes"):
            _walk_shapes(shape.shapes, here, attempt)


def _read_placeholder(shape) -> None:
    if shape.is_placeholder:
        placeholder_format = shape.placeholder_format
        placeholder_format.idx
        placeholder_format.type


def _read_fill(shape) -> None:
    fill = getattr(shape, "fill", None)
    if fill is not None:
        fill.type


def _read_text(shape) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    for paragraph in shape.text_frame.paragraphs:
        paragraph.level
        paragraph.alignment
        for run in paragraph.runs:
            run.text
            font = run.font
            font.size
            font.bold
            font.italic
            font.name
            font.color.type


def _read_table(shape) -> None:
    if not getattr(shape, "has_table", False):
        return
    for row in shape.table.rows:
        row.height
        for cell in row.cells:
            cell.text
            cell.is_merge_origin
            cell.is_spanned


def _read_chart(shape) -> None:
    if not getattr(shape, "has_chart", False):
        return
    chart = shape.chart
    chart.chart_type
    chart.has_legend
    for plot in chart.plots:
        list(plot.categories)
        for series in plot.series:
            series.name
            list(series.values)
