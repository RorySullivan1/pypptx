"""Read access to Office 2016+ "chartex" charts: waterfall, histogram, treemap and friends.

A |ChartEx| is reached from a graphic frame, `graphic_frame.chartex`, when
`graphic_frame.has_chartex` is |True|. It exposes the chart type and each series' name,
categories and values as cached in the chart part. Editing chartex data is not supported; new
chartex charts are added with `shapes.add_chart()` like classic charts.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.chart import XL_CHART_TYPE

if TYPE_CHECKING:
    from pptx.oxml.chart.chartex import CT_ChartExDimension, CT_ChartExSeries, CT_ChartExSpace
    from pptx.parts.chartex import ChartExPart

# -- series `layoutId` -> chart type; a Pareto chart is a histogram ("clusteredColumn") series
# -- plus a "paretoLine" series, so it is recognized separately --
_LAYOUT_CHART_TYPES = {
    "boxWhisker": XL_CHART_TYPE.BOX_WHISKER,
    "clusteredColumn": XL_CHART_TYPE.HISTOGRAM,
    "funnel": XL_CHART_TYPE.FUNNEL,
    "paretoLine": XL_CHART_TYPE.PARETO,
    "regionMap": XL_CHART_TYPE.REGION_MAP,
    "sunburst": XL_CHART_TYPE.SUNBURST,
    "treemap": XL_CHART_TYPE.TREEMAP,
    "waterfall": XL_CHART_TYPE.WATERFALL,
}


class ChartEx:
    """A chartex chart (waterfall, histogram, Pareto, box & whisker, treemap, sunburst, funnel,
    region map), read-only.

    Not intended to be constructed directly; use `GraphicFrame.chartex`.
    """

    def __init__(self, chartSpace: CT_ChartExSpace, chart_part: ChartExPart):
        self._chartSpace = chartSpace
        self._chart_part = chart_part

    @property
    def chart_type(self) -> XL_CHART_TYPE | None:
        """Member of `XL_CHART_TYPE` for this chart, e.g. `XL_CHART_TYPE.WATERFALL`.

        |None| when the chart has no series or its series layout is not one pypptx recognizes.
        """
        layout_ids = [series.layoutId for series in self._chartSpace.series_lst]
        if "paretoLine" in layout_ids:
            return XL_CHART_TYPE.PARETO
        if not layout_ids:
            return None
        return _LAYOUT_CHART_TYPES.get(layout_ids[0])

    @property
    def part(self) -> ChartExPart:
        """The |ChartExPart| holding this chart."""
        return self._chart_part

    @property
    def series(self) -> tuple[ChartExSeries, ...]:
        """The series of this chart, in order.

        A Pareto chart's cumulative-percentage line is a series of its own, with `layout_id`
        "paretoLine" and no data of its own.
        """
        return tuple(ChartExSeries(ser, self._chartSpace) for ser in self._chartSpace.series_lst)


class ChartExSeries:
    """One series of a |ChartEx| chart, read-only."""

    def __init__(self, ser: CT_ChartExSeries, chartSpace: CT_ChartExSpace):
        self._ser = ser
        self._chartSpace = chartSpace

    @property
    def categories(self) -> tuple[str, ...]:
        """The leaf-level category label of each point, in order; empty when there are none.

        For hierarchical categories (treemap, sunburst) use `category_paths` for the full path.
        """
        return tuple(path[-1] for path in self.category_paths)

    @property
    def category_paths(self) -> tuple[tuple[str, ...], ...]:
        """The category of each point as a path from the top level down to the leaf.

        A single-level category gives a 1-tuple like `("Q1",)`. A parent label missing from the
        cache at a point (a blank cell) is taken from the nearest point above it.
        """
        dim = self._cat_dim
        if dim is None or not dim.lvl_lst:
            return ()
        # -- first cached level is the leaf level; reverse so paths run top-down --
        levels = [lvl.texts() for lvl in reversed(dim.lvl_lst)]
        point_count = max(len(texts) for texts in levels)
        paths: list[tuple[str, ...]] = []
        last: list[str] = [""] * len(levels)
        for idx in range(point_count):
            path: list[str] = []
            for depth, texts in enumerate(levels):
                text = texts[idx] if idx < len(texts) else None
                if text is None:
                    text = last[depth] if depth < len(levels) - 1 else ""
                last[depth] = text
                path.append(text)
            paths.append(tuple(path))
        return tuple(paths)

    @property
    def hidden(self) -> bool:
        """|True| when this series is hidden."""
        return self._ser.hidden

    @property
    def layout_id(self) -> str:
        """The series layout, e.g. "waterfall", "clusteredColumn" (histogram) or "treemap"."""
        return self._ser.layoutId

    @property
    def name(self) -> str | None:
        """Cached name of this series, or |None| if it has none."""
        return self._ser.name

    @property
    def values(self) -> tuple[float | None, ...]:
        """The cached value of each point, in order; |None| where a point has no value."""
        dim = self._val_dim
        if dim is None or not dim.lvl_lst:
            return ()
        return tuple(
            None if text is None or text == "" else float(text)
            for text in dim.lvl_lst[0].texts()
        )

    @property
    def _cat_dim(self) -> CT_ChartExDimension | None:
        data = self._data
        return None if data is None else data.cat_dim

    @property
    def _data(self):
        data_id = self._ser.data_id
        return None if data_id is None else self._chartSpace.data_for(data_id)

    @property
    def _val_dim(self) -> CT_ChartExDimension | None:
        data = self._data
        return None if data is None else data.val_dim
