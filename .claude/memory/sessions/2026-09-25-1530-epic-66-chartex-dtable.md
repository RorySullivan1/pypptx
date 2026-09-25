# 2026-09-25 15:30 · epic-66-chartex-dtable

**Goal:** Implement epic #66 (sub-issues #67-#70: chartex read/add, chart data table) and open a PR

## What happened
- #70 (`c:dTable`) delegated to a `python-developer` worktree agent, merged in: `Chart.has_data_table`,
  `DataTable` (horizontal_border/vertical_border/outline/show_keys); new dTable = PowerPoint's plain default
  (borders + outline, no keys). It also gave `CT_PlotArea.catAx/valAx` successors so axes insert before dTable/spPr.
- #67 `ChartExPart` (LazyXmlPart) + `oxml/chart/chartex.py`; `cx`/`mc` namespaces (mc added BEFORE `ve` in
  `_nsmap` so `pfxmap` still maps that URI to `ve`); `RT.CHART_EX`; mc:* tags get `CT_MarkupCompatibilityElement`.
- Shape tree: `iter_shape_elms` yields the chartex `p:graphicFrame` inside `mc:AlternateContent/mc:Choice`
  (only when its graphicData uri is chartex). `oxml.shapes.shared.tree_elm(elm)` = the spTree child to move/
  remove/duplicate/group; used in shapetree + `BaseShape.is_in_group/parent_group`.
- #68 `ChartEx.chart_type` (paretoLine ⇒ PARETO) and series name/values/categories/category_paths.
- #69 `XL_CHART_TYPE` WATERFALL 119, HISTOGRAM 118, PARETO 122, BOX_WHISKER 121, TREEMAP 117, SUNBURST 120,
  FUNNEL 123, REGION_MAP 140; `add_chart()` routes the six addable types via `chart/chartexwriter.py`.
- `BaseShape.has_chartex` / `has_smartart` default False (has_smartart was missing since #62).
- `examples/modern_charts.py` writes every new chart type — the file for the human PowerPoint check.

## Gotchas & dead ends
- `tests/opc/test_package.py` PartFactory test leaked a mock SlidePart into the global
  `PartFactory.part_type_for`; any later in-process test loading a file got mock slides. Fixed with monkeypatch.
- lxml default elements (mc:*) have no namespace-aware `xpath()` → register a BaseOxmlElement subclass.
- Chartex hierarchy: cx:lvl order is leaf first; workbook repeats parent labels per row; strDim `cx:f dir="row"`
  for multi-level. Treemap/sunburst values are `numDim type="size"`. Funnel needs Requires cx2 (2015/10/21),
  the rest cx1 (2015/9/8). Leaf-first order and dir="row" are from memory of Office output — human check.

## State at end
- 3562 tests pass; all examples run; new modules pyright-clean; ruff count unchanged vs main.

## Open threads
- Human PowerPoint checks: each added chartex type opens without repair and looks right (examples/modern_charts.py);
  data table shows under a column chart.
