# 2026-09-24 18:44 · epic-79-shape-line-fill-table

**Goal:** Close epic #79 (sub-issues #80-#86)

## What happened
- Four parallel `python-developer` agents in git worktrees, grouped by file to limit conflicts: #80 freeform;
  #81+#86 line/autofit; #82+#83 fill.py; #84+#85 table.py. Merged into the branch; one PR (#96) closed all 7.
- #80 `FreeformBuilder.add_cubic_bezier/add_quadratic_bezier/add_arc`; ops expose `.points` for extents.
- #81 `LineFormat.cap_style/join_style/miter_limit` (`MSO_LINE_CAP_STYLE`, `MSO_LINE_JOIN_STYLE`; lim via ST_Percentage).
- #82 `FillFormat.picture()/tile()`; `from_fill_parent(..., part_provider)` — only `Shape.fill` passes one.
- #83 `FillFormat.gradient_type` (`MSO_GRADIENT_TYPE`) + `gradient_fill_to_rect`.
- #84 `rows/columns.add(index)` / `.remove(x)` in `CT_Table.insert_tr/remove_tr/insert_gridCol/remove_gridCol`.
- #85 `_Cell.text_direction`; #86 `TextFrame.font_scale/line_spacing_reduction` (read-only).

## Gotchas & dead ends
- Concurrent worktree agents each `pip install -e` → the editable install points at whichever ran last. Always test
  with `PYTHONPATH=src`. Worktrees also share one `git stash` — don't pop another agent's entry.
- `a:arcTo` stAng/swAng are *visual* angles (ray from centre), not parametric; convert with
  `t = atan2(wR·sinθ, hR·cosθ)` or extents drift on ellipses (fixed in 7256920).
- Table merges: fuzzing (random merge + add/remove, strict span/flag checker) found 2-D merge bugs the agent's tests
  missed — origin move dropped the orthogonal span; promotion must key on "next cell is a continuation", not on the
  removed cell's own span (bare merges carry spans on the origin only); clamp decrements at 1 (69b414d).
- `a:path` tag is shared by gradFill and custGeom (`CT_Path2D` registered); gradient code uses generic lxml calls.

## State at end
- PR #96 merged (46f6de2); #79–#86 closed. 3225 tests pass; all examples run.

## Open threads
- Human PowerPoint checks: picture/tiled fill, radial gradient, elliptical arc, table insert inside a merge.
- 4 new pyright errors in dml/fill.py (`_Fill` base access, same pre-existing pattern as gradient_angle).
