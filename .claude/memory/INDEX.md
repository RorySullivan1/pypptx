# MEMORY INDEX  ·  keep ≤ ~80 lines, ≤ ~200 chars per line

## State            (rewrite in place — current truth only, ≤ ~10 lines)
- pypptx: pure-Python `.pptx` library (dist `pypptx`, import `pptx`); architecture + conventions in `CLAUDE.md`.
- All Claude tooling lives in `.claude/` (inventory: `.claude/CATALOG.md`); reference docs in `.claude/context/`.
- Roadmap in `.meta/roadmap/` (stages 01–04, v0.1.0 → v1.0.0); cursor `.meta/version` = v0.2.0 shipped; next v0.2.1.
- Epic #87 closed (PR #94): master text styles, theme style lists, `SlideMaster.apply_theme`.
- Epic #79 closed (PR #96): freeform curves/arcs, line cap/join, picture fill, non-linear gradients, table row/col add/remove, cell text direction, autofit readouts.
- Assets adopted from claudeBrain @ `8e281af` by selection; shared ones are copies — improve upstream, re-copy.

## Decisions        (append-only; supersede, never delete)
- [2026-09-24] Claude assets consolidated: internal_docs → `.claude/context/`, dev_map → `.meta/roadmap/` (hooks read `.meta/`)
  — sessions/2026-09-24-1505-adopt-claudebrain-assets.md
- [2026-09-24] Adopted factory core + roadmap tier + Python/GitHub/docs families. EXCLUDED (not our stacks): VSTO, VBA,
  Power Platform/SharePoint/Power BI, quant, branding→presentation, Outlook/print HTML. Prose tier deferred to v0.2.2 — same log
- [2026-09-24] Keep pypptx's branch-per-version naming `vX.Y.Z` over `/version-set`'s `claude/<label>-<slug>` default (CLAUDE.md § Roadmap) — same log
- [2026-09-24] Text-style levels are zero-based like `_Paragraph.level`; `phClr` → `MSO_THEME_COLOR.PLACEHOLDER` (17)
  — sessions/2026-09-24-1819-epic-87-text-styles-themes.md
- [2026-09-24] Arc angles are visual (DrawingML), converted to parametric for extents; table merge edits promote the next
  continuation cell and copy the orthogonal span — sessions/2026-09-24-1844-epic-79-shape-line-fill-table.md

## Threads          (open items; remove when closed)
- Epic #87 human checks: master text styles rendering in PowerPoint; `apply_theme` with a real PowerPoint .thmx.
- Epic #79 human checks: picture/tiled fill, radial gradient, elliptical arc, table insert inside a merge.
- v0.2.2 docstring audit: consider adopting the factory's prose tier (`prose_budget.py`, `prose-auditor`, `/prose-review`).
- `unverified` surfaces (verification-surface.md): type checker, Sphinx build, real-world corpus.

## Log              (append-only pointers)
- 2026-09-24 1505 | Consolidate into .claude/ + adopt claudeBrain assets | sessions/2026-09-24-1505-adopt-claudebrain-assets.md
- 2026-09-24 1819 | Close epic #87 (text styles, theme lists, apply_theme) | sessions/2026-09-24-1819-epic-87-text-styles-themes.md
- 2026-09-24 1844 | Close epic #79 (shape/line/fill/table completeness) | sessions/2026-09-24-1844-epic-79-shape-line-fill-table.md
