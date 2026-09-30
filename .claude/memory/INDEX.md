# MEMORY INDEX  ·  keep ≤ ~80 lines, ≤ ~200 chars per line

## State            (rewrite in place — current truth only, ≤ ~10 lines)
- pypptx: pure-Python `.pptx` library (dist `pypptx`, import `pptx`); architecture + conventions in `CLAUDE.md`.
- All Claude tooling lives in `.claude/` (inventory: `.claude/CATALOG.md`); reference docs in `.claude/context/`.
- Package 0.3.2; v0.3.0–v0.3.2 shipped. 0.3.2 also carries #114 (shape tags) and #115 (3D charts).
- Cursor v0.4.0 (real-world round-trip; card still an overview — expand before building); then v0.5.0 … v0.7.0.
- API reference: `sphinx-build -W -b html docs docs/_build/html` (docs extra). Only #51 open on GitHub.
- No git tags or Releases yet: sessions can't push tags. Owner tags by hand from `.meta/release-tags.md` (v0.1.0–v0.3.2).
- Sessions push only to their assigned branch, not `vX.Y.Z`; `.meta/version` records the real branch.
- **Actions is billing-blocked (since 2026-09-27): all evidence is local.** LF pinned; `tests.yml` covers 3.9–3.13/Win/macOS.

## Decisions        (append-only; supersede, never delete)
- [2026-09-24..25] 10 earlier decisions (claudeBrain adoption, text styles, arcs, element-class keying, SmartArt,
  chartex, AlternateContent, hot-path caches) — sessions/ARCHIVE-2026.md
- [2026-09-26] Roadmap renumbered: shipped epics = v0.3.0; docs versions v0.2.1/v0.2.2 → v0.3.1/v0.3.2; stage-03 plans
  v0.3.0–v0.6.0 → v0.4.0–v0.7.0 (package versions must increase) — .meta/roadmap/stages/03-hardening-and-gaps/v0.3.0.md
- [2026-09-26] Pin `* text=auto eol=lf` repo-wide + explicit `binary` fixtures: the snippet loaders are
  byte-exact (`"rb"`, split on `"

"`) so a CRLF tree breaks 51 tests; `osx_dirs_fixture` injects `posixpath`
  rather than skipping on Windows — sessions/2026-09-26-2219-windows-crlf-acceptance.md
- [2026-09-29] Tags are one part PER owner, referenced from its `p:custDataLst/p:tags` (`p:nvPr` for a shape,
  `p:cSld` for the slide), related from the slide part; slide tags resolve by that reference, not by reltype
  — sessions/2026-09-29-1211-issue-114-shape-tags.md
- [2026-09-29] 3D chart writers share `_Base3DChartXmlWriter`; only THREE_D_COLUMN/AREA/LINE get a `c:serAx` + perspective
  camera. New chart XML uses POSITIVE axis ids (`unsignedInt`) — sessions/2026-09-29-2205-issue-115-3d-charts.md
- [2026-09-30] Sphinx `|Name|` substitutions are GENERATED in docs/conf.py from pptx's classes (no hand list); enum `:ref:`
  labels live on docs/api/enum.rst; `-W` is the docs gate — sessions/2026-09-30-0353-v032-docstrings-sphinx.md

## Threads          (open items; remove when closed)
- Epic #87 human checks: master text styles rendering in PowerPoint; `apply_theme` with a real PowerPoint .thmx.
- Epic #79 human checks: picture/tiled fill, radial gradient, elliptical arc, table insert inside a merge.
- Epic #56 human checks: custom show plays the right slides; loop/kiosk take effect; last-embedded-font removal opens cleanly.
- v0.3.2 docstring audit: consider adopting the factory's prose tier (`prose_budget.py`, `prose-auditor`, `/prose-review`).
- Epic #62 human check: edited SmartArt node text shows in PowerPoint (use a PowerPoint-authored deck).
- Epic #66 human checks: each added chartex type opens without repair (examples/modern_charts.py); data table renders.
- Epic #75 human checks: added audio plays; trim/fade take effect; 3D model/zoom deck matches the selection pane.
- Issue #51 (checklist trimmed 2026-09-25): open items are proxy-tuple caching (paragraphs/runs/shapes) and a "no slower than
  baseline" `pytest -m perf` gate in CI; plan + baseline in `.claude/context/performance.md`.
- `unverified` surfaces (verification-surface.md): type checker, real-world corpus.
- Anyone with a Windows clone predating `b64d2b9` needs `git add --renormalize .` once after pulling, or git
  reports every text file modified (its dirty check compares the size recorded in the index, not content).
- `tests.yml` caveats: `paths-ignore` (md/.claude/.meta) must go if `tests` ever becomes a REQUIRED check, or
  docs-only PRs wait forever; ~17 billable min/run, 10 of them the single macOS cell (x10 multiplier).
- The #51 perf gate is still absent, deliberately — shared-runner timing noise would make it flake.
- **Actions billing blocks every job.** Fix the spending limit/payment, or cut cost: the lone macOS cell is
  ~10 of ~17 billable min per run. Until then no PR can be CI-verified.
- #115 human check: new 3D charts (column/bar/line/pie/area) open without repair; camera/walls visibly change. No `series_axis` API.
- #114 human check: shape-tagged deck opens in PowerPoint without repair. Follow-up idea: `has_tags` (reading creates a part).
- v0.4.0 card expanded (pilot over 85 Apache POI decks). Owner decisions before building: OK to commit ALv2 POI test
  files; `shape_type` of a geometry-less `p:sp` → AUTO_SHAPE (recommended) vs None.

## Log              (append-only pointers)
- 2026-09-24 1505 | Consolidate into .claude/ + adopt claudeBrain assets | sessions/2026-09-24-1505-adopt-claudebrain-assets.md
- 2026-09-24 1819 | Close epic #87 (text styles, theme lists, apply_theme) | sessions/2026-09-24-1819-epic-87-text-styles-themes.md
- 2026-09-24 1844 | Close epic #79 (shape/line/fill/table completeness) | sessions/2026-09-24-1844-epic-79-shape-line-fill-table.md
- 2026-09-25 1202 | Close epic #56 (package-level presentation parts) | sessions/2026-09-25-1202-epic-56-presentation-parts.md
- 2026-09-25 1400 | Epic #62 SmartArt read + node text edit (PR #101 merged) | sessions/2026-09-25-1400-epic-62-smartart.md
- 2026-09-25 1530 | Epic #66 chartex + data table (PR #103 merged) | sessions/2026-09-25-1530-epic-66-chartex-dtable.md
- 2026-09-25 1600 | Epic #75 audio, trim/fade, AlternateContent shapes (PR #105 merged) | sessions/2026-09-25-1600-epic-75-media-altcontent.md
- 2026-09-25 1630 | Issue #51: qn/XPath caches (−39%/deck); perf harness PR #109 | sessions/2026-09-25-1630-issue-51-perf-caches.md
- 2026-09-26 | Roadmap reconcile + package 0.3.0 | .meta/roadmap/INDEX.md
- 2026-09-26 2219 | Windows/LF acceptance + cross-platform CI (PR #112 merged b64d2b9) | sessions/2026-09-26-2219-windows-crlf-acceptance.md
- 2026-09-27 1829 | v0.3.1 examples (PR #116); #114/#115 filed; Actions billing-blocked | sessions/2026-09-27-1829-v031-examples.md
- 2026-09-28 | v0.3.1 marked shipped, cursor → v0.3.2, package 0.3.1 | .meta/version
- 2026-09-29 1211 | #114 shape tags (`BaseShape.tags`), slide tags keyed by cSld ref | sessions/2026-09-29-1211-issue-114-shape-tags.md
- 2026-09-29 2205 | #115 3D chart authoring: 13 THREE_D_* types via add_chart | sessions/2026-09-29-2205-issue-115-3d-charts.md
- 2026-09-30 0353 | v0.3.2 started: Sphinx scaffold, -W clean, docstring gaps closed | sessions/2026-09-30-0353-v032-docstrings-sphinx.md
- 2026-09-30 | v0.3.2 shipped (PR #120, 33a3e02): package 0.3.2, cursor → v0.4.0 | .meta/version
- 2026-09-30 | Release-tag commits recorded (.meta/release-tags.md); v0.4.0 card expanded from a POI-corpus pilot | .meta/roadmap/stages/03-hardening-and-gaps/v0.4.0.md
