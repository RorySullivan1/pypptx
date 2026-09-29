# MEMORY INDEX  ·  keep ≤ ~80 lines, ≤ ~200 chars per line

## State            (rewrite in place — current truth only, ≤ ~10 lines)
- pypptx: pure-Python `.pptx` library (dist `pypptx`, import `pptx`); architecture + conventions in `CLAUDE.md`.
- All Claude tooling lives in `.claude/` (inventory: `.claude/CATALOG.md`); reference docs in `.claude/context/`.
- Package 0.3.1. Roadmap `.meta/roadmap/`: v0.3.0 (epics #87/#79/#75/#66/#62/#56 + #51, card in stage 03) and
  v0.3.1 (examples, #116; 3D-chart example deferred to #115) shipped; cursor v0.3.2 (docstrings/Sphinx), then v0.4.0 … v0.7.0.
- No git tags or GitHub Releases yet: sessions can't push tags — the owner tags `v0.3.0`/`v0.3.1` by hand.
- Assets adopted from claudeBrain @ `8e281af` by selection; shared ones are copies — improve upstream, re-copy.
- Since `b64d2b9`: `.gitattributes` pins LF (CRLF tree = 52 failures), `pyparsing` is a dev dep, `tests.yml` runs
  pytest on Linux 3.9–3.13 + Win + macOS — **but Actions is billing-blocked (2026-09-27), so evidence is local**.

## Decisions        (append-only; supersede, never delete)
- [2026-09-24] Claude assets consolidated: internal_docs → `.claude/context/`, dev_map → `.meta/roadmap/` (hooks read `.meta/`)
  — sessions/2026-09-24-1505-adopt-claudebrain-assets.md
- [2026-09-24] Adopted factory core + roadmap tier + Python/GitHub/docs families. EXCLUDED (not our stacks): VSTO, VBA,
  Power Platform/SharePoint/Power BI, quant, branding→presentation, Outlook/print HTML. Prose tier deferred to v0.3.2 (was v0.2.2) — same log
- [2026-09-24] Keep pypptx's branch-per-version naming `vX.Y.Z` over `/version-set`'s `claude/<label>-<slug>` default (CLAUDE.md § Roadmap) — same log
- [2026-09-24] Text-style levels are zero-based like `_Paragraph.level`; `phClr` → `MSO_THEME_COLOR.PLACEHOLDER` (17)
  — sessions/2026-09-24-1819-epic-87-text-styles-themes.md
- [2026-09-24] Arc angles are visual (DrawingML), converted to parametric for extents; table merge edits promote the next
  continuation cell and copy the orthogonal span — sessions/2026-09-24-1844-epic-79-shape-line-fill-table.md
- [2026-09-25] Element classes are keyed by tag alone: a tag shared across contexts (`p:custShow`, `p:sld`, `a:path`) gets ONE class;
  a test fails on duplicate `register_element_cls` tags — sessions/2026-09-25-1202-epic-56-presentation-parts.md
- [2026-09-25] SmartArt text edits DROP the cached dsp:drawing part (not patch it); `MSO_SHAPE_TYPE.SMART_ART` is canonical,
  `IGX_GRAPHIC` an alias — sessions/2026-09-25-1400-epic-62-smartart.md
- [2026-09-25] Chartex charts are a SIBLING API (`GraphicFrame.chartex`, `has_chart` False, `shape_type` CHART); the
  shape is the `mc:Choice` graphicFrame and `tree_elm()` gives its wrapper — sessions/2026-09-25-1530-epic-66-chartex-dtable.md
- [2026-09-25] Non-chartex `mc:AlternateContent` shapes are one read-only `AlternateContentShape` each, represented by the
  Fallback shape (else Choice) — sessions/2026-09-25-1600-epic-75-media-altcontent.md
- [2026-09-25] Hot-path caching: `qn` is lru_cached; `BaseOxmlElement.xpath` evaluates per-thread cached compiled XPath
  (`xmlchemy.compiled_xpath`) — sessions/2026-09-25-1630-issue-51-perf-caches.md
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
- `unverified` surfaces (verification-surface.md): type checker, Sphinx build, real-world corpus.
- Anyone with a Windows clone predating `b64d2b9` needs `git add --renormalize .` once after pulling, or git
  reports every text file modified (its dirty check compares the size recorded in the index, not content).
- `tests.yml` caveats: `paths-ignore` (md/.claude/.meta) must go if `tests` ever becomes a REQUIRED check, or
  docs-only PRs wait forever; ~17 billable min/run, 10 of them the single macOS cell (x10 multiplier).
- The #51 perf gate is still absent, deliberately — shared-runner timing noise would make it flake.
- **Actions billing blocks every job.** Fix the spending limit/payment, or cut cost: the lone macOS cell is
  ~10 of ~17 billable min per run. Until then no PR can be CI-verified.
- #115 human check: new 3D charts (column/bar/line/pie/area) open without repair; camera/walls visibly change. No `series_axis` API.
- #114 human check: shape-tagged deck opens in PowerPoint without repair. Follow-up idea: `has_tags` (reading creates a part).

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
