# Memory archive — 2026

Decisions folded out of INDEX.md to keep it in budget. Still in force unless superseded.

## Decisions archived 2026-09-30

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
