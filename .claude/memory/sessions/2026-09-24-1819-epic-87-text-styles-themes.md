# 2026-09-24 18:19 · epic-87-text-styles-themes

**Goal:** Close out epic #87 (master text styles, theme style lists, theme replacement)

## What happened
- PR #94 merged to main (4f10c81): closed #88, #89, #90; epic-autoclose then closed #87.
- #88: `CT_TextListStyle` (a:lstStyle, p:title/body/otherStyle, p:defaultTextStyle), `CT_SlideMasterTextStyles`;
  API in new `pptx/text/styles.py` — `SlideMaster.text_styles.*[level]`, `Presentation.default_text_style`.
- #89: `Theme.fill_styles` / `line_styles` / `background_fill_styles` (read-only FillFormat/LineFormat).
- #90: `SlideMaster.apply_theme(source)` + `SlideMasterPart.apply_theme`; `OFC_THEME_OVERRIDE` → XmlPart.
- Bug fixes: `ST_TextIndent` (negative indents raised); `MSO_THEME_COLOR.PLACEHOLDER` (phClr raised on read).

## Gotchas & dead ends
- Text-style levels are zero-based (match `_Paragraph.level`); `body[0]` == `a:lvl1pPr`.
- cxml `xml()` comparisons fail on ns-decl placement when an `a:` child is added under a `p:` root — assert via xpath.
- Mutation-testing by cp-restoring a file within the same second leaves a stale .pyc — clear `__pycache__`.
- Unit tests avoid real `Presentation()` in-process (roundtrip tests use a subprocess); `Package.open(template)` is fine.
- Repo has no test CI; only workflow is epic-autoclose.

## State at end
- main has all epic #87 work; 3099 tests pass; no new pyright errors.

## Open threads
- Human-gated: open output in PowerPoint to confirm master text-style changes render on inheriting slides.
- `apply_theme` untested on a PowerPoint-authored .thmx (tests use a minimal hand-built one).
