# 2026-09-25 12:02 · epic-56-presentation-parts

**Goal:** Implement epic #56 (sub-issues #57-#61) and ship it

## What happened
- Three parallel `python-developer` worktree agents: #57+#58 (presentation.xml), #59+#60 (new parts), #61 (fonts).
  Merged into one branch; PR #99 merged (bd626fc); #56–#61 closed (epic via the autoclose workflow).
- #57 `Presentation.notes_width/notes_height` (`CT_NotesSize`).
- #58 `Presentation.custom_shows` (`pptx/custom_show.py`); `Slides.delete` prunes refs via
  `PresentationPart.drop_custom_show_refs`; removing the selected show resets `slide_show_settings` to all slides.
- #59 `PresPropsPart` + `Presentation.slide_show_settings` (`PP_SLIDE_SHOW_TYPE` in `enum/pres.py`); part made on first write.
- #60 `TableStylesPart` + read-only `Presentation.table_styles` (id→name, `default_id`).
- #61 `Presentation.embedded_fonts` (`pptx/fonts.py`, `oxml/embeddedfont.py`); `remove()` drops rels, and on the last
  font removes `p:embeddedFontLst` and sets `embedTrueTypeFonts` off. Font CTs registered as blob `Part`.
- New `LazyXmlPart` (opc/package.py): keeps loaded bytes, parses on first `_element` access → untouched parts are
  byte-stable (reading via the API counts as touching).

## Gotchas & dead ends
- Element classes are looked up by tag alone: `p:custShow` is both a custShowLst definition and the showPr
  selection; two agents registered different classes and the later one silently broke custom shows. Now one class
  (`CT_CustomShow`) + a test that fails on any duplicate `register_element_cls` tag. `p:sld` and `a:path` are
  shared tags too — handled with generic lxml calls, not a second class.
- Merging parallel additions to `CT_Presentation._tag_seq` needs every `successors=_tag_seq[N:]` index re-derived
  (order: notesSz, embeddedFontLst, custShowLst, kinsoku, defaultTextStyle, extLst).
- The shipped template's tableStyles.xml has only `def`, no `a:tblStyle` — `table_styles` is empty on a new deck.

## State at end
- `main` @ bd626fc; 3402 tests pass; all examples run; no new pyright/ruff findings vs main.

## Open threads
- Human PowerPoint checks: custom show plays the right slides; loop/kiosk take effect; a file with its last embedded
  font removed opens without repair.
- No real embedded-font .pptx in tests/test_files — tests build the fixture in code.
