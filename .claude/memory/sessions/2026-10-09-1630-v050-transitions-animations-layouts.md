# 2026-10-09 16:30 · v050-transitions-animations-layouts

**Goal:** Build v0.5.0: read transitions/animations, custom layouts, handout master

## What happened
- Card expanded from a 511-deck pilot (POI test-data + LibreOffice sd/qa PPTX + corpus; 464 PowerPoint-saved), owner
  accepted all 5 recommendations, then built in 7 steps on `claude/epic-87-closeout-hfo5vj`:
  - c34efc7 corpus +4 POI decks (customGeo, bug68703, EmbeddedVideo, 2411-Performance_Up; SVN r1938723).
  - 5d030a8 OXML: oxml/transition.py (`effective_transition` resolves mc:AlternateContent), oxml/timing.py (moved
    timing/media classes from oxml/slide.py), enum/animation.py, ST_TLTime, p159 namespace.
  - b1f4add API: `SlideTransition` on slide/layout/master; `Slide.animations` → `Animation`.
  - 207a157 handout master (`HandoutMasterPart`, `Presentation.handout_master`, never created).
  - 181aeb8 layouts: `SlideLayouts.add_slide_layout/duplicate`, `LayoutShapes.add_placeholder`,
    `PresentationPart.next_master_or_layout_id()`; FIXED import_slide's imported master keeping the source's whole
    sldLayoutIdLst (dangling rIds → slide_layouts raised; ids collided).
  - 901c8a1 deckwalk reads timing (`walk_timing`), read-then-save fidelity test, pinned corpus facts; 97a9f16 examples,
    Sphinx pages, manifest/gaps.
- 9cff1c1: six EXISTING tests were never collected (and_*/but_* without "it"); renamed (all pass) + tests/test_test_names.py guard.

## Gotchas & dead ends
- presetID is per class. MsoAnimEffect == file presetID only for entrance/exit 1–31 (checked vs LibreOffice's
  oox/source/ppt/commontimenodecontext.cxx); emphasis/path IDs differ (emph 8 Spin vs msoAnimEffectSpin 61).
- Legacy getters WRITE on read (upstream python-pptx): `Font.color` makes the fill solid, `_Paragraph.alignment`
  adds a:pPr. So the full deckwalk can't back a read-then-save test; `walk_timing` can.
- lxml can't C14N bug68703's customXml/item3.xml → fidelity compare falls back to bytes on C14NError.
- cxml can't express a digit-bearing attribute prefix (`p14:dur`): use literal XML.
- Pilot: LibreOffice's tdf169781.pptx (deliberately broken) lists a layout rel that doesn't exist → slide_layouts raises.

## State at end
- Suite 4366 passed; coverage 94%; sphinx -W clean; ruff 22 (= baseline); pyright no new errors on touched files.
- Work-count budget 1,009,255 (28 decks, walk includes timing reads).

## Open threads
- Human checks: custom_layouts.py output in PowerPoint (gallery, no repair, "1_" naming); transition/animation values
  vs PowerPoint panes (customGeo, bug68703); whether master/layout transitions apply to slides.
- Follow-up candidates: make Font.color / paragraph.alignment read-only on read; layout non-placeholder shapes (v0.5.x).
