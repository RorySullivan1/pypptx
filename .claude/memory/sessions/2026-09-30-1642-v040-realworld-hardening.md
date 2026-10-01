# 2026-09-30 16:42 · v040-realworld-hardening

**Goal:** Build v0.4.0: real-world corpus, damaged input, exception audit, Hypothesis

## What happened
- Owner decisions: POI files OK **for testing only**; geometry-less `p:sp` → `shape_type` AUTO_SHAPE.
- Corpus: 24 decks + 4 fuzz files from Apache POI r1938723 in `tests/test_files/real_world/` (SOURCES.md
  with SHA-256/app/features, LICENSE, NOTICE). Web-scraped POI decks deliberately excluded.
- `tests/test_real_world_roundtrip.py` (+ `tests/unitutil/deckwalk.py`): deep read walk, edit/save/reload,
  unedited save changes no XML (blank text ignored) and keeps every content type; damaged-input cases.
- **59 tests had never run**: pytest only collected `it_*`; `and_it_*`/`but_it_*` now collected. 5 failed:
  they expected ValueError where code raised ShapeError/ChartError → domain errors now also subclass
  ValueError (Shape/Slide/Chart/TableError). New InvalidValueError(ValueError), InvalidTypeError(TypeError),
  InvalidPackageError(PackageError). All bare API-reachable raises converted keeping their builtin base.
- Damaged packages → InvalidPackageError (not-a-zip, corrupt member, no [Content_Types].xml, no main part);
  `_ZipPkgReader.__del__` guarded.
- Effects: BaseShape reads shadow/glow/reflection/soft_edge via `_effect_properties`, three_d via
  `_three_d_properties`; groups use grpSpPr (no 3D), graphic frames use the table's tblPr / chart's c:spPr.
- Hypothesis over all 59 writable simple types found 5 real bugs (bool written as "True", NaN/inf,
  ST_Angle overflow, ST_PositiveFixedAngle writing 21600000, ST_UniversalMeasure KeyError) — all fixed.
- Review round: corpus had never been committed (.gitignore `*.pptx`); fixed + `!tests/test_files/**/*.pptx`.
  Stricter fidelity test (rels + blobs) found `#fragment` hyperlink rels dropped on load — fixed, kept internal
  on copy. UnsupportedEffectError added; effect reads no longer write XML; MANIFEST.in prunes the corpus.
- Final: 4131 passed; coverage 93.49%; sphinx -W clean; ruff 22 = main; pyright same as main.

## Gotchas & dead ends
- Pilot claim "`fill` fails with shape_type" was the pilot script's own ordering; the walker proved it false.
- Errors in `__del__` go to `sys.unraisablehook`, not stderr — capture the hook to test them.
- `--hypothesis-show-statistics` reports nothing for `@given` nested inside a parametrized test; measure
  examples with a counter instead (did: 200 per type).
- `.gitignore` ignores `*.pptx`: new test decks must be un-ignored or they silently stay local.
- Keep the builtin base when converting a raise: a ValueError raised for a wrong TYPE becomes
  InvalidValueError, not InvalidTypeError, or `except ValueError` callers break.

## State at end
- v0.4.0 implemented on the session branch, PR #122; review + goal audit done. Needs owner sign-off on scope
  additions (see card "As built") and the PowerPoint check before merge.

## Open threads
- Human check: open 3 round-tripped corpus decks (one chart, one SmartArt) + a table/chart shadow in PowerPoint.
- `AlternateContentShape` effects follow its represented element; not exercised.
- Shipped 2026-10-01: owner merged PR #122 (2c64b2d), accepting the scope additions; package 0.4.0, cursor → v0.5.0.
