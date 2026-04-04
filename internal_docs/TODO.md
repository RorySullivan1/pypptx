# pypptx TODO

Planned additions and improvements for the pypptx library.

**Strategy:** Bottom-up, types-first. Complete the OXML layer (XML element classes,
simple types, enumerations) before building the Python API wrappers on top. This
avoids context-switching between layers and exposes schema gaps early.

---

## 1. Core Reorganization & Structural Improvements (DONE)

### 1.1 Project Packaging
- [x] Added `pyproject.toml` with metadata, dependencies (`lxml`, `typing_extensions`, `XlsxWriter`, `Pillow`), dev extras (`pytest`, `pytest-cov`), and pytest configuration
- [x] Package installs and runs via `pip install -e ".[dev]"`

### 1.2 Exception Hierarchy (DONE)
- [x] Audited all ~130 raise sites across 38 files
- [x] Added domain-specific exceptions: `ShapeError`, `SlideError`, `ChartError`, `TableError`, `PackageError`
- [x] Updated key call sites in `shapes/`, `slide.py`, `chart/`, `table.py`, `parts/presentation.py`
- [x] Updated corresponding test assertions
- [x] Removed backward-compatibility alias for `PythonPptxError`
- [x] Migrated remaining generic `ValueError`/`TypeError` raises in API modules to domain exceptions

### 1.3 Shape Tree & Lookup (DONE)
- [x] Added `_BaseShapes.get_by_name(name, default)` for name-based lookup
- [x] Added `_BaseShapes.get_by_id(shape_id, default)` for ID-based lookup
- [x] Added `_BaseShapes.__contains__` supporting both name (str) and shape object
- [x] Built cached index dicts (`_name_index`, `_id_index`) with invalidation on all mutation paths

### 1.4 Proxy & Base Class Cleanup
- [x] Removed `(object)` explicit base class from 90 files (Python 2 artifact)
- [x] Reviewed proxy hierarchy — `ElementProxy`/`ParentedElementProxy`/`PartElementProxy` is sound
- [x] `del` statements in `__init__.py` are intentional namespace cleanup, retained

### 1.5 Part Loading & Memory (DONE)
- [x] Evaluated — blobs are loaded eagerly in `_PackageLoader._parts` via `PartFactory`
- [x] Refactored `_ZipPkgReader` to read blobs on demand instead of loading all into memory at once

### 1.6 Type Annotations
- [x] `py.typed` marker already present
- [x] `from __future__ import annotations` in all 92 source files
- [ ] Incremental annotation improvements deferred to feature work

### 1.7 Test Infrastructure
- [x] pytest configured in `pyproject.toml` (test paths, class/function patterns)
- [x] 2644 tests passing, 97% code coverage
- [ ] Add round-trip integration tests (deferred)

### 1.8 OXML Layer (DONE)
- [x] Added `p14` and `p15` namespace prefixes for sections and modern comments
- [x] `xmlchemy.py` and `simpletypes.py` audit moved to Section 2 (OXML Foundation Audit)

---

## 2. OXML Foundation Audit (DONE)

Audit and harden the OXML infrastructure before adding new element types.

### 2.1 `simpletypes.py` Audit (DONE)
- [x] Reviewed all ~50 existing simple type converters for completeness and correctness
- [x] Removed dead Python 2 `basestring` check from `validate_string`
- [x] Added missing simple types needed by planned features:
  - `ST_TextColumnCount` — int 1..16 (text columns on `a:bodyPr`)
  - `ST_TextMargin` — int 0..51206400 EMU (paragraph indent/margin on `a:pPr`)
  - `ST_TextNonNegativePoint` — int >= 0 centipoints (kerning on `a:rPr`)
  - `ST_TextPoint` — union of int centipoints and universal measure (character spacing on `a:rPr`)
  - `ST_PositiveCoordinate32` — non-negative 32-bit int EMU (column spacing on `a:bodyPr`)
- [x] Existing converters handle edge cases correctly (range validation, type checking)

### 2.2 `xmlchemy.py` Audit (DONE)
- [x] All needed descriptors present: `ZeroOrOne`, `ZeroOrMore`, `OneOrMore`, `OneAndOnlyOne`, `ZeroOrOneChoice`, `Choice`, `OptionalAttribute`, `RequiredAttribute`
- [x] `OptionalAttribute` supports boolean via `XsdBoolean`
- [x] **Limitation:** No built-in extension element support — decorative flag on `cNvPr` will require manual `a:ext` element handling (not a descriptor gap, just a known pattern)

### 2.3 Namespace Registry (DONE)
- [x] `p14`, `p15` confirmed present for sections and modern comments
- [x] Added `a14` (`http://schemas.microsoft.com/office/drawing/2010/main`) for DrawingML 2010 extensions
- [x] Added `adec` (`http://schemas.microsoft.com/office/drawing/2017/decorative`) for decorative flag support

---

## 3. OXML Type Definitions — Missing Attributes on Existing Elements (DONE)

Add attributes and child element declarations to element classes that already
exist but are incomplete.

### 3.1 `oxml/shapes/shared.py` — `CT_NonVisualDrawingProps` (`cNvPr`) (DONE)
- [x] `descr` — optional `XsdString` attribute (alternative text for accessibility)
- [x] `title` — optional `XsdString` attribute (shape title for accessibility)
- [x] `hidden` — optional `XsdBoolean` attribute (shape visibility)
- [ ] Decorative flag — requires manual `a:ext` element handling (deferred to Section 8)

### 3.2 `oxml/shapes/shared.py` — Shape lock elements (DONE)
- [x] Created `CT_Locking` class with `noChangeAspect` optional boolean attribute
- [x] Registered for `a:spLocks`, `a:picLocks`, `a:cxnSpLocks`, `a:grpSpLocks`

### 3.3 `oxml/text.py` — `CT_TextCharacterProperties` (`a:rPr`) (DONE)
- [x] `strike` — optional `XsdString` attribute (enum type deferred to Section 5)
- [x] `baseline` — optional `ST_Percentage` attribute (superscript/subscript)
- [x] `cap` — optional `XsdString` attribute (enum type deferred to Section 5)
- [x] `spc` — optional `ST_TextPoint` attribute (character spacing)
- [x] `kern` — optional `ST_TextNonNegativePoint` attribute (kerning threshold)

### 3.4 `oxml/text.py` — `CT_TextParagraphProperties` (`a:pPr`) (DONE)
- [x] `indent` — optional `ST_TextMargin` attribute (first-line indent)
- [x] `marL` — optional `ST_TextMargin` attribute (left margin)
- [x] `rtl` — optional `XsdBoolean` attribute (right-to-left text direction)
- [x] `hangingPunct` — optional `XsdBoolean` attribute
- [x] `fontAlgn` — optional `XsdString` attribute (enum type deferred to Section 5)

### 3.5 `oxml/text.py` — `CT_TextBodyProperties` (`a:bodyPr`) (DONE)
- [x] `vert` — optional `XsdString` attribute (enum type deferred to Section 5)
- [x] `numCol` — optional `ST_TextColumnCount` attribute (text columns 1..16)
- [x] `spcCol` — optional `ST_PositiveCoordinate32` attribute (column spacing)

### 3.6 `oxml/shapes/shared.py` — `CT_LineProperties` (`a:ln`) (DONE)
- [x] `cmpd` — optional `XsdString` attribute (enum type deferred to Section 5)

### 3.7 `oxml/presentation.py` — `CT_Presentation` (DONE)
- [x] `firstSlideNum` — optional `XsdInt` attribute (first slide number)

---

## 4. OXML Type Definitions — New Element Classes (DONE)

New `CT_*` element classes defined and registered with `xmlchemy`.

### 4.1 Shadow Elements — `oxml/dml/effect.py` (DONE)
- [x] `CT_OuterShadowEffect` (`a:outerShdw`) — `blurRad`, `dist`, `dir`, `algn`, `rotWithShape` attributes
- [x] `CT_InnerShadowEffect` (`a:innerShdw`) — `blurRad`, `dist`, `dir` attributes
- [x] `CT_EffectList` (`a:effectLst`) — declares `glow`, `innerShdw`, `outerShdw`, `reflection`, `softEdge` as `ZeroOrOne` children

### 4.2 Additional Effect Elements — `oxml/dml/effect.py` (DONE)
- [x] `CT_ReflectionEffect` (`a:reflection`) — blur, start/end alpha, distance, direction, rotWithShape
- [x] `CT_GlowEffect` (`a:glow`) — radius attribute, color child
- [x] `CT_SoftEdgesEffect` (`a:softEdge`) — radius attribute

### 4.3 Table Cell Border Elements — `oxml/table.py` (DONE)
- [x] `CT_TableCellProperties` (`a:tcPr`) declares `a:lnL`, `a:lnR`, `a:lnT`, `a:lnB` as `ZeroOrOne` children
- [x] Added `a:lnTlToBr`, `a:lnBlToTr` diagonal border children on `a:tcPr`
- [x] Added `tblStyle` attribute on `CT_TableProperties` (`a:tblPr`)

### 4.4 Headers & Footers — `oxml/slide.py` (DONE)
- [x] `CT_HeaderFooter` (`p:hf`) — boolean attributes: `sldNum`, `hdr`, `ftr`, `dt`
- [x] Registered as `p:hf` (already referenced in tag sequences of NotesMaster, SlideLayout, SlideMaster)

### 4.5 Sections — `oxml/section.py` (DONE)
- [x] `CT_SectionList` (`p14:sectionLst`) — container with `ZeroOrMore` section children
- [x] `CT_Section` (`p14:section`) — `name` attribute, `ZeroOrMore` `p14:sldId` children
- [x] `CT_SectionSlideIdListEntry` (`p14:sldId`) — slide reference within section

### 4.6 Comments — `oxml/comment.py` (DONE)
- [x] `CT_Comment` (`p:cm`) — `authorId`, `idx` attributes; `pos` and `text` children
- [x] `CT_CommentList` (`p:cmLst`) — container for comments
- [x] `CT_CommentAuthor` (`p:cmAuthor`) — `id`, `name`, `initials`, `lastIdx`, `clrIdx`
- [x] `CT_CommentAuthorList` (`p:cmAuthorLst`) — container for authors

### 4.7 Picture Format Effects — `oxml/dml/picture.py` (DONE)
- [x] `CT_LuminanceEffect` (`a:lum`) — `bright`, `contrast` attributes
- [x] `CT_GrayscaleEffect` (`a:grayscl`) — no attributes needed
- [x] `CT_DuotoneEffect` (`a:duotone`) — color children

### 4.8 3D Formatting — `oxml/dml/threed.py` (DONE)
- [x] `CT_Shape3D` (`a:sp3d`) — `extrusionH`, `contourW`, `prstMaterial`; bevel children
- [x] `CT_Bevel` (`a:bevelT`/`a:bevelB`) — `w`, `h`, `prst` attributes
- [x] `CT_Scene3D` (`a:scene3d`) — `camera` and `lightRig` children
- [x] `CT_Camera` (`a:camera`) — `prst`, `fov` attributes
- [x] `CT_LightRig` (`a:lightRig`) — `rig`, `dir` attributes

### 4.9 Theme Elements — `oxml/theme.py` (DONE)
- [x] `CT_ColorScheme` (`a:clrScheme`) — all 12 color slots (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink)
- [x] `CT_FontScheme` (`a:fontScheme`) — `majorFont` and `minorFont` children
- [x] `CT_FontCollection` (`a:majorFont`/`a:minorFont`) — latin, ea, cs children
- [x] `CT_BaseStyles` (`a:themeElements`) — `clrScheme` and `fontScheme` children
- [x] `CT_OfficeStyleSheet` updated with `themeElements` child

### 4.10 WordArt — `oxml/text.py` (DONE)
- [x] `CT_PresetTextShape` (`a:prstTxWarp`) — `prst` attribute, `avLst` child
- [x] Registered as `ZeroOrOne` child (`prstTxWarp`) on `CT_TextBodyProperties`

---

## 5. Enumerations (DONE)

`BaseXmlEnum` definitions in `pptx/enum/` with OXML attribute wiring.

### Text enumerations — `enum/text.py` (DONE)
- [x] `MSO_TEXT_STRIKE_TYPE` — `noStrike`, `sngStrike`, `dblStrike` → wired to `strike` on `CT_TextCharacterProperties`
- [x] `MSO_TEXT_CAPS` — `none`, `all`, `small` → wired to `cap` on `CT_TextCharacterProperties`
- [x] `MSO_TEXT_FONT_ALIGN` — `auto`, `t`, `ctr`, `base`, `b` → wired to `fontAlgn` on `CT_TextParagraphProperties`
- [x] `MSO_TEXT_VERTICAL_TYPE` — `horz`, `vert`, `vert270`, `wordArtVert`, `eaVert`, `mongolianVert`, `wordArtVertRtl` → wired to `vert` on `CT_TextBodyProperties`
- [x] `MSO_PRESET_TEXT_SHAPE` — 41 text warp presets → wired to `prst` on `CT_PresetTextShape`

### DML enumerations — `enum/dml.py` (DONE)
- [x] `MSO_LINE_COMPOUND_TYPE` — `sng`, `dbl`, `thickThin`, `thinThick`, `tri` → wired to `cmpd` on `CT_LineProperties`
- [x] `MSO_RECT_ALIGNMENT` — 9 positions (tl, t, tr, l, ctr, r, bl, b, br) → wired to `algn` on `CT_OuterShadowEffect`

### Deferred
- [ ] Shadow style enumeration — deferred to Section 12 (Python API shadow implementation)

---

## 5.1 OXML Foundation Gap Fixes (DONE)

Close gaps where elements exist in `_tag_seq` / successor lists but lack
`ZeroOrOne`/`ZeroOrMore` descriptor declarations, making them inaccessible
programmatically.

### Slide elements (DONE)
- [x] `p:hf` declared as `ZeroOrOne` on `CT_NotesMaster`, `CT_SlideLayout`, `CT_SlideMaster`
- [x] `p:transition` declared as `ZeroOrOne` on `CT_Slide`, `CT_SlideLayout`

### Text character properties (DONE)
- [x] `a:ea`, `a:cs`, `a:sym` declared as `ZeroOrOne` (`CT_TextFont`) on `CT_TextCharacterProperties`
- [x] `a:highlight` declared as `ZeroOrOne` on `CT_TextCharacterProperties`
- [x] Registered `a:ea`, `a:cs`, `a:sym` element classes as `CT_TextFont`

### Bullet elements on paragraph properties (DONE)
- [x] Declared all 11 bullet child elements on `CT_TextParagraphProperties`:
  `buClrTx`, `buClr`, `buSzTx`, `buSzPct`, `buSzPts`, `buFontTx`, `buFont`,
  `buNone`, `buAutoNum`, `buChar`, `buBlip`

### Color modifiers (DONE)
- [x] Added `tint`, `shade`, `satMod`, `satOff`, `lumMod`, `lumOff`, `alpha` on `_BaseColorElement`
- [x] Registered `a:alpha`, `a:satMod`, `a:satOff`, `a:shade`, `a:tint` as `CT_Percentage`

### Blip fill (DONE)
- [x] `a:tile` and `a:stretch` declared as `ZeroOrOne` on `CT_BlipFillProperties`

### Shape style (DONE)
- [x] `p:style` declared as `ZeroOrOne` on `CT_Shape` and `CT_Connector`

---

## 6. Python API — Slide Lifecycle (NEXT)

_Prerequisites: Section 3.7 (`firstSlideNum` attribute)_

- [ ] Delete a slide — remove `sldId` from `sldIdLst`, delete slide part and relationships
- [ ] Reorder slides — reorder `sldId` entries within `sldIdLst`, rename parts
- [ ] Duplicate a slide — deep-clone slide part, remap relationships for images/charts/media
- [ ] Slide number property — compute from position in `sldIdLst` + `firstSlideNum`
- [ ] `Presentation.first_slide_number` — expose `firstSlideNum` attribute

---

## 7. Python API — Shape Lifecycle

_Prerequisites: Sections 3.1 (`hidden`), 3.2 (`noChangeAspect`)_

- [ ] Delete a shape — remove element from `spTree`, clean up related relationships (images, hyperlinks)
- [ ] Duplicate a shape — clone element in `spTree`, assign new shape ID, clone related parts
- [ ] Z-order control — `bring_to_front()`, `send_to_back()`, `bring_forward()`, `send_backward()` via element reordering in `spTree`
- [ ] Shape visibility — expose `hidden` attribute on `cNvPr` as `shape.visible` property
- [ ] Lock aspect ratio — expose `noChangeAspect` on shape lock elements
- [ ] Parent group reference — back-reference from child shape to containing `GroupShape`

---

## 8. Python API — Accessibility

_Prerequisites: Section 3.1 (`descr`, `title`, decorative flag)_

- [ ] `shape.alternative_text` — read/write `descr` attribute on `cNvPr`
- [ ] `shape.title` — read/write `title` attribute on `cNvPr`
- [ ] `shape.decorative` — read/write decorative flag (extension element on `cNvPr`)

---

## 9. Python API — Text & Font Enhancements

_Prerequisites: Sections 3.3, 3.4, 3.5 (text attributes), Section 5 (enumerations)_

### 9.1 Font Properties
- [ ] Strikethrough — `strike` attribute on `a:rPr` (single, double, none)
- [ ] Superscript / subscript — `baseline` attribute on `a:rPr`
- [ ] Character caps — `cap` attribute on `a:rPr` (none, all, small)
- [ ] Character spacing — `spc` attribute on `a:rPr`
- [ ] Kerning — `kern` attribute on `a:rPr`
- [ ] Font shadow — effect list child on `a:rPr`
- [ ] East Asian font name — `a:ea` element on `a:rPr`
- [ ] Complex script font name — `a:cs` element on `a:rPr`

### 9.2 Paragraph Properties
- [ ] Bullet formatting — `a:buChar`, `a:buAutoNum`, `a:buFont`, `a:buClr`, `a:buSzPct`, `a:buSzPts`, `a:buNone`
- [ ] First-line indent — `indent` attribute on `a:pPr`
- [ ] Left margin — `marL` attribute on `a:pPr`
- [ ] Tab stops — `a:tabLst` with `a:tab` children (position, alignment)
- [ ] Text direction / RTL — `rtl` attribute on `a:pPr`
- [ ] Hanging punctuation — `hangingPunct` attribute on `a:pPr`
- [ ] Baseline alignment — `fontAlgn` attribute on `a:pPr`

### 9.3 Text Frame Properties
- [ ] Orientation — `vert` attribute on `a:bodyPr` (horizontal, vertical, stacked, etc.)
- [ ] Text columns — `numCol` and `spcCol` attributes on `a:bodyPr`
- [ ] `has_text` property — boolean check for non-empty text content

---

## 10. Python API — Table Improvements

_Prerequisites: Section 4.3 (table cell border elements)_

- [ ] Cell borders — per-edge control via `a:lnL`, `a:lnR`, `a:lnT`, `a:lnB` within `a:tcPr` (color, weight, dash style per edge)
- [ ] Diagonal borders — `a:lnTlToBr`, `a:lnBlToTr` within `a:tcPr`
- [ ] Table style — get/set built-in style GUID via `tblStyle` attribute on `a:tblPr`

---

## 11. Python API — Line & Connector Enhancements

_Prerequisites: Section 3.6 (`cmpd` attribute)_

- [ ] Arrowhead formatting — `a:headEnd` and `a:tailEnd` within `a:ln` (type, width, length)
- [ ] Compound line style — `cmpd` attribute on `a:ln` (single, double, thick-thin, etc.)
- [ ] Line transparency — alpha modifier on line fill color
- [ ] Line visibility — no-fill vs filled state
- [ ] Line pattern — pattern fill on lines
- [ ] Query connected shapes — read `a:stCxn` / `a:endCxn` attributes on connectors

---

## 12. Python API — Shadow (Full Implementation)

_Prerequisites: Section 4.1 (shadow elements)_

The current `ShadowFormat` is a stub exposing only `inherit`. Replace with full implementation:

- [ ] Shadow type — outer (`a:outerShdw`), inner (`a:innerShdw`) within `a:effectLst`
- [ ] Blur radius — `blurRad` attribute
- [ ] Distance and direction — `dist` and `dir` attributes
- [ ] Shadow color with transparency — color child element with alpha
- [ ] Alignment — `algn` attribute
- [ ] Rotate with shape — `rotWithShape` attribute
- [ ] Visibility — presence/absence of shadow element
- [ ] Shadow style enumeration

---

## 13. Python API — Additional Visual Effects

_Prerequisites: Section 4.2 (effect elements)_

- [ ] Reflection — `a:reflection` in `a:effectLst` (blur, start/end alpha, distance, direction, scale)
- [ ] Glow — `a:glow` in `a:effectLst` (radius, color)
- [ ] Soft edge — `a:softEdge` in `a:effectLst` (radius)

---

## 14. Python API — Headers & Footers

_Prerequisites: Section 4.4 (`CT_HeaderFooter`)_

- [ ] Slide-level header/footer configuration — `p:hf` element
- [ ] Date/time placeholder — automatic vs fixed, format string
- [ ] Footer text placeholder
- [ ] Slide number placeholder
- [ ] Per-slide show/hide overrides
- [ ] Notes and handout header/footer support

---

## 15. Python API — Metadata & Organization

_Prerequisites: Section 4.5 (sections and tags elements)_

- [ ] Tags — key-value string pairs on shapes and slides (separate `tags[N].xml` parts linked via relationships)
- [ ] Sections — named slide groups via `p14:sectionLst` in presentation extensions (add, remove, rename, reorder, list sections)
- [ ] Custom document properties — beyond core properties (custom key-value metadata)

---

## 16. Python API — Comments

_Prerequisites: Section 4.6 (comment elements)_

- [ ] Slide comments — add, read, delete comments (`comments[N].xml` parts)
- [ ] Comment authors — manage author list (`commentAuthors.xml` part)
- [ ] Comment positioning — x/y coordinates on slide
- [ ] Comment metadata — author, datetime, text

---

## 17. Python API — Picture Format Enhancements

_Prerequisites: Section 4.7 (picture format effects)_

- [ ] Brightness — `a:lum` bright attribute on blip
- [ ] Contrast — `a:lum` contrast attribute on blip
- [ ] Grayscale / black-and-white / washout — `a:grayscl`, `a:duotone` effects on blip
- [ ] Transparency color
- [ ] Original image dimensions — expose from `ImagePart`

---

## 18. Chart Buildout — Full Chart Capabilities (NEXT)

High-value for reporting automation. The existing chart subsystem covers 9 of
16+ chart types and lacks trendlines, error bars, secondary axes, and several
formatting APIs. This section closes all gaps.

_Prerequisites: None — chart OXML and API are self-contained subsystems._

### 18.1 Missing OXML Chart Type Classes (DONE)

Define CT_* element classes for chart types that have enumerations but no OXML
implementation, then register them in `__init__.py`.

- [x] `CT_Bar3DChart` (`c:bar3DChart`) — `barDir`, `grouping`, `varyColors`, `ser`, `gapWidth`, `gapDepth`, `shape`, `axId` children
- [x] `CT_Line3DChart` (`c:line3DChart`) — `grouping`, `varyColors`, `ser`, `gapDepth`, `axId` children
- [x] `CT_Pie3DChart` (`c:pie3DChart`) — `varyColors`, `ser` children
- [x] `CT_StockChart` (`c:stockChart`) — `ser`, `axId`, `hiLowLines`, `upDownBars` children
- [x] `CT_SurfaceChart` (`c:surfaceChart`) — `wireframe`, `ser`, `bandFmts`, `axId` children
- [x] `CT_Surface3DChart` (`c:surface3DChart`) — same structure as surface
- [x] `CT_OfPieChart` (`c:ofPieChart`) — `ofPieType`, `varyColors`, `ser`, `gapWidth`, `splitType`, `splitPos`, `custSplit`, `secondPieSize`, `serLines` children
- [x] Register all new classes in `oxml/__init__.py`
- [x] Wire into `PlotFactory` / `PlotTypeInspector` so existing API recognizes them
- [x] API plot classes: `Bar3DPlot`, `Line3DPlot`, `Pie3DPlot`, `StockPlot`, `SurfacePlot`, `Surface3DPlot`, `OfPiePlot`

### 18.2 Trendlines — OXML + API (DONE)

Most-requested chart feature for reporting automation.

- [x] `CT_Trendline` (`c:trendline`) — `name`, `spPr`, `trendlineType`, `order`, `period`, `forward`, `backward`, `intercept`, `dispRSqr`, `dispEq`, `trendlineLbl` children
- [x] `CT_TrendlineType` (`c:trendlineType`) — `val` attribute (linear, exponential, logarithmic, movingAvg, polynomial, power)
- [x] `CT_TrendlineLabel` (`c:trendlineLbl`) — layout, `numFmt`, `spPr`, `txPr`, `tx` children
- [x] Declare `c:trendline` as `ZeroOrMore` on `CT_SeriesComposite` (+ `c:errBars`)
- [x] Register new classes in `__init__.py` (+ `c:forward`, `c:backward`, `c:intercept`, `c:dispRSqr`, `c:dispEq`, `c:period`)
- [x] `Trendline` API class — type, order, period, forward/backward, intercept, display R², display equation, name, format
- [x] `TrendlineCollection` on series — add/remove/iterate trendlines via `series.trendlines`
- [x] `XL_TRENDLINE_TYPE` enumeration (EXPONENTIAL, LINEAR, LOGARITHMIC, MOVING_AVERAGE, POLYNOMIAL, POWER)
- [x] Fixed `_SeriesFactory` to handle all 16 chart types (was missing area3D, bar3D, line3D, pie3D, ofPie, stock, surface, surface3D)

### 18.3 Error Bars — OXML + API (DONE)

- [x] `CT_ErrBars` (`c:errBars`) — `errDir`, `errBarType`, `errValType`, `noEndCap`, `plus`, `minus`, `spPr` children + `val_val` property for `c:val` (tag conflict with `CT_NumDataSource`)
- [x] `CT_ErrBarType` / `CT_ErrValType` / `CT_ErrDir` — val-attribute elements for error bar configuration
- [x] `c:errBars` declared as `ZeroOrMore` on `CT_SeriesComposite` (done in 18.2)
- [x] Register new classes in `__init__.py` (+ `c:noEndCap` as `CT_Boolean`)
- [x] `ErrorBars` API class — type, direction, include (both/plus/minus), value, has_end_cap, format
- [x] `ErrorBarsCollection` — has_error_bars, add(type, value), remove, iteration
- [x] Access via `series.error_bars` property on all series types
- [x] `XL_ERROR_BAR_TYPE` (CUSTOM, FIXED_VALUE, PERCENT, ST_DEV, ST_ERROR)
- [x] `XL_ERROR_BAR_DIRECTION` (X, Y)
- [x] `XL_ERROR_BAR_INCLUDE` (BOTH, MINUS_VALUES, PLUS_VALUES)

### 18.4 Secondary Axis — API (DONE)

OXML already supports multiple axes via `ZeroOrMore`. API exposure added.

- [x] `Chart.secondary_value_axis` — access second `c:valAx` if present
- [x] `Chart.secondary_category_axis` — access second `c:catAx`/`c:dateAx` if present
- [x] Fix `Chart.value_axis` heuristic — was returning `valAx_lst[1]` when count > 1, now returns `[0]` (primary)
- [ ] Axis `axId` / `crossAx` cross-referencing — associate axes with correct plots (deferred to 18.5)
- [ ] `Plot.axis_ids` property — expose which axis IDs a plot references (deferred to 18.5)
- [ ] Combo chart axis assignment — when adding a second plot, assign correct axis pair (deferred to 18.5)

### 18.5 Combo Charts — API

Reading existing combo charts works. Need creation support.

- [ ] `Chart.add_plot()` — add a second plot type to an existing chart (e.g. line on top of bar)
- [ ] Axis assignment for new plots — auto-create secondary axes when needed
- [ ] `Chart.plots` — already exists, verify iteration over multiple `xChart` elements works correctly

### 18.6 3D Chart Properties

- [ ] `CT_View3D` (`c:view3D`) — `rotX`, `rotY`, `rAngAx`, `perspective`, `depthPercent`, `heightPercent` children
- [ ] Declare `c:view3D` as `ZeroOrOne` on `CT_Chart` (likely already in `_tag_seq`)
- [ ] `Chart.view_3d` API — rotation X/Y, right-angle axes, perspective, depth/height percent
- [ ] `CT_Surface` / `CT_Floor` / `CT_SideWall` / `CT_BackWall` — `spPr`, `thickness` children
- [ ] Declare floor/walls on `CT_Chart`
- [ ] `Chart.floor`, `Chart.back_wall`, `Chart.side_wall` API properties

### 18.7 Chart & Plot Area Formatting

- [ ] Plot area position — manual layout via `c:layout` / `c:manualLayout` (x, y, w, h)
- [ ] Plot area formatting — fill and line via `c:spPr` on `c:plotArea`
- [ ] Chart area formatting — fill and line via `c:spPr` on `c:chartSpace`
- [ ] `c:dispBlanksAs` — gap, zero, span for missing data points

### 18.8 Line/Area/Bar Chart Special Elements

- [ ] `c:hiLowLines` — high-low lines on line/stock charts (spPr for formatting)
- [ ] `c:upDownBars` — up/down bars on line/stock charts (gapWidth, upBars, downBars with spPr)
- [ ] `c:dropLines` — drop lines on line/area charts (spPr for formatting)
- [ ] `c:serLines` — series connector lines on bar/pie charts (spPr for formatting)
- [ ] Expose on relevant plot types as properties with formatting access

### 18.9 Data Label Enhancements

- [ ] Separator text — `c:separator` element on `c:dLbls`
- [ ] Leader line formatting — `c:leaderLines` spPr on `c:dLbls`
- [ ] Individual data label override — `c:dLbl` per-point customization (already partially implemented)

### 18.10 Axis Enhancements

- [ ] Axis label rotation — `c:txPr` on axis with `a:bodyPr` rot attribute
- [ ] Display units — `c:dispUnits` (hundreds, thousands, millions, etc.)
- [ ] Log scale — `c:logBase` attribute on `c:scaling`
- [ ] Axis crossing at specific value — `c:crossesAt` (OXML exists, needs API)
- [ ] Category axis label offset — `c:lblOffset` value exposure
- [ ] Date axis base time unit — `c:baseTimeUnit`, `c:majorTimeUnit`, `c:minorTimeUnit`

---

## 19. Python API — 3D Shape Formatting

_Prerequisites: Section 4.8 (3D elements)_

- [ ] Shape extrusion — `a:sp3d` (extrusion height, contour width, material)
- [ ] Bevel — top and bottom bevel profiles (`a:bevelT`, `a:bevelB`)
- [ ] 3D scene — camera preset, rotation, field of view (`a:scene3d` -> `a:camera`)
- [ ] Lighting rig — type and direction (`a:scene3d` -> `a:lightRig`)

---

## 20. Python API — Theme Access

_Prerequisites: Section 4.9 (theme elements)_

- [ ] Read/write theme color schemes — `a:clrScheme` in `theme.xml`
- [ ] Read/write theme font schemes — major and minor font families
- [ ] Theme effect schemes

---

## 21. Python API — WordArt & Text Effects

_Prerequisites: Section 4.10 (WordArt elements)_

- [ ] Preset text warp — `a:prstTxWarp` on `a:bodyPr`
- [ ] Text-level 3D scene and fill/outline

---

## 22. Python API — Callout Shapes

- [ ] Callout-specific formatting via adjustment handles on callout preset geometries (accent bar, angle, length, gap)

---

## 23. Python API — Slide Import & Cross-Presentation Operations

- [ ] Import slides from another `.pptx` — clone parts, remap relationships, deduplicate shared resources
- [ ] Merge presentations — combine slide decks with master/layout reconciliation

---

## 24. Python API — Bulk / Range Operations

- [ ] `ShapeRange`-like API for operating on multiple shapes at once (align, distribute, format)
- [ ] Batch shape property updates without repeated XML tree walks

---

## Out of Scope

The following are **not planned** as they require capabilities beyond XML file manipulation:

- Rendering engine (slide-to-image, PDF export, slideshow playback)
- Print subsystem
- Application/window/UI state management
- Clipboard operations
- VBA macro execution
- Image reprocessing/compression
- Animation and transition authoring
- Visual line-break calculation (`TextRange.Lines()`)
