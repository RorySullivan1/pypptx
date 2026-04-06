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
- [x] Incremental annotation improvements — return types, parameter types, and `-> None` on setters across 15+ key API files

### 1.7 Test Infrastructure
- [x] pytest configured in `pyproject.toml` (test paths, class/function patterns)
- [x] 2644 tests passing, 97% code coverage
- [x] Add round-trip integration tests — 8 tests covering blank pres, shapes, accessibility, fonts, slide ops, duplication, hidden shapes, slide numbers

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
- [x] Decorative flag — implemented via manual `a:ext` element handling (see Section 8)

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
- [x] Shadow style enumeration — handled by `ShadowFormat.shadow_type` returning `"outer"`/`"inner"`/`None`

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

## 5.2 OXML Gap Remediation (DONE)

Close remaining OXML descriptor gaps discovered during audit before
building Python API layers (Sections 6-17).

### Line end elements (DONE)
- [x] `CT_LineEndProperties` class with `type`, `w`, `len` attributes
- [x] `headEnd` and `tailEnd` ZeroOrOne descriptors on `CT_LineProperties`
- [x] Registered `a:headEnd` and `a:tailEnd`

### Effect color children (DONE)
- [x] `ZeroOrOneChoice` for `eg_colorChoice` on `CT_OuterShadowEffect`, `CT_InnerShadowEffect`, `CT_GlowEffect`

### Blip effect chain (DONE)
- [x] `duotone`, `grayscl`, `lum` ZeroOrOne descriptors on `CT_Blip`
- [x] `rLink` attribute on `CT_Blip`

### CT_DuotoneEffect (DONE)
- [x] `color_elms` property for accessing the two color children

### Tags OXML (DONE)
- [x] `CT_TagList` and `CT_StringTag` in new `oxml/tags.py`
- [x] Registered `p:tag` and `p:tagLst`

### CT_Presentation cleanup (DONE)
- [x] Full `_tag_seq` with proper child ordering
- [x] `notesMasterIdLst` and `handoutMasterIdLst` ZeroOrOne descriptors

### Color element attributes (DONE)
- [x] `CT_HslColor` — `hue`, `sat`, `lum` required attributes
- [x] `CT_PresetColor` — `val` required attribute
- [x] `CT_ScRgbColor` — `r`, `g`, `b` required attributes
- [x] `CT_SystemColor` — `val` required, `lastClr` optional attributes

---

## 5.3 OXML Final Micro-Gaps (DONE)

### Text character properties (DONE)
- [x] `effectLst` ZeroOrOne descriptor on `CT_TextCharacterProperties` (font shadow support)

### Tab stops (DONE)
- [x] `tabLst` ZeroOrOne descriptor on `CT_TextParagraphProperties`
- [x] `CT_TabStopList` class with `tab = ZeroOrMore("a:tab")`
- [x] `CT_TabStop` class with `pos` and `algn` attributes
- [x] Registered `a:tab` and `a:tabLst`

---

## 6. Python API — Slide Lifecycle (DONE)

_Prerequisites: Section 3.7 (`firstSlideNum` attribute)_

- [x] Delete a slide — `Slides.delete(slide)` removes sldId and drops relationship
- [x] Reorder slides — `Slides.move(old_idx, new_idx)` reorders sldId entries
- [x] Duplicate a slide — `Slides.duplicate(slide)` deep-clones slide part, remaps relationships for images/charts/media
- [x] Slide number property — `Slide.slide_number` computed from position + firstSlideNum
- [x] `Presentation.first_slide_number` — read/write firstSlideNum attribute

---

## 7. Python API — Shape Lifecycle (DONE)

_Prerequisites: Sections 3.1 (`hidden`), 3.2 (`noChangeAspect`)_

- [x] Delete a shape — `_BaseGroupShapes.remove_shape(shape)` removes element from spTree
- [x] Duplicate a shape — `_BaseGroupShapes.duplicate_shape(shape)` deep-clones element, assigns new ID/name, remaps rId references
- [x] Z-order control — `move_shape_to_front()`, `move_shape_to_back()` via element reordering
- [x] Shape visibility — `BaseShape.hidden` read/write property on cNvPr
- [x] Lock aspect ratio — `BaseShape.lock_aspect_ratio` r/w property via lock elements on cNv*Pr
- [x] Parent group reference — `BaseShape.is_in_group` and `BaseShape.parent_group` for group containment

---

## 8. Python API — Accessibility (DONE)

_Prerequisites: Section 3.1 (`descr`, `title`, decorative flag)_

- [x] `shape.alternative_text` — read/write `descr` attribute on `cNvPr`
- [x] `shape.title` — read/write `title` attribute on `cNvPr`
- [x] `shape.decorative` — read/write decorative flag (extension element on `cNvPr`)

---

## 9. Python API — Text & Font Enhancements

_Prerequisites: Sections 3.3, 3.4, 3.5 (text attributes), Section 5 (enumerations)_

### 9.1 Font Properties (DONE)
- [x] Strikethrough — `font.strikethrough` via `strike` attribute on `a:rPr`
- [x] Superscript / subscript — `font.baseline` via `baseline` attribute on `a:rPr`
- [x] Character caps — `font.caps` via `cap` attribute on `a:rPr`
- [x] Character spacing — `font.spacing` via `spc` attribute on `a:rPr`
- [x] Kerning — `font.kerning` via `kern` attribute on `a:rPr`
- [x] Font shadow — `Font.has_shadow` r/w via `a:effectLst/a:outerShdw` on `a:rPr`
- [x] East Asian font name — `font.east_asian_name` via `a:ea` element on `a:rPr`
- [x] Complex script font name — `font.complex_script_name` via `a:cs` element on `a:rPr`

### 9.2 Paragraph Properties (DONE)
- [x] Bullet formatting — `paragraph.bullet_char`, `paragraph.bullet_type`
- [x] First-line indent — `paragraph.indent` via `indent` attribute on `a:pPr`
- [x] Left margin — `paragraph.margin_left` via `marL` attribute on `a:pPr`
- [x] Tab stops — `paragraph.tab_stops`, `add_tab_stop()`, `clear_tab_stops()`
- [x] Text direction / RTL — `paragraph.rtl` via `rtl` attribute on `a:pPr`
- [x] Hanging punctuation — `paragraph.hanging_punctuation` via `hangingPunct` on `a:pPr`
- [x] Baseline alignment — `paragraph.font_alignment` via `fontAlgn` on `a:pPr`

### 9.3 Text Frame Properties (DONE)
- [x] Orientation — `text_frame.text_orientation` via `vert` attribute on `a:bodyPr`
- [x] Text columns — `text_frame.columns` via `numCol`, `text_frame.column_spacing` via `spcCol`
- [x] `has_text` property — boolean check for non-empty text content

---

## 10. Python API — Table Improvements (DONE)

_Prerequisites: Section 4.3 (table cell border elements)_

- [x] Cell borders — `_Cell.border_left/right/top/bottom` via `_CellBorderAdapter` + `LineFormat`
- [x] Diagonal borders — `_Cell.border_diagonal_down` / `border_diagonal_up` via `a:lnTlToBr`, `a:lnBlToTr`
- [x] Table style — `Table.table_style_id` via `tblStyle` attribute on `a:tblPr`

---

## 11. Python API — Line & Connector Enhancements (DONE)

_Prerequisites: Section 3.6 (`cmpd` attribute)_

- [x] Arrowhead formatting — `begin/end_arrowhead_type/width/length` on `LineFormat`
- [x] Compound line style — `LineFormat.compound_type` via `cmpd` attribute on `a:ln`
- [x] Line transparency — `line.color.alpha` via `ColorFormat.alpha` on the line's solid fill color
- [x] Line visibility — `LineFormat.no_fill` read-only check
- [x] Line pattern — `line.fill.patterned()` + `line.fill.pattern`/`fore_color`/`back_color` via `FillFormat`
- [x] Query connected shapes — `Connector.begin_connection`/`end_connection`

---

## 12. Python API — Shadow (Full Implementation) (DONE)

_Prerequisites: Section 4.1 (shadow elements)_

- [x] Shadow type — `ShadowFormat.shadow_type` (outer/inner/none)
- [x] Blur radius — `ShadowFormat.blur_radius`
- [x] Distance and direction — `ShadowFormat.distance`, `ShadowFormat.direction`
- [x] Shadow color with transparency — `ShadowFormat.color` returns `ColorFormat`; alpha via `color.alpha` (0.0–1.0)
- [x] Alignment — `ShadowFormat.alignment`
- [x] Rotate with shape — `ShadowFormat.rotate_with_shape`
- [x] Visibility — presence/absence of shadow element

---

## 13. Python API — Additional Visual Effects (DONE)

_Prerequisites: Section 4.2 (effect elements)_

- [x] Reflection — `ReflectionFormat` with enabled, blur_radius, start/end_opacity, distance, clear
- [x] Glow — `GlowFormat` with enabled, radius, clear
- [x] Soft edge — `SoftEdgeFormat` with enabled, radius, clear
- [x] Exposed on `BaseShape` as lazyproperty accessors

---

## 14. Python API — Headers & Footers (DONE)

_Prerequisites: Section 4.4 (`CT_HeaderFooter`)_

- [x] Header/footer configuration — `HeaderFooter` class with show_date, show_footer, show_header, show_slide_number
- [x] Exposed on `SlideLayout`, `SlideMaster`, `NotesMaster` as `header_footer` lazyproperty
- [x] Date/time placeholder — `date_format` (read-only field type string), `is_date_auto` (auto vs fixed)
- [x] Footer text placeholder content management — `footer_text` read/write property
- [x] Per-slide show/hide overrides — `p:hf` on `CT_Slide`, `Slide.header_footer` lazyproperty

---

## 15. Python API — Metadata & Organization (DONE)

_Prerequisites: Section 4.5 (sections and tags elements)_

- [x] Tags — `Slide.tags` via `TagsPart` with dict-like API (get/set/del/contains/iter/items)
- [x] Sections — `Presentation.sections` with `Sections`/`Section` classes (add, remove, rename, iterate, slide IDs)
- [x] Custom document properties — `Presentation.custom_properties` via `CustomPropertiesPart` with typed values (str/int/float/bool)

---

## 16. Python API — Comments (DONE)

_Prerequisites: Section 4.6 (comment elements)_

- [x] Slide comments — `Slide.comments` returns `SlideComments` with add, iterate, len, clear, indexed access
- [x] Comment authors — `CommentAuthorsPart` on package with get_or_add_author, auto-incrementing IDs
- [x] Comment positioning — `Comment.position` returns (x, y) tuple
- [x] Comment metadata — `Comment.author`, `Comment.text` (r/w), `Comment.datetime`, `Comment.delete()`
- [x] Part registration — `CommentsPart` and `CommentAuthorsPart` registered in PartFactory

---

## 17. Python API — Picture Format Enhancements (DONE)

_Prerequisites: Section 4.7 (picture format effects)_

- [x] Brightness — `Picture.brightness` via `a:lum` bright attribute on blip
- [x] Contrast — `Picture.contrast` via `a:lum` contrast attribute on blip
- [x] Grayscale — `Picture.is_grayscale` via `a:grayscl` on blip
- [x] Transparency color — `Picture.transparency_color` read/write via `a:clrChange` on blip
- [x] Original image dimensions — `Picture.image_width` / `image_height` via `ImagePart._native_size`

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
- [x] Axis `axId` / `crossAx` cross-referencing — implemented in 18.5 via `Plot.axis_ids` and `add_secondary_axes()`
- [x] `Plot.axis_ids` property — implemented in 18.5
- [x] Combo chart axis assignment — implemented in 18.5 via `Chart.add_plot()` with `use_secondary_axis`

### 18.5 Combo Charts — API (DONE)

- [x] `Plot.axis_ids` property — tuple of integer axis IDs the plot references
- [x] `BaseChartElement.axId_vals` / `axId_lst` — OXML axis ID accessors
- [x] Register `c:axId` as `CT_UnsignedInt`
- [x] `CT_PlotArea.add_secondary_axes()` — creates hidden secondary catAx + right-positioned valAx
- [x] `CT_PlotArea.add_xChart()` — adds xChart element with axis references
- [x] `Chart.add_plot(plot_type, use_secondary_axis, grouping)` — high-level API for overlay plots
- [x] Supports "line", "bar", "area", "scatter" overlay types
- [x] `Chart.plots` iteration verified working with multiple xChart elements

### 18.6 3D Chart Properties (DONE)

- [x] `CT_View3D` — `rotX`, `rotY`, `rAngAx`, `perspective`, `depthPercent`, `hPercent` children with val properties
- [x] `CT_RotX`, `CT_RotY`, `CT_Perspective`, `CT_DepthPercent`, `CT_HPercent` element classes
- [x] `CT_Surface` — shared class for `c:floor`, `c:sideWall`, `c:backWall` with `thickness` and `spPr`
- [x] Declared `c:view3D`, `c:floor`, `c:sideWall`, `c:backWall` as `ZeroOrOne` on `CT_Chart`
- [x] `Chart.view_3d` API — `View3D` with rot_x, rot_y, right_angle_axes, perspective, depth_percent, height_percent
- [x] `Chart.floor`, `Chart.back_wall`, `Chart.side_wall` — `ChartSurface` with format and thickness

### 18.7 Chart & Plot Area Formatting (DONE)

- [x] `CT_ManualLayout` — added `y`, `w`, `h`, `yMode`, `wMode`, `hMode` descriptors + left/top/width/height properties
- [x] Plot area layout — `PlotArea` API with left, top, width, height (fraction of chart dimensions)
- [x] Plot area formatting — `PlotArea.format` via `c:spPr` on `c:plotArea`
- [x] Chart area formatting — `Chart.chart_format` via `c:spPr` on `c:chartSpace`
- [x] `Chart.display_blanks_as` — read/write "gap"/"zero"/"span" for missing data points
- [x] `Chart.plot_area` property exposing `PlotArea` object
- [x] Registered `c:h`, `c:hMode`, `c:w`, `c:wMode`, `c:y`, `c:yMode`, `c:dispBlanksAs`

### 18.8 Line/Area/Bar Chart Special Elements (DONE)

- [x] `c:dropLines` — registered as CT_ChartLines, descriptors on AreaChart, LineChart, Line3DChart, StockChart
- [x] `c:hiLowLines` — registered as CT_ChartLines, descriptors on LineChart, StockChart
- [x] `c:upDownBars` — CT_UpDownBars with gapWidth/upBars/downBars, descriptors on LineChart, StockChart
- [x] `c:serLines` — registered as CT_ChartLines, descriptors on BarChart, OfPieChart
- [x] API: DropLines, HiLowLines, SeriesLines (format), UpDownBars (gap_width, up/down bars format)
- [x] has_drop_lines/has_hi_low_lines/has_up_down_bars/has_series_lines toggle properties on plots

### 18.9 OXML Foundation — Plot Type Descriptor Gaps

_OXML descriptors must exist before API properties can be built on them._

Missing descriptors on chart type OXML classes. All items are `ZeroOrOne` unless noted.

- [x] `CT_Area3DChart` — convert to `_tag_seq` pattern; add `varyColors`, `ser` (ZeroOrMore), `dLbls`, `dropLines`, `gapDepth` descriptors
- [x] `CT_Bar3DChart` — add `c:gapDepth`, `c:shape` descriptors
- [x] `CT_BubbleChart` — add `c:varyColors`, `c:showNegBubbles`, `c:sizeRepresents` descriptors
- [x] `CT_DoughnutChart` — add `c:firstSliceAng`, `c:holeSize` descriptors
- [x] `CT_PieChart` — add `c:firstSliceAng` descriptor
- [x] `CT_OfPieChart` — add `c:ofPieType`, `c:gapWidth`, `c:splitType`, `c:splitPos`, `c:custSplit`, `c:secondPieSize` descriptors
- [x] `CT_LineChart` — add `c:marker`, `c:smooth` descriptors
- [x] `CT_Line3DChart` — add `c:gapDepth` descriptor
- [x] `CT_RadarChart` — add `c:radarStyle` descriptor
- [x] `CT_ScatterChart` — add `c:scatterStyle`, `c:dLbls` descriptors
- [x] `CT_Surface3DChart` — add `c:wireframe`, `c:bandFmts` descriptors
- [x] `CT_SurfaceChart` — add `c:wireframe`, `c:bandFmts` descriptors

### 18.10 OXML Foundation — Series, Legend & Chart-Level Descriptor Gaps

- [x] `CT_SeriesComposite` — add `c:explosion` descriptor (pie/doughnut slice explosion)
- [x] `CT_SeriesComposite` — add `c:shape` descriptor (bar shape: box, cone, cylinder, pyramid)
- [x] `CT_SeriesComposite` — add `c:bubble3D` descriptor (3D bubble effect)
- [x] `CT_DPt` — add `c:explosion`, `c:bubble3D` descriptors (per-point overrides)
- [x] `CT_Legend` — add `c:legendEntry` (ZeroOrMore) descriptor (per-entry formatting)
- [x] `CT_Legend` — add `c:spPr` descriptor (legend box formatting)
- [x] `CT_TrendlineLabel` — add `c:layout`, `c:tx`, `c:txPr` descriptors
- [x] `CT_Chart` — add `c:plotVisOnly` descriptor (plot visible cells only)
- [x] `CT_Chart` — add `c:showDLblsOverMax` descriptor (show labels over max)
- [x] `CT_ChartSpace` — add `c:roundedCorners` descriptor (rounded chart border)

### 18.11 OXML Foundation — Data Label Descriptor Gaps

- [x] `CT_DLbls` — add `spPr` descriptor (data label collection formatting)
- [x] `CT_DLbls` — add `separator` descriptor
- [x] `CT_DLbls` — add `showLeaderLines` descriptor
- [x] `CT_DLbls` — add `leaderLines` descriptor
- [x] `CT_DLbls` — add `showBubbleSize` descriptor
- [x] `CT_DLbl` — add `numFmt` descriptor
- [x] `CT_DLbl` — add `showLegendKey`, `showVal`, `showCatName`, `showSerName`, `showPercent`, `showBubbleSize` descriptors
- [x] `CT_DLbl` — add `separator` descriptor
- [x] `CT_DLbl` — add `layout` descriptor (per-point label positioning)

### 18.12 OXML Foundation — Axis Descriptor Gaps

- [x] `CT_Scaling` — add `c:logBase` descriptor (enables logarithmic scale)
- [x] `CT_ValAx` — add `c:dispUnits` descriptor (hundreds, thousands, millions, etc.)
- [x] `CT_ValAx` — add `c:axPos` descriptor (axis position: l, r, t, b)
- [x] `CT_ValAx` — add `c:crossBetween` descriptor (midpoint vs between categories)
- [x] `CT_CatAx` — add `c:crossAx` descriptor (cross-axis reference)
- [x] `CT_CatAx` — add `c:axPos`, `c:auto`, `c:lblAlgn` descriptors
- [x] `CT_CatAx` — add `c:tickLblSkip`, `c:tickMarkSkip`, `c:noMultiLvlLbl` descriptors
- [x] `CT_DateAx` — add `c:crossAx` descriptor
- [x] `CT_DateAx` — add `c:baseTimeUnit`, `c:majorUnit`, `c:majorTimeUnit`, `c:minorUnit`, `c:minorTimeUnit` descriptors

### 18.13 API — Data Label Enhancements

_Prerequisites: 18.11 (data label descriptors)_

- [x] Separator text — `c:separator` element on `c:dLbls`
- [x] Leader line formatting — `c:leaderLines` spPr on `c:dLbls`
- [x] Individual data label override — `c:dLbl` per-point customization (already partially implemented)
- [x] `DataLabels.separator` — read/write separator string property
- [x] `DataLabels.show_leader_lines` — read/write boolean
- [x] `DataLabels.leader_lines` — ChartFormat for leader line formatting
- [x] `DataLabels.show_bubble_size` — read/write boolean
- [x] Per-point data label API — `DataLabel` class with show* overrides, separator, format, position
- [x] Pie/doughnut API — `DoughnutPlot.hole_size`, `PiePlot.first_slice_angle`, `DoughnutPlot.first_slice_angle`
- [x] Of-pie API — `OfPiePlot.split_type`, `OfPiePlot.split_pos`, `OfPiePlot.second_pie_size`, `OfPiePlot.gap_width`

### 18.14 API — Axis Enhancements

_Prerequisites: 18.12 (axis descriptors)_

- [x] Log scale — `ValueAxis.log_scale` read/write float (logBase on scaling)
- [x] Cross between — `ValueAxis.cross_between` read/write string
- [x] Axis crossing at specific value — `ValueAxis.crosses_at` (already existed)
- [x] Category axis tick skip — `CategoryAxis.tick_label_skip`, `CategoryAxis.tick_mark_skip`
- [x] Category axis label offset — `TickLabels.offset` (already existed)
- [x] Date axis time units — `DateAxis.base_time_unit`, `DateAxis.major_time_unit`, `DateAxis.minor_time_unit`
- [x] Chart-level API — `Chart.plot_visible_only`, `Chart.show_data_labels_over_max`, `Chart.rounded_corners`
- [x] Legend API — `Legend.format` (ChartFormat for legend box formatting)
- [x] Axis label rotation — `TickLabels.label_rotation` (rot on a:bodyPr, degrees)
- [x] Display units — `CT_DispUnits`, `CT_BuiltInUnit` + `ValueAxis.display_units`
- [x] Series API — `PieSeries.explosion`, `BarSeries.bar_shape`, `BubbleSeries.bubble_3d`

---

## 19. Python API — 3D Shape Formatting

_Prerequisites: Section 4.8 (3D elements)_

- [x] Shape extrusion — `ThreeDFormat.extrusion_height`, `contour_width`, `material`
- [x] Bevel — `ThreeDFormat.bevel_top`/`bevel_bottom` with `Bevel.width`, `height`, `preset`
- [x] 3D scene — `Scene3D.camera` with `Camera.preset`, `field_of_view`
- [x] Lighting rig — `Scene3D.light_rig` with `LightRig.rig_type`, `direction`

---

## 20. Python API — Theme Access

_Prerequisites: Section 4.9 (theme elements)_

- [x] Read/write theme color schemes — `a:clrScheme` in `theme.xml`
- [x] Read/write theme font schemes — major and minor font families
- [x] Theme effect schemes — `Theme.effect_scheme` returning `EffectScheme` with indexed/named access to 3 `EffectStyle` objects (subtle/moderate/intense)

---

## 21. ~~Python API — WordArt & Text Effects~~ (OUT OF SCOPE)

_Moved to Out of Scope — decorative effects rarely needed in programmatic document generation.
OXML layer support exists (`CT_PresetTextShape` on `CT_TextBodyProperties`) for direct access if needed._

- ~~Preset text warp — `a:prstTxWarp` on `a:bodyPr`~~
- ~~Text-level 3D scene and fill/outline~~

---

## 22. Python API — Callout Shapes (DONE)

- [x] Callout-specific formatting via adjustment handles on callout preset geometries — already supported through `Shape.adjustments` (indexed read/write `AdjustmentCollection`) and `MSO_SHAPE.LINE_CALLOUT_*` presets with default adjustment values in `spec.py`

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
- WordArt / text warp effects (OXML layer exists for direct access via `CT_PresetTextShape`)
- Text-level 3D scene, fill, and outline effects
