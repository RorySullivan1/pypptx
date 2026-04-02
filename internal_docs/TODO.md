# pypptx TODO

Planned additions and improvements for the pypptx library, organized by priority and area.

---

## 1. Core Reorganization & Structural Improvements

These items address the internal architecture and code quality of the existing codebase before adding new features.

### 1.1 Project Packaging
- [ ] Add `pyproject.toml` with project metadata, dependencies, and build configuration
- [ ] Define explicit dependency on `lxml` and `typing_extensions`
- [ ] Add optional dependency group for text layout features (Pillow, font metrics)
- [ ] Configure pytest in pyproject.toml

### 1.2 Exception Hierarchy
- [ ] Audit exception usage across codebase — many operations silently return `None` instead of raising
- [ ] Add domain-specific exceptions (e.g., `ShapeNotFoundError`, `InvalidOperationError`, `SlideNotFoundError`)
- [ ] Remove backward-compatibility alias for `PythonPptxError` once stable

### 1.3 Shape Tree & Lookup Performance
- [ ] Build name-indexed dictionary on `_BaseShapes` for O(1) shape-by-name lookup
- [ ] Build ID-indexed dictionary for O(1) shape-by-ID lookup
- [ ] Add `Shapes.__contains__` and `Shapes.get(name, default)` for dict-like access
- [ ] Invalidate indexes on mutation (add/remove shape)

### 1.4 Proxy & Base Class Cleanup
- [ ] Review `ElementProxy` / `ParentedElementProxy` / `PartElementProxy` hierarchy for clarity
- [ ] Evaluate whether `Subshape` pattern (in text module) should be unified with `ParentedElementProxy`
- [ ] Remove `object` as explicit base class (Python 2 artifact) across all classes
- [ ] Audit `del` statements in `__init__.py` — evaluate whether namespace cleanup is necessary

### 1.5 Part Loading & Memory
- [ ] Evaluate lazy loading of image/media blobs (defer read until `.blob` accessed)
- [ ] Profile memory usage on large presentations with many embedded images
- [ ] Consider streaming writes for large media parts during save

### 1.6 Type Annotations
- [ ] Audit and complete type annotations across all public APIs
- [ ] Add `py.typed` marker for PEP 561 compliance
- [ ] Ensure all return types are annotated (many properties return untyped)

### 1.7 Test Infrastructure
- [ ] Add pytest configuration and test runner setup
- [ ] Audit test coverage — identify untested public API surface
- [ ] Add integration tests that round-trip (create -> save -> reopen -> verify)
- [ ] Standardize test fixtures — reduce duplication across test modules

### 1.8 OXML Layer
- [ ] Audit `xmlchemy.py` base classes for consistency and documentation
- [ ] Review `simpletypes.py` for completeness against ECMA-376 simple types used
- [ ] Ensure namespace declarations in `ns.py` cover all namespaces needed for planned features

---

## 2. Slide Lifecycle

- [ ] Delete a slide — remove `sldId` from `sldIdLst`, delete slide part and relationships
- [ ] Reorder slides — reorder `sldId` entries within `sldIdLst`, rename parts
- [ ] Duplicate a slide — deep-clone slide part, remap relationships for images/charts/media
- [ ] Slide number property — compute from position in `sldIdLst` + `firstSlideNum`
- [ ] Import slides from another `.pptx` file — clone parts, merge masters/layouts, deduplicate images
- [ ] `Presentation.first_slide_number` — expose `firstSlideNum` attribute

---

## 3. Shape Lifecycle

- [ ] Delete a shape — remove element from `spTree`, clean up related relationships (images, hyperlinks)
- [ ] Duplicate a shape — clone element in `spTree`, assign new shape ID, clone related parts
- [ ] Z-order control — `bring_to_front()`, `send_to_back()`, `bring_forward()`, `send_backward()` via element reordering in `spTree`
- [ ] Shape visibility — expose `hidden` attribute on `cNvPr` as `shape.visible` property
- [ ] Lock aspect ratio — expose `noChangeAspect` on shape lock elements
- [ ] Parent group reference — back-reference from child shape to containing `GroupShape`

---

## 4. Accessibility

- [ ] `shape.alternative_text` — read/write `descr` attribute on `cNvPr`
- [ ] `shape.title` — read/write `title` attribute on `cNvPr`
- [ ] `shape.decorative` — read/write decorative flag (extension element on `cNvPr`)

---

## 5. Text & Font Enhancements

### 5.1 Font Properties
- [ ] Strikethrough — `strike` attribute on `a:rPr` (single, double, none)
- [ ] Superscript / subscript — `baseline` attribute on `a:rPr`
- [ ] Character caps — `cap` attribute on `a:rPr` (none, all, small)
- [ ] Character spacing — `spc` attribute on `a:rPr`
- [ ] Kerning — `kern` attribute on `a:rPr`
- [ ] Font shadow — effect list child on `a:rPr`
- [ ] East Asian font name — `a:ea` element on `a:rPr`
- [ ] Complex script font name — `a:cs` element on `a:rPr`

### 5.2 Paragraph Properties
- [ ] Bullet formatting — `a:buChar`, `a:buAutoNum`, `a:buFont`, `a:buClr`, `a:buSzPct`, `a:buSzPts`, `a:buNone`
- [ ] First-line indent — `indent` attribute on `a:pPr`
- [ ] Left margin — `marL` attribute on `a:pPr`
- [ ] Tab stops — `a:tabLst` with `a:tab` children (position, alignment)
- [ ] Text direction / RTL — `rtl` attribute on `a:pPr`
- [ ] Hanging punctuation — `hangingPunct` attribute on `a:pPr`
- [ ] Baseline alignment — `fontAlgn` attribute on `a:pPr`

### 5.3 Text Frame Properties
- [ ] Orientation — `vert` attribute on `a:bodyPr` (horizontal, vertical, stacked, etc.)
- [ ] Text columns — `numCol` and `spcCol` attributes on `a:bodyPr`
- [ ] `has_text` property — boolean check for non-empty text content

---

## 6. Table Improvements

- [ ] Cell borders — per-edge control via `a:lnL`, `a:lnR`, `a:lnT`, `a:lnB` within `a:tcPr` (color, weight, dash style per edge)
- [ ] Diagonal borders — `a:lnTlToBr`, `a:lnBlToTr` within `a:tcPr`
- [ ] Table style — get/set built-in style GUID via `tblStyle` attribute on `a:tblPr`

---

## 7. Line & Connector Enhancements

- [ ] Arrowhead formatting — `a:headEnd` and `a:tailEnd` within `a:ln` (type, width, length)
- [ ] Compound line style — `cmpd` attribute on `a:ln` (single, double, thick-thin, etc.)
- [ ] Line transparency — alpha modifier on line fill color
- [ ] Line visibility — no-fill vs filled state
- [ ] Line pattern — pattern fill on lines
- [ ] Query connected shapes — read `a:stCxn` / `a:endCxn` attributes on connectors to determine which shapes are connected and at which connection point

---

## 8. Shadow (Full Implementation)

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

## 9. Additional Visual Effects

- [ ] Reflection — `a:reflection` in `a:effectLst` (blur, start/end alpha, distance, direction, scale)
- [ ] Glow — `a:glow` in `a:effectLst` (radius, color)
- [ ] Soft edge — `a:softEdge` in `a:effectLst` (radius)

---

## 10. Headers & Footers

- [ ] Slide-level header/footer configuration — `p:hf` element
- [ ] Date/time placeholder — automatic vs fixed, format string
- [ ] Footer text placeholder
- [ ] Slide number placeholder
- [ ] Per-slide show/hide overrides
- [ ] Notes and handout header/footer support

---

## 11. Metadata & Organization

- [ ] Tags — key-value string pairs on shapes and slides (separate `tags[N].xml` parts linked via relationships)
- [ ] Sections — named slide groups via `p14:sectionLst` in presentation extensions (add, remove, rename, reorder, list sections)
- [ ] Custom document properties — beyond core properties (custom key-value metadata)

---

## 12. Comments

- [ ] Slide comments — add, read, delete comments (`comments[N].xml` parts)
- [ ] Comment authors — manage author list (`commentAuthors.xml` part)
- [ ] Comment positioning — x/y coordinates on slide
- [ ] Comment metadata — author, datetime, text

---

## 13. Picture Format Enhancements

- [ ] Brightness — `a:lum` bright attribute on blip
- [ ] Contrast — `a:lum` contrast attribute on blip
- [ ] Grayscale / black-and-white / washout — `a:grayscl`, `a:duotone` effects on blip
- [ ] Transparency color
- [ ] Original image dimensions — expose from `ImagePart`

---

## 14. Chart Enhancements

- [ ] Plot area — position (x, y, width, height) and formatting
- [ ] Chart area — formatting (fill, line)
- [ ] Display blanks as — `c:dispBlanksAs` (gap, zero, span)
- [ ] Secondary value and category axes — second axis pair with `axId` cross-references
- [ ] 3D chart view — rotation, elevation, perspective attributes on `c:view3D`
- [ ] 3D chart surfaces — floor, walls, back wall, side wall formatting

---

## 15. 3D Shape Formatting

- [ ] Shape extrusion — `a:sp3d` (extrusion height, contour width, material)
- [ ] Bevel — top and bottom bevel profiles (`a:bevelT`, `a:bevelB`)
- [ ] 3D scene — camera preset, rotation, field of view (`a:scene3d` -> `a:camera`)
- [ ] Lighting rig — type and direction (`a:scene3d` -> `a:lightRig`)

---

## 16. Theme Access

- [ ] Read/write theme color schemes — `a:clrScheme` in `theme.xml`
- [ ] Read/write theme font schemes — major and minor font families
- [ ] Theme effect schemes

---

## 17. WordArt & Text Effects

- [ ] Preset text warp — `a:prstTxWarp` on `a:bodyPr`
- [ ] Text-level 3D scene and fill/outline

---

## 18. Callout Shapes

- [ ] Callout-specific formatting via adjustment handles on callout preset geometries (accent bar, angle, length, gap)

---

## 19. Slide Import & Cross-Presentation Operations

- [ ] Import slides from another `.pptx` — clone parts, remap relationships, deduplicate shared resources
- [ ] Merge presentations — combine slide decks with master/layout reconciliation

---

## 20. Bulk / Range Operations

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
