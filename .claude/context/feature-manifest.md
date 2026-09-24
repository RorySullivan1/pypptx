# pypptx Project Manifest

## Purpose

pypptx is a Python library for creating, reading, and modifying Microsoft PowerPoint (.pptx) presentations. It operates directly on the Office Open XML (OOXML/ECMA-376) file format — a ZIP archive of XML parts — without requiring PowerPoint or any rendering engine.

The library provides a Pythonic object model that mirrors the structure of the VBA PowerPoint object model where feasible within the constraints of static file manipulation.

## Version

0.1.0

---

## Scope

### What pypptx Is

A file-format library. It reads and writes the XML inside `.pptx` files. Every feature it offers must be expressible as a modification to the XML parts and relationships that compose the OOXML package.

### What pypptx Is Not

- Not a rendering engine. It cannot convert slides to images or PDF.
- Not a runtime environment. It cannot play slideshows, execute macros, interact with a clipboard, or manage application windows.
- Not a print subsystem. It cannot send presentations to a printer.

---

## Architectural Layers

The codebase is organized in four layers, from low-level to high-level:

| Layer | Packages | Role |
|---|---|---|
| OPC | `pptx.opc` | Open Packaging Convention — ZIP I/O, content types, relationships |
| OXML | `pptx.oxml` | lxml-based element classes mapped to ECMA-376 XML schema types |
| Parts | `pptx.parts` | Package parts — slide, image, chart, media, core properties |
| API | `pptx.api`, `pptx.presentation`, `pptx.slide`, `pptx.shapes`, `pptx.text`, `pptx.chart`, `pptx.dml`, `pptx.table`, `pptx.action` | User-facing Pythonic object model |

---

## Implemented Features

### Presentation Lifecycle
- Create blank presentations from the built-in template
- Open existing `.pptx`, `.pptm` (macro-enabled), and `.potx` (template) files
- Save to file path or file-like stream
- Core document properties (title, author, subject, etc.)
- Custom document properties — typed values (str/int/float/bool) via `Presentation.custom_properties`
- First-slide-number — `Presentation.first_slide_number` r/w (`firstSlideNum`)

### Slide Management
- Add slides from existing slide layouts
- Delete slides — `Slides.delete(slide)` with relationship cleanup
- Duplicate slides — `Slides.duplicate(slide)` deep-clones the part and remaps relationships (images, charts, media)
- Reorder slides — `Slides.move(old_idx, new_idx)`
- Import slides from another presentation — `Slides.import_slide(slide)` with resource deduplication
- Merge presentations — `Slides.merge(presentation)` preserving order, with layout matching and automatic master/theme import
- Access slides by index or slide ID
- Computed slide number — `Slide.slide_number` from position + `firstSlideNum`
- Slide background and background fill

### Slide Hierarchy
- Slide masters and slide layouts (read, iterate, remove layouts)
- Notes master and notes slides (create, read, access notes text)
- Layout-to-master and slide-to-layout relationships
- Placeholder inheritance chain (master → layout → slide)
- Headers & footers — `header_footer` on slide, layout, master, and notes master; per-slide show/hide of date, footer, slide number; date format (read-only field-type string) and auto-vs-fixed mode; footer text r/w

### Sections, Tags & Comments
- Sections — `Presentation.sections` with add, remove, rename, iterate, and per-section slide-ID enumeration (`p14:sectionLst`)
- Tags — dict-like API on slides (and shapes) via `TagsPart` (get/set/del/contains/iter/items)
- Slide comments — `Slide.comments` with add/iterate/clear/indexed access; per-comment `author`, `text` (r/w), `datetime`, `position`, `delete()`
- Comment authors — auto-managed via package-level `CommentAuthorsPart`

### Shapes
- **AutoShapes** — 180+ preset geometries (`MSO_SHAPE`), adjustment handles, text, fill, line
- **Pictures** — insert from file or stream, crop (all four edges), line formatting
- **Tables** — create with rows/columns, cell access, merge/split cells, cell margins, vertical anchor, fill per cell, banding properties (first/last row/col, horizontal/vertical banding), table style (`tblStyle`), per-edge + diagonal borders with full `LineFormat`
- **Charts** — see Charts section below
- **Connectors** — straight/elbow/curve types, begin/end positioning, connect to shape connection points; queryable connected shapes via `Connector.begin_connection`/`end_connection`
- **Group shapes** — group existing shapes, access child shapes
- **Freeform shapes** — programmatic path construction with line segments
- **Graphic frames** — charts, tables, and OLE objects
- **Placeholders** — typed placeholders (title, body, picture, chart, table) with insert operations
- **Movies** — insert video with poster frame
- **Callout shapes** — `MSO_SHAPE.LINE_CALLOUT_*` presets with adjustment handles
- **Bulk operations** — `ShapeRange` API for alignment, distribution, and batch property updates without repeated XML walks

### Shape Properties (Common)
- Position (left, top) and size (width, height) in EMUs
- Rotation
- Name (read/write) and shape ID (read-only)
- Shape type enumeration
- Click action and hyperlink
- Placeholder format (index, type)
- Visibility — `BaseShape.hidden`
- Lock aspect ratio — `BaseShape.lock_aspect_ratio`
- Group containment — `BaseShape.is_in_group`, `BaseShape.parent_group`
- Z-order — `move_shape_to_front()`, `move_shape_to_back()`
- Accessibility — `alternative_text`, `title`, `decorative`

### Text
- Text frames with auto-size control (none, shrink text, resize shape)
- `has_text` boolean
- Vertical anchor, word wrap, four-edge margins
- Text orientation (`vert`) and columns (`numCol`, `spcCol`)
- Paragraphs — alignment, level, line spacing, space before/after, first-line indent (`indent`), left margin (`marL`), RTL, hanging punctuation, baseline / font alignment (`fontAlgn`)
- Tab stops — `paragraph.tab_stops`, `add_tab_stop`, `clear_tab_stops`
- Bullet formatting — character, numbered, picture; bullet font, color, size; no-bullet
- Runs with per-run font formatting
- Font — bold, italic, underline (16 styles), color, size, name, language, fill
- Strikethrough, super/subscript (`baseline`), caps (all/small/none), character spacing (`spc`), kerning (`kern`), font shadow
- East Asian (`a:ea`) and complex-script (`a:cs`) font names
- Hyperlinks on runs
- Fit text to shape (requires font-metrics file)

### Drawing / Formatting
- **Fill** — solid, gradient (angle, stops with color and position), patterned (50+ patterns), background (no fill), foreground/background colors
- **Line** — width, color, dash style (8 presets), solid/pattern fill, no-fill check, transparency (color alpha), arrowheads (head + tail: type, width, length), compound type (`MSO_LINE_COMPOUND_TYPE`: single, double, thick-thin, etc.)
- **Color** — RGB, theme color with brightness adjustment, color-type detection, alpha
- **Shadow** — full read/write via `ShadowFormat`: type (outer/inner/none), blur radius, distance, direction, alignment, color with transparency, rotate-with-shape, visibility
- **Reflection / Glow / Soft Edge** — full read/write via `ReflectionFormat`, `GlowFormat`, `SoftEdgeFormat` on `BaseShape`
- **Chart formatting** — fill and line on chart elements via `ChartFormat` proxy

### 3D Formatting
- Shape extrusion — `ThreeDFormat.extrusion_height`, `contour_width`, `material`
- Bevel — top and bottom profiles via `Bevel.width`, `height`, `preset`
- Scene — `Scene3D.camera` with `Camera.preset`, `field_of_view`
- Lighting rig — `Scene3D.light_rig` with `LightRig.rig_type`, `direction`

### Charts
- 16 chart families (75+ enumerated `XL_CHART_TYPE` values) — bar, line, pie, scatter, bubble, area, radar, stock, surface, doughnut, of-pie, plus 3D variants (bar3D, line3D, pie3D, area3D, surface3D)
- Series formatting — markers, fill, line; pie/bar/bubble per-series properties (explosion, bar shape, bubble-3D)
- Axes — value, category, date; primary + secondary; log scale; tick / label skip; label offset and rotation; display units; date axis time units; axis crossing
- Legends — `Legend.format` (`ChartFormat` for box formatting); per-entry overrides
- Data labels — collection and per-point: position, separator, leader lines, show val / cat / ser / percent / bubble size toggles
- Trendlines — `XL_TRENDLINE_TYPE` (linear, exponential, logarithmic, moving-average, polynomial, power), order/period, forward/backward, intercept, R² and equation display
- Error bars — type (custom, fixed value, percent, stddev, stderr), direction (X/Y), include (both/plus/minus), value, end-cap
- 3D view — rotation X/Y, right-angle axes, perspective, depth percent, height percent
- 3D surfaces — `Chart.floor`, `Chart.back_wall`, `Chart.side_wall` with format and thickness
- Plot area — `PlotArea` layout (left/top/width/height as fractions) and formatting
- Chart area — `Chart.chart_format`, `rounded_corners`, `plot_visible_only`, `show_data_labels_over_max`, `display_blanks_as`
- Drop lines, hi-lo lines, up-down bars, series lines on applicable plot types
- Replace chart data; combo charts via `Chart.add_plot(plot_type, use_secondary_axis, grouping)`
- Pie / doughnut — first-slice angle, hole size; of-pie split type/pos/second-pie-size; gap-width

### Picture Format
- Brightness — `Picture.brightness` (`a:lum` bright)
- Contrast — `Picture.contrast` (`a:lum` contrast)
- Grayscale — `Picture.is_grayscale` (`a:grayscl`)
- Transparency color — `Picture.transparency_color` r/w (`a:clrChange`)
- Original image dimensions — `Picture.image_width` / `image_height` via `ImagePart`

### Theme
- Color schemes — read/write `a:clrScheme`
- Font schemes — major and minor font families
- Effect schemes — `Theme.effect_scheme` returning `EffectScheme` with indexed/named access to subtle/moderate/intense `EffectStyle` objects

### Enumerations
- 180+ auto shape types (`MSO_SHAPE`)
- 75+ chart types (`XL_CHART_TYPE`)
- 137 language identifiers
- Shape types, placeholder types, connector types, media types
- Fill types, color types, line dash styles, pattern types, theme colors
- Line compound types (`MSO_LINE_COMPOUND_TYPE`)
- Text alignment, auto size, underline types, vertical anchor, font caps, font baseline
- Chart axis types, data label positions, legend positions, marker styles, tick marks
- Trendline types, error-bar type / direction / include
- Action types

### OLE Objects
- Embed DOCX, PPTX, XLSX with icon representation
- Access OLE blob and prog ID
- Custom OLE objects with arbitrary files

### Utilities
- Length units: EMU, Inches, Cm, Mm, Pt, Centipoints (all interconvertible)
- `@lazyproperty` cache decorator for derived collections

---

## Known Limitations

These are inherent constraints of a file-format library that edits XML. They cannot be resolved within the scope of this project.

### Requires a Rendering Engine (Out of Scope)
- Export slides to images (PNG, JPEG, etc.)
- Export presentation to PDF/XPS
- Slideshow playback
- Print support
- Text reflow calculation (knowing where line breaks fall within a paragraph)
- Compress/reprocess embedded images

### Requires Application Runtime (Out of Scope)
- Application object and window management
- Clipboard operations (copy, cut, paste shapes/slides between sessions)
- UI selection state (selecting shapes, slides, text ranges)
- Format painter (pick up / apply formatting from UI context)
- Default shape properties (application-level state)
- VBA macro execution (macros are preserved in the file but cannot be run or edited)
- AddIns, CommandBars, FileDialog

### Decorative Effects (Out of Scope)
- WordArt / preset text warp (`a:prstTxWarp` on `a:bodyPr`) — OXML layer exists (`CT_PresetTextShape`) for direct access if needed, but no high-level API is planned
- Text-level 3D scene, fill, and outline effects

### Text Range Model Limitations
- VBA's `TextRange` provides `Characters()`, `Words()`, `Sentences()`, `Lines()`, `Find()`, `Replace()`, `InsertBefore()`, `InsertAfter()`. The `Lines()` method requires a rendering engine to know where visual line breaks occur. The string-search operations are theoretically possible but complicated by the run-splitting model where a single visual word may span multiple XML runs — planned for a future milestone (see `.meta/roadmap/stages/03-hardening-and-gaps/v0.6.0.md`).

---

## Feature Gaps

These are features that **are representable in the OOXML schema** and are therefore implementable, but are not yet present in the library.

### Custom Slide Layouts
- Create new slide layout parts and link to a slide master
- Define placeholder positions and types on custom layouts

### Transitions & Animations (read-only target — see `.meta/roadmap/stages/03-hardening-and-gaps/v0.4.0.md`)
- Slide transition timing and effect parameters (`p:transition`)
- Read access to the animation timing tree on a slide

### Cross-run Text Operations (see `.meta/roadmap/stages/03-hardening-and-gaps/v0.6.0.md`)
- Find / replace across runs within a paragraph
- Run-aware string manipulation (`InsertBefore`, `InsertAfter`, `Characters`, `Words`)
