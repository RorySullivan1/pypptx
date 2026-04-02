# pypptx Project Manifest

## Purpose

pypptx is a Python library for creating, reading, and modifying Microsoft PowerPoint (.pptx) presentations. It operates directly on the Office Open XML (OOXML/ECMA-376) file format — a ZIP archive of XML parts — without requiring PowerPoint or any rendering engine.

The library provides a Pythonic object model that mirrors the structure of the VBA PowerPoint object model where feasible within the constraints of static file manipulation.

## Version

0.0.01

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
- Create blank presentations from built-in template
- Open existing `.pptx` files (including macro-enabled and template variants)
- Save to file path or file-like stream
- Access and modify core document properties (title, author, subject, etc.)

### Slide Management
- Add slides from existing slide layouts
- Access slides by index or slide ID
- Query slide count and slide position
- Access slide background and background fill

### Slide Hierarchy
- Slide masters and slide layouts (read, iterate, remove layouts)
- Notes master and notes slides (create, read, access notes text)
- Layout-to-master and slide-to-layout relationships
- Placeholder inheritance chain (master -> layout -> slide)

### Shapes
- **AutoShapes** — all preset geometries (180+ shape types), adjustment handles, text, fill, line
- **Pictures** — insert from file or stream, crop (all four edges), line formatting
- **Tables** — create with rows/columns, cell access, merge/split cells, cell margins, vertical anchor, fill per cell, banding properties (first/last row/col, horizontal/vertical banding)
- **Charts** — 75+ chart types, category/XY/bubble data, axes, legends, data labels, series formatting, markers, chart title, gridlines, replace data
- **Connectors** — straight/elbow/curve types, begin/end positioning, connect to shape connection points
- **Group shapes** — group existing shapes, access child shapes
- **Freeform shapes** — programmatic path construction with line segments
- **Graphic frames** — charts, tables, and OLE objects
- **Placeholders** — typed placeholders (title, body, picture, chart, table) with insert operations
- **Movies** — insert video with poster frame

### Shape Properties (Common)
- Position (left, top) and size (width, height) in EMUs
- Rotation
- Name (read/write)
- Shape ID (read-only)
- Shape type enumeration
- Click action and hyperlink
- Placeholder format (index, type)

### Text
- Text frames with auto-size control (none, shrink text, resize shape)
- Paragraphs with alignment, level, line spacing, space before/after
- Runs with per-run font formatting
- Font: bold, italic, underline (16 styles), color, size, name, language, fill
- Hyperlinks on runs
- Fit text to shape (requires font metrics file)
- Margins (top, bottom, left, right)
- Vertical anchor
- Word wrap

### Drawing / Formatting
- **Fill** — solid, gradient (angle, stops with color and position), patterned (50+ patterns), background (no fill), foreground/background colors
- **Line** — width, color, dash style (8 styles), fill
- **Color** — RGB, theme color with brightness adjustment, color type detection
- **Shadow** — inherit flag only (stub implementation)
- **Chart formatting** — fill and line on chart elements via ChartFormat proxy

### Enumerations
- 180+ auto shape types (MSO_SHAPE)
- 75+ chart types (XL_CHART_TYPE)
- 137 language identifiers
- Shape types, placeholder types, connector types, media types
- Fill types, color types, line dash styles, pattern types, theme colors
- Text alignment, auto size, underline types, vertical anchor
- Chart axis types, data label positions, legend positions, marker styles, tick marks
- Action types

### OLE Objects
- Embed DOCX, PPTX, XLSX with icon representation
- Access OLE blob and prog ID
- Custom OLE objects with arbitrary files

### Utilities
- Length units: EMU, Inches, Cm, Mm, Pt, Centipoints (all interconvertible)
- Lazy property decorator for performance

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

### Text Range Model Limitations
- VBA's `TextRange` provides `Characters()`, `Words()`, `Sentences()`, `Lines()`, `Find()`, `Replace()`, `InsertBefore()`, `InsertAfter()`. The `Lines()` method requires a rendering engine to know where visual line breaks occur. The string-search operations are theoretically possible but complicated by the run-splitting model where a single visual word may span multiple XML runs.

---

## Feature Gaps

These are features that **are representable in the OOXML schema** and are therefore implementable, but are not yet present in the library. They are grouped by the area of the object model they belong to.

### Slide Lifecycle
- Delete a slide (remove from `sldIdLst`, delete part and relationships)
- Reorder slides (reorder `sldId` entries in `sldIdLst`)
- Duplicate a slide (deep-clone slide part, remap relationships)
- Import slides from another presentation (clone parts, merge masters/layouts)
- Slide number (computed from position + `firstSlideNum`)

### Shape Lifecycle
- Delete a shape (remove element from `spTree`, clean up relationships)
- Duplicate a shape (clone element, assign new ID, clone related parts)
- Z-order control (reorder elements within `spTree`; document order = z-order)
- Shape visibility (`hidden` attribute on `cNvPr`)
- Lock aspect ratio (`noChangeAspect` on shape locks)
- Parent group reference (Python-side back-reference during tree traversal)

### Accessibility
- Alternative text (`descr` attribute on `cNvPr`)
- Shape title (`title` attribute on `cNvPr`)
- Decorative flag (extension element on `cNvPr`)

### Text & Font
- Strikethrough (`strike` attribute on `a:rPr`)
- Superscript / subscript (`baseline` attribute on `a:rPr`)
- Font shadow (effect list child on `a:rPr`)
- Character caps — none, all caps, small caps (`cap` attribute on `a:rPr`)
- Character spacing (`spc` attribute on `a:rPr`)
- Kerning (`kern` attribute on `a:rPr`)
- Baseline offset for super/subscript positioning (`baseline` attribute on `a:rPr`)
- East Asian font name (`a:ea` element)
- Complex script font name (`a:cs` element)
- Text frame orientation (`vert` attribute on `a:bodyPr`)
- Text frame columns (`numCol`, `spcCol` on `a:bodyPr`)
- `HasText` property on text frame

### Paragraph
- Bullet formatting — character bullets (`a:buChar`), numbered bullets (`a:buAutoNum`), picture bullets, bullet font (`a:buFont`), bullet color (`a:buClr`), bullet size (`a:buSzPct`, `a:buSzPts`), no bullet (`a:buNone`)
- First-line indent (`indent` attribute on `a:pPr`)
- Left margin (`marL` attribute on `a:pPr`)
- Tab stops (`a:tabLst` with `a:tab` children)
- Text direction / RTL (`rtl` attribute on `a:pPr`)
- Hanging punctuation (`hangingPunct` attribute)
- Baseline alignment (`fontAlgn` attribute)

### Table
- Cell borders — per-edge control: left, right, top, bottom, diagonal (`a:lnL`, `a:lnR`, `a:lnT`, `a:lnB`, `a:lnTlToBr`, `a:lnBlToTr` within `a:tcPr`) with full line formatting (color, weight, dash style)
- Table style application (built-in style GUIDs via `tblStyle` attribute)

### Line & Connector
- Arrowhead formatting — head and tail: type, width, length (`a:headEnd`, `a:tailEnd` within `a:ln`)
- Compound line style — single, double, thick-thin, etc. (`cmpd` attribute on `a:ln`)
- Line transparency (alpha modifier on line fill color)
- Line visibility (no-fill vs filled)
- Line pattern
- Query connected shapes on connectors (read `a:stCxn` / `a:endCxn` attributes)

### Shadow (Full Implementation)
- Shadow type — outer, inner, perspective (`a:outerShdw`, `a:innerShdw` within `a:effectLst`)
- Blur radius, distance, direction, alignment
- Shadow color with transparency
- Rotate with shape flag
- Visibility

### Additional Effects
- Reflection (`a:reflection` in effect list)
- Glow (`a:glow` in effect list)
- Soft edge (`a:softEdge` in effect list)

### 3D Formatting
- Extrusion depth, contour, material (`a:sp3d`)
- Bevel (top and bottom profiles)
- 3D scene — camera preset, rotation, field of view (`a:scene3d` → `a:camera`)
- Lighting rig — type, direction (`a:scene3d` → `a:lightRig`)

### Headers & Footers
- Slide-level header/footer configuration (`p:hf` element)
- Date/time, footer text, slide number placeholders on masters and layouts

### Metadata & Organization
- Tags — key-value string pairs on shapes and slides (separate `tags[N].xml` parts)
- Sections — named groups of slides (`p14:sectionLst` in presentation extensions)
- Custom document properties (beyond core properties)

### Comments
- Slide comments — author, text, position, datetime (`comments[N].xml` parts + `commentAuthors.xml`)

### Picture Format
- Brightness and contrast (`a:lum` on blip)
- Grayscale / black-and-white / washout (`a:grayscl`, `a:duotone` on blip)
- Transparency color
- Original image dimensions (from ImagePart)

### Chart
- Plot area positioning and formatting
- Chart area formatting
- Display blanks as (gap, zero, span)
- Secondary value and category axes
- 3D chart view (rotation, elevation, perspective)
- 3D chart surfaces (floor, walls)

### Themes
- Read/write theme color schemes
- Read/write theme font schemes (major and minor fonts)
- Theme effect schemes

### Callout Shapes
- Callout-specific formatting (accent bar, angle, length, gap) via adjustment handles on callout preset geometries

### WordArt / Text Effects
- Preset text warp (`a:prstTxWarp` on `a:bodyPr`)
- Text-level 3D scene and fill/outline

### Performance
- Shape lookup by name — build and maintain a name-indexed dictionary for O(1) access
- Shape lookup by ID — same approach with shape ID keys
- Lazy part loading — defer reading of image/media blobs until accessed
- Bulk shape operations — operate on multiple shapes without repeated XML tree walks

### Custom Slide Layouts
- Create new slide layout parts and link to a slide master
- Define placeholder positions and types on custom layouts

### Slide Import / Cross-Presentation Operations
- Import slides from another `.pptx` file
- Merge presentations (clone parts, remap relationships, deduplicate masters/layouts/images)
