# pypptx

Create, read, and modify PowerPoint .pptx files in pure Python.

## Overview

pypptx is a Python library for working with Microsoft PowerPoint presentations. It operates directly on the Office Open XML (OOXML) file format without requiring PowerPoint or any other software to be installed. Create presentations from scratch, open and modify existing ones, or merge slides across files — all with a clean, Pythonic API.

## Installation

```bash
pip install pypptx
```

Dependencies (installed automatically): lxml, Pillow, XlsxWriter, typing\_extensions.

## Quick Start

```python
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()

# Add a title slide
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "Hello, pypptx!"
slide.placeholders[1].text = "Automating PowerPoint with Python"

# Add a blank slide with a formatted text box
slide2 = prs.slides.add_slide(prs.slide_layouts[5])
txBox = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(2))
tf = txBox.text_frame
tf.text = "This was created programmatically."
p = tf.add_paragraph()
p.text = "Second paragraph with custom formatting."
p.font.size = Pt(24)
p.font.bold = True

prs.save("hello.pptx")
```

## Features

- **Presentations** — create blank, open existing (`.pptx` / `.pptm` / `.potx`), save to file or stream, core + custom document properties, notes page size, slide-show settings (loop, kiosk/browse, pen color, slide range), custom shows, table-style listing, embedded fonts (list and remove)
- **Slides** — add, delete, duplicate, reorder, import across presentations, merge, computed slide numbers
- **180+ shape types** — autoshapes, pictures, tables, connectors, group shapes, freeforms, placeholders, OLE objects, movies, audio clips, callouts; 3D models, zooms and equations are listed read-only
- **Shape control** — z-order, visibility, lock aspect ratio, group containment, hyperlinks, click actions
- **Rich text formatting** — bold, italic, underline, color, size, shadow, strikethrough, caps, superscript/subscript, character spacing, kerning, bullets, tab stops, columns, RTL, vertical orientation
- **Fill styles** — solid, gradient (linear, radial, rectangular, path), pattern, picture (stretched or tiled), background (no fill)
- **Line formatting** — width, color, dash styles, compound styles, cap and join styles, arrowheads, transparency, pattern fill
- **75+ chart types** — bar, line, pie, scatter, bubble, area, radar, stock, surface, doughnut, of-pie (plus 3D variants); combo charts with trendlines, error bars, secondary axes, drop / hi-lo / up-down lines, log scale, 3D view, data labels with leader lines, data tables; Office 2016+ waterfall, histogram, box & whisker, treemap, sunburst and funnel charts (add and read; Pareto and region map read)
- **SmartArt** — detect SmartArt graphic frames, walk the node tree with text and levels, edit node text (PowerPoint re-lays out the diagram on open)
- **Tables** — create, insert/delete rows and columns, merge/split cells, borders (including diagonal), cell fills, cell text direction, banding, table styles
- **Visual effects** — outer/inner shadow with transparency, reflection, glow, soft edges
- **3D formatting** — extrusion, bevel, 3D scene, lighting, camera
- **Headers & footers** — per-slide date/time, footer text, slide-number overrides
- **Metadata** — tags, sections, custom document properties
- **Comments** — add, read, delete with author tracking and positioning
- **Picture effects** — brightness, contrast, grayscale, transparency color
- **Theme access** — color schemes, font schemes, effect schemes, fill/line style lists; replace a master's theme from another deck or a `.thmx`
- **Master text styles** — per-level title/body/other formatting on slide masters, plus the presentation default text style
- **Bulk operations** — ShapeRange API for alignment, distribution, batch property setting
- **Accessibility** — alternative text, title, decorative flag on shapes

## Examples

See the [`examples/`](examples/) directory for runnable scripts:

| Script | Description |
|--------|-------------|
| [`hello_world.py`](examples/hello_world.py) | Minimal presentation with title slide and core properties |
| [`add_textbox.py`](examples/add_textbox.py) | Text boxes with rich font formatting and paragraph alignment |
| [`add_picture.py`](examples/add_picture.py) | Insert and format images |
| [`shapes_and_fills.py`](examples/shapes_and_fills.py) | AutoShapes with solid, gradient, and pattern fills |
| [`shape_effects.py`](examples/shape_effects.py) | Shadow, glow, reflection, and 3D formatting |
| [`add_table.py`](examples/add_table.py) | Tables with merged cells, borders, and cell fills |
| [`add_chart.py`](examples/add_chart.py) | Bar and line charts with data labels and legends |
| [`modern_charts.py`](examples/modern_charts.py) | Waterfall, histogram, box & whisker, treemap, sunburst and funnel charts, and a chart data table |
| [`slide_operations.py`](examples/slide_operations.py) | Duplicate, reorder, and delete slides |
| [`merge_presentations.py`](examples/merge_presentations.py) | Import and merge slides across presentations |
| [`bullet_lists.py`](examples/bullet_lists.py) | Bullet formatting, numbered lists, and indentation |
| [`connectors_and_groups.py`](examples/connectors_and_groups.py) | Connectors with arrowheads and group shapes |
| [`comments_and_metadata.py`](examples/comments_and_metadata.py) | Comments, tags, sections, and custom properties |

## Requirements

- Python >= 3.9

## Limitations

- **Not a rendering engine** — cannot convert slides to images or PDF
- **Not a runtime** — cannot play slideshows or execute VBA macros
- **Not a print system** — cannot send presentations to a printer

## License

MIT
