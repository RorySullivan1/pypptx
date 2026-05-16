# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Identity

- Distribution name: **`pypptx`** (on PyPI). Import name: **`pptx`**. Sources live at `src/pptx/`, NOT `src/pypptx/`. Test code and user code both `from pptx import ...`.
- Pure-Python library for reading/writing PowerPoint `.pptx` (OOXML / ECMA-376) files. No PowerPoint, no rendering engine, no runtime dependency — every feature must be expressible as an edit to XML parts inside the package's ZIP container.
- Python `>= 3.9`. Hard deps: `lxml`, `Pillow`, `XlsxWriter`, `typing_extensions`.

## Common Commands

```bash
# Install in editable mode with dev extras
pip install -e ".[dev]"

# Run the full test suite (configured in pyproject.toml -> [tool.pytest.ini_options])
pytest

# Run a single test file / class / function
pytest tests/test_slide.py
pytest tests/test_slide.py::DescribeSlide
pytest tests/test_slide.py::DescribeSlide::it_provides_access_to_its_shapes

# Coverage
pytest --cov=pptx --cov-report=term-missing
```

Test discovery is non-default: classes match `Describe*`, functions match `it_*` or `test_*` (BDD style). Keep new tests in that shape or pytest will silently skip them.

## Architecture

Four layers, low to high. Each layer only depends on the ones below it.

| Layer | Where | Responsibility |
|---|---|---|
| **OPC** | `pptx/opc/` | Open Packaging Convention — ZIP I/O, content-types, part relationships, `PackURI`. |
| **OXML** | `pptx/oxml/` | lxml custom element classes mapped to ECMA-376 schema types. Element classes carry attribute/child accessors; the parser inflates them automatically when their tag is encountered. |
| **Parts** | `pptx/parts/` | One class per content-type — `SlidePart`, `ChartPart`, `ImagePart`, etc. A part owns one XML element tree (or a blob) plus its outbound relationships. |
| **API** | `pptx/api.py`, `pptx/presentation.py`, `pptx/slide.py`, `pptx/shapes/`, `pptx/text/`, `pptx/chart/`, `pptx/dml/`, `pptx/table.py`, `pptx/action.py` | User-facing Pythonic proxies. Modeled loosely on the VBA PowerPoint object model. |

Entry point is `pptx.api.Presentation(path_or_stream=None)`. With no argument it loads `src/pptx/templates/default.pptx`. The returned `Presentation` is a proxy whose `.part` is a `PresentationPart`, which owns the package graph through `pptx.package.Package` (`OpcPackage` subclass).

### How content-types resolve to part classes

`pptx/__init__.py` builds a `content_type → Part class` map and registers it onto `PartFactory.part_type_for`. When the package reader walks the ZIP it consults this map to instantiate the right Part subclass. **New part types must be registered here** — adding the class alone is not enough.

### Element proxy hierarchy (`pptx/shared.py`)

- `ElementProxy` — wraps one lxml element. Equality is "same underlying element".
- `ParentedElementProxy` — adds `._parent` (and `.part` via the parent) for objects that need ancestor context (most shapes).
- `PartElementProxy` — wraps the root element of a part (e.g. a slide). Owns its `.part` directly.

Most API objects subclass one of these. Keep them as value objects — no mutable Python-side state; mutate XML instead.

### oxml / xmlchemy

`pptx/oxml/xmlchemy.py` provides a declarative meta-system for defining custom lxml element classes (child accessors, attribute typing via `simpletypes`, choice groups). Classes are registered with `register_element_cls("ns:tag", Cls)` so the lxml parser auto-promotes elements to the right Python class on load. When adding XML support, add the element class in `pptx/oxml/...`, register it, then build the API proxy on top — do not parse XML directly from API code.

### Units

`pptx/util.py` defines `Length` (int subclass storing EMUs) with subclasses `Emu`, `Inches`, `Cm`, `Mm`, `Pt`, `Centipoints`. All public APIs that accept measurements take `Length` — pass `Inches(1)`, not `914400`. PowerPoint stores font sizes in centipoints internally, which is why that subclass exists.

`@lazyproperty` (in `pptx/util.py`) is the standard cache decorator — used pervasively for derived collections (`slides`, `slide_layouts`, etc.). Prefer it over manual `_cache` attributes.

## Testing Conventions

- `tests/unitutil/mock.py` — wrappers over `unittest.mock` (`class_mock`, `instance_mock`, `method_mock`, `property_mock`). Use these rather than raw `patch()`; they auto-register teardown.
- `tests/unitutil/cxml.py` — a homegrown **C**ompact **XML** **E**xpression **L**anguage. `cxml.element("p:sp/p:nvSpPr/p:cNvPr{id=1,name=foo}")` produces a parsed lxml element. Used everywhere to build fixture XML compactly. Read the module's docstring before writing complex test XML by hand.
- Tests are organized to mirror `src/pptx/` (e.g. `tests/oxml/`, `tests/shapes/`, `tests/parts/`).
- Round-trip tests live in `tests/test_roundtrip.py` and exercise real `.pptx` save/load cycles.

## Project Conventions

- `from __future__ import annotations` is at the top of every source file; `TYPE_CHECKING` blocks are used for proxy-class imports to avoid cycles between layers.
- `py.typed` ships with the package — keep annotations accurate; downstream users rely on them.
- The exception hierarchy in `pptx/exc.py` (`ShapeError`, `SlideError`, `ChartError`, `TableError`, `PackageError`, etc.) is the preferred raise target inside the API layer. Avoid bare `ValueError`/`TypeError` in user-facing code paths.
- `*.pptx` is gitignored — example scripts write to the working directory and outputs are not checked in.

## Internal Docs

- `internal_docs/MANIFEST.md` — authoritative scope/feature inventory and explicit out-of-scope items. Consult before proposing a feature that "PowerPoint can do" — many such features (rendering, slideshow, print, macro execution) are deliberately excluded.
- `internal_docs/TODO.md` — granular roadmap of feature gaps with their OOXML schema location. Strategy stated there: **bottom-up, types-first** — complete the OXML layer (element classes, simple types, enums) before building API wrappers on top.
- `dev_map/` — versioned development roadmap. One `vMAJOR.MINOR.PATCH.md` file per planned milestone between the current release and `v1.0.0`. Each file follows the template in `dev_map/_TEMPLATE.md` (Objective / Enhancements / Dependencies / Out of Scope / Test Plan). Work for a given version happens on a branch of the same name (`vX.Y.Z`); the milestone file is expanded with implementation detail at the time that branch is opened. See `dev_map/README.md` for status and conventions.
- `examples/` — runnable scripts demonstrating each feature area; useful as informal smoke tests when changing public APIs.
