# Development Roadmap

This directory holds versioning milestones charting the path from the current
release to `v1.0.0`. Each milestone is an **overview** — concise enough to scan,
detailed enough to scope the branch. Files are expanded with implementation
detail at the time development on that branch begins.

## Conventions

- **One file per version**, named `vMAJOR.MINOR.PATCH.md` (e.g. `v0.2.0.md`).
- **One branch per version**, named identically (e.g. `v0.2.0`). All work for
  that milestone happens on that branch; the branch merges into `main` at
  release.
- **Template per file**: Objective, Enhancements / Fixes, Dependencies, Out of
  Scope for this Version, Test Plan. (See `_TEMPLATE.md`.)
- **Sequencing**: Files are written ahead of time as overviews. Detail is added
  to the file at the same time the corresponding `vX.Y.Z` branch is opened. Do
  not implement work for a version before its overview is fleshed out.

## Current Status

| Version | Theme | Status |
|---|---|---|
| `v0.1.0` | Initial alpha + enum expansion | **Released** |
| `v0.2.0` | Documentation & MANIFEST resync | In Progress |
| `v0.2.1` | Example scripts for under-demonstrated areas | Planned |
| `v0.2.2` | Docstring audit & Sphinx scaffold | Planned |
| `v0.3.0` | Stability & real-world round-trip | Planned |
| `v0.4.0` | Transitions, animations (read), custom layouts | Planned |
| `v0.5.0` | Performance & developer experience | Planned |
| `v0.6.0` | Cross-run text find/replace & text helpers | Planned |
| `v1.0.0` | API freeze, comprehensive docs, SemVer commitment | Planned |

## Related

- `../internal_docs/MANIFEST.md` — authoritative scope & feature inventory
- `../internal_docs/TODO.md` — granular feature buildout log (sections 1–24)
- `../CLAUDE.md` — references this directory under "Development Roadmap"
