# Development Map — pypptx

> Auto-surfaced at session start. Keep to ~one screen; detail lives in `stages/`.
> The cursor (`←`) is the next version to build. Statuses: planned · in-progress · shipped.
> Author/re-slice with `/roadmap-set`; reconcile with `/roadmap-status`. Each version is
> built on a branch named after it (`vX.Y.Z`); its card is expanded with implementation
> detail when that branch opens — don't start a version whose card is still an overview.

## Objective
A pure-Python, dependency-light library for reading and writing PowerPoint `.pptx` files
(OOXML / ECMA-376) with a Pythonic, VBA-inspired object model — no PowerPoint, no rendering
engine — reaching a stable, documented, SemVer-committed `1.0`.
Scope and exclusions: `.claude/context/feature-manifest.md`.

## Stages
1. **Foundation** — the four layers + OXML type expansion. *Milestone: installable 0.1.0 alpha.*
2. **Truthful docs** — manifest/README, examples, docstrings + Sphinx match the code. *Milestone: every feature group documented, demonstrated, and typed.*
3. **Hardening & remaining gaps** — real-world round-trip, transitions/animations (read), custom layouts, performance, cross-run text. *Milestone: real decks round-trip; only deliberate gaps remain.*
4. **Stable release** — API freeze, hosted docs, SemVer commitment. *Milestone: 1.0.0 on PyPI.*

## Versions
| Version | Stage | Goal (one line)                                         | Status |
|---------|-------|---------------------------------------------------------|--------|
| v0.1.0  | 01    | Initial alpha + enum expansion                          | shipped |
| v0.2.0  | 02    | Documentation & MANIFEST resync                         | shipped |
| v0.3.0  | 03    | Feature-gap epics: SmartArt, chartex, media, text styles… (recorded after the fact) | shipped |
| v0.3.1  | 02    | Example scripts for under-demonstrated areas            | shipped |
| v0.3.2  | 02    | Docstring audit & Sphinx scaffold                       | shipped |
| v0.4.0  | 03    | Stability & real-world round-trip                       | shipped |
| v0.5.0  | 03    | Transitions, animations (read), custom layouts          | in-progress ← cursor |
| v0.6.0  | 03    | Performance & developer experience (harness done)       | planned |
| v0.7.0  | 03    | Cross-run text find/replace & text helpers              | planned |
| v1.0.0  | 04    | API freeze, comprehensive docs, SemVer commitment       | planned (sketch) |
