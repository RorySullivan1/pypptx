# context/

Reference docs Claude can deep-read when a task needs them. `CLAUDE.md` points
here so the main session stays lean — Claude only opens what's relevant.

This layer has **two tiers**:

1. **Project docs** (flat `*.md` below) — longer-form reference that `CLAUDE.md` points at:
   the scope manifest, the feature-gap log, the verification surface. Plain markdown, no
   frontmatter, listed in the **Manifest** below. They hold *facts*; how-to guidance lives
   in `../skills/` (see below).
2. **Reference notes** (`notes/*.md` + auto-generated `INDEX.md`) — small, declarative,
   read-on-demand reference cards (a concept, an external-system fact, a schema, a system
   map). `INDEX.md` is the always-loaded catalog (surfaced at SessionStart); each note is
   read only when its topic is relevant. The `knowledge-router` skill decides what earns a
   note, and its `context.py` engine creates notes and regenerates the catalog so it can't
   drift. Run `python ../skills/knowledge-router/scripts/context.py list` to see them.

## When a skill and a brief cover the same ground, the skill wins

The two tiers overlap by design, but that does not license the same guidance in both
places — and left unstated it produces exactly that. Two rules settle it:

- **A skill is loaded by its own description when the task matches; a brief is loaded
  only if something points at it.** Reachability, not scope, is what makes a home
  canonical. Where both cover a topic, **the skill is canonical** and the brief must not
  restate it.
- **A brief earns its place only by what no skill carries** — a whole-stack stance a
  task-scoped skill cannot express — and it must be reachable: referenced from the
  installing project's `CLAUDE.md`, or it is dead weight. A brief nothing points at is
  not a fallback; it is a second copy that cannot be corrected because nobody reads it.

## Manifest (project-instruction briefs)

| File | What it's for |
|---|---|
| `feature-manifest.md` | Authoritative scope & feature inventory — what pypptx implements, the remaining gaps, and what is deliberately out of scope (rendering, slideshow, print, macros). Read before proposing a feature "PowerPoint can do". |
| `feature-gaps.md` | Granular feature buildout log (sections 1–24) with the OOXML schema location of each item. States the **bottom-up, types-first** strategy. |
| `verification-surface.md` | Who can confirm each surface worked — the per-surface `agent-runnable` / `human-gated` / `unverified` table, written by the `establish-verification` workflow. |

Reference notes are catalogued automatically in `INDEX.md` — not listed here.
