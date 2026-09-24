# `.claude/` — project infrastructure

This directory holds the typed building blocks Claude Code uses for a project.
Each subdirectory is a distinct *layer* with a distinct job. The layers compose.

> Adopted from the **claudeBrain** factory (`RorySullivan1/claudeBrain`,
> `example-project/.claude/`) by selection — see `memory/INDEX.md` § Decisions for what was
> taken, what was left out, and why. Shared assets are copies: improve them upstream in the
> factory, then re-copy, rather than forking them here.

## The composability stack

```
hooks      ← enforcement, underneath everything (Claude cannot skip these)
─────────────────────────────────────────────────────────────────────────
workflows  ▸  commands  ▸  agents  ▸  skills
(orchestrate)  (one-shot)   (isolated)  (expertise)
```

- **hooks/** — Deterministic shell scripts run by the harness on lifecycle events
  (`PreToolUse`, `PostToolUse`, `SessionStart`, …). They are the enforcement layer
  *underneath* the prompt stack — the model cannot choose to skip them. Use for
  anything that must *always* happen: formatting, branch guards, write protection,
  cache warming. Configured in `settings.json`.
- **workflows/** — Multi-step autonomous orchestrations. Claude executes a scripted
  sequence that can loop, branch, and spawn agents. Each is a markdown file.
- **commands/** — Single-shot, stateless prompt templates — saved prompts you'd
  otherwise retype. One file per command (`/<name>`).
- **agents/** — Isolated subagents spawned with clean context. They do focused work
  and return only a summary, so they don't bleed context into the main session.
- **skills/** — Domain-expertise bundles that tell Claude *how to think and behave*
  for a task type. Applied within a session or an agent's context. One folder per
  skill containing `SKILL.md`; the folder name equals the skill's `name:` frontmatter.

## Supporting files

- **context/** — Reference docs (architecture notes, schemas, stack instructions).
  `CLAUDE.md` points here; Claude deep-reads only what's relevant to the task. See
  `context/README.md` for the manifest.
- **settings.json** — Permissions, model, and hook configuration.
- **memory/** — Cross-session state via the `session-memory` skill: an auto-loaded
  `INDEX.md` plus append-only `sessions/*.md` logs (loaded/persisted by the lifecycle
  hooks in `settings.json`). Replaces a static `DECISIONS.md` log.
  **One unit of work at a time.** `memory/INDEX.md` § State and `.meta/version` both name
  *the* current work. Runtime is per-worktree and safe; parallel branches collide only at
  merge, loudly, and State is the one hunk you must never resolve by taking a side. The
  collision is reproducible — `hooks/probes/probe_parallel_state.py`.
- **CATALOG.md** — A generated, **on-demand** inventory of every skill, agent, command, and
  workflow with a one-line purpose. `CLAUDE.md` references it by path instead of enumerating
  assets (skills/agents already auto-load by their `description:`). Produced by
  `hooks/catalog.py` (a mechanical generator), kept fresh by a `PostToolUse` auto-rebuild +
  a `SessionStart` staleness warning, and regenerated with the `/reindex` command.

## Status in this project

- **skills/** — the Python family (`python-development` / `-review` / `-maintenance` /
  `-deployment`, `coding-standards`), the GitHub family, `technical-documentation-drafter`,
  `development-mapping`, and the operational set (`session-memory`, `agent-finder`,
  `knowledge-router`, `token-optimizer`, `skill-distiller`, `claim-grounding`).
- **agents/** — `python-developer` (the executor for `src/pptx/`), `software-architect`,
  `goal-auditor`, `github-operator`, `token-manager`.
- **commands/** — `/version-set`, `/version-ship`, `/roadmap-set`, `/roadmap-status`,
  `/epic`, `/issue`, `/reindex`.
- **workflows/** — `advance-roadmap-step`, `ship-version`, `verify-claims`,
  `establish-verification`.
- **hooks/** — memory + context lifecycle hooks, git guards (version / roadmap / asset
  integrity), context-economy guards, and the `build-hooks.py` / `catalog.py` generators,
  compiled into `settings.json`.
- **context/** — the scope manifest, the feature-gap log, and the verification surface.
- **memory/** is active; the roadmap lives in `../.meta/roadmap/`, the cursor in
  `../.meta/version`.

For the full, current list of any layer, read **`CATALOG.md`** (regenerate with `/reindex`)
— this README describes the layers; the catalog enumerates them.
