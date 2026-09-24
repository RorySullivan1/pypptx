# 2026-09-24 15:05 · adopt-claudebrain-assets

**Goal:** Consolidate pypptx's Claude assets into .claude/ and adopt the claudeBrain factory's portable core

## What happened
- **Consolidated** the repo-root Claude docs into the standard layout:
  `internal_docs/MANIFEST.md` → `.claude/context/feature-manifest.md`,
  `internal_docs/TODO.md` → `.claude/context/feature-gaps.md` (git renames, content unchanged
  except `dev_map/` links), and `dev_map/` → `.meta/roadmap/` (INDEX + 4 stages + 8 cards).
  dev_map's template sections map onto the card shape: Enhancements → Goals, Test Plan →
  Objectives / acceptance; Objective / Dependencies / Out of scope kept as-is.
- **Adopted** the factory's portable core from `RorySullivan1/claudeBrain` @ `8e281af`
  (`example-project/.claude/`) by selection, per its `init-project` workflow:
  - Dogfooded floor: all shared hooks + generators, `session-memory`, `agent-finder`,
    `knowledge-router`, `token-optimizer`, `skill-distiller`, `claim-grounding`,
    `github-issues`, `github-pull-requests`, `token-manager`, `/version-set`, `/version-ship`,
    `/reindex`, `/epic`, `/issue`, `ship-version`, `verify-claims`, `establish-verification`.
  - Roadmap tier: `development-mapping`, `/roadmap-set`, `/roadmap-status`,
    `advance-roadmap-step`, `goal-auditor`, `roadmap_status.py`, `roadmap_guard.py`.
  - Families: Python (`python-development`/`-review`/`-maintenance`/`-deployment`,
    `coding-standards`, `python-developer` — rewritten for `src/pptx/`), `software-architect`,
    GitHub (`github-comments`, `github-releases`, `github-operator`), `technical-documentation-drafter`.
- Installed the GitHub skills' `installs.json` targets into `.github/` (PR + issue templates,
  `epic-autoclose.yml`). `.claude/settings.local.json` untracked + gitignored (personal).
- `establish-verification`: wrote `context/verification-surface.md` — pytest and example
  scripts proven able to fail; PowerPoint rendering human-gated; types/Sphinx/corpus unverified.

## Gotchas & dead ends
- `claim-grounding/reviews/ledger.jsonl` is per-project data (the factory's held 104 rows about
  VSTO/VBA/etc.) — reset to empty here, not copied.
- Copied assets referencing factory-only tooling (`/add-skill`, `author-asset`) now say so
  explicitly; `/reindex` lost its "second tree" step (pypptx has one tree).
- PR #53 shipped v0.2.0 but left dev_map saying "In Progress"; recorded as shipped.
  `pyproject.toml` is still `0.1.0` — "shipped" means merged, not published.

## State at end
- `.claude/` holds all Claude tooling; `.meta/` holds roadmap + cursor (v0.2.0 shipped;
  v0.2.1 next). build-hooks, catalog, asset_integrity green; pytest 3019 passed.

## Open threads
- Start v0.2.1 with `/version-set v0.2.1` on branch `v0.2.1` (expand its card first).
