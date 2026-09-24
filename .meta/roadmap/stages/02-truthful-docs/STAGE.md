# Stage 02 — Truthful docs  ◀ current stage

**Goal:** make what pypptx *says* match what it *does*, so every later version builds on
accurate baselines: the scope manifest and README (v0.2.0), a runnable example per feature
area (v0.2.1), then docstrings + a local Sphinx build (v0.2.2).

**Versions:** v0.2.0 (MANIFEST + README resync — shipped, PR #53), v0.2.1 (example scripts
for under-demonstrated areas), v0.2.2 (docstring audit + Sphinx scaffold).

**Milestone (stage exit):** every feature group in `feature-manifest.md` is named in the
README, has an example script that round-trips, and has typed docstrings that build to
HTML locally with no Sphinx warnings.
