# The verification surface — who can confirm this worked

pypptx's answer, per surface, to one question: **can the agent confirm this itself, or must a
human — and why?** Every workflow that ends in "done" (`advance-roadmap-step`, `ship-version`)
reads this file instead of re-deriving what it can check.

Produced by the **`establish-verification`** workflow on adoption (2026-09-24); refresh it with
the same workflow when a surface or check changes. The 2026-09-27 rows were stamped by hand
after PR #112 (cross-platform CI, LF line endings, the `pyparsing` gap) — a full re-derivation
via the workflow has not been run since adoption.

## The three tiers

| Tier | Means | Requires |
|---|---|---|
| `agent-runnable` | A command the agent can execute here, whose result is the verdict | The exact command, and evidence it has been seen to **fail** |
| `human-gated` | Confirmation needs something the agent cannot reach | The blocker, and what the human is asked to report back |
| `unverified` | Nobody has a check yet | An owner and a next action — this tier is a debt marker, not a resting state |

### Rules that make the tiers mean something

- **Name the command, not the capability.** "The tooling is tested" is not a check; `pytest -q
  tests/test_util.py` is. If the *check* column holds nothing you could paste into a shell, the row is
  `unverified` wearing a disguise.
- **A check that cannot fail is not a check.** Before a row may be marked `agent-runnable`, the
  command must have been observed to fail on deliberately broken input at least once — the same
  controls discipline `claim-grounding` requires of a probe. Record that in *last-run*.
- **`human-gated` is a real answer, not a failure.** It is what lets a workflow place an
  honest gate instead of reporting success it cannot see. What makes the row useful is the
  **reason** (so nobody re-litigates it) and the **report contract** (what the human returns —
  often just a binary "worked / didn't", which is all an air gap can carry).
- **Split a surface rather than calling it "mixed".** If part of a surface is machine-checkable
  and part isn't, that is two rows. "Mixed" hides which half is actually covered.
- **A stale `agent-runnable` row is a claim about the past.** *last-run* is what distinguishes a
  check that works from one that used to.
- **Say whether the check gates or only reports.** Some checks are advisory by
  contract — `asset_integrity.py` never vetoes a commit and reports its verdict in its output. A caller that reads exit status alone will record a
  pass that never happened. Where the two differ, the *check* column says which to read.
- **Never promote a tier to close a gap.** Downgrading is free; upgrading needs a command and a
  failure. Reporting `unverified` is always cheaper than a wrong "verified".


## The surface table — pypptx

| Surface | Check | Tier | Reason | Last run |
|---|---|---|---|---|
| Library behaviour (`src/pptx/`) incl. round-trip | `pytest -q` (targeted first: `pytest -q tests/<area>/`); CI runs it on every push/PR via `.github/workflows/tests.yml` | `agent-runnable` | Pure Python; the suite includes save/load round-trips in `tests/test_roundtrip.py` | 2026-09-27 — 3657 passed on Windows locally **and** 7/7 CI cells green on `main` (Linux 3.9–3.13, Windows 3.13, macOS 3.13; run 36324126203). Seen to fail: 2026-09-24 `tests/test_util.py` went 2 failed after mutating `_EMUS_PER_INCH`; 2026-09-26 a CRLF working tree gave 38 failed + 14 errors |
| Public API as used (`examples/`) | `python examples/<script>.py` — exit status is the verdict; CI runs all of them in the ubuntu/py3.13 cell | `agent-runnable` | Scripts build and save a `.pptx` end to end; outputs are gitignored — run from a scratch cwd | 2026-09-27 — all 15 runnable scripts exit 0 in CI and locally on Windows (`batch_process.py` excluded: it takes IN_DIR/OUT_DIR). Seen to fail: 2026-09-24 a bad layout index raised `SlideError` |
| Declared dependencies are complete | `python -m venv <tmp> && <tmp>/pip install -e ".[dev]" && pytest -q`; every CI cell does exactly this | `agent-runnable` | A developer machine carries packages the wheel never declares, so an undeclared test dep stays invisible until a clean install | 2026-09-27 — clean venv (lxml 6.1.3, pillow 12.3.0, pyparsing 3.3.3, xlsxwriter 3.2.9, pytest 9.1.1): 3657 passed. Seen to fail: 2026-09-27, `pyparsing` undeclared → all 7 CI cells failed at collection (run 36290495812) |
| Output opens and renders correctly | open the saved `.pptx` in PowerPoint (or LibreOffice) and look | `human-gated` | No renderer is part of the project or reachable as a check; the human reports "opens without repair prompt / looks right" per file | — |
| Real-world file compatibility | none yet — a PowerPoint-authored corpus arrives in v0.4.0 | `unverified` | Owner: v0.4.0 card; next: add the corpus + round-trip driver | — |
| Type annotations (`py.typed`) | none — no mypy/pyright configured | `unverified` | Owner: v0.3.2 / v1.0.0 cards; next: configure a type checker in `pyproject.toml` | — |
| API docs (Sphinx) | none — `docs/` does not exist yet | `unverified` | Owner: v0.3.2 card; next: `sphinx-build -W` once the scaffold lands | — |
| Asset shape (`.claude/`) | `asset_integrity.py` fed a git-commit hook payload — **advisory: reports, never vetoes; read the output** | `agent-runnable` | Pure file-shape analysis | 2026-09-24 — silent on the real tree; caught a deliberate `name:`/folder mismatch |

**Reading `—` in *last-run*.** It means the tier is **claimed, not proved** — nobody has yet
watched that command fail on broken input. Treat such a row as `unverified` until a run stamps it.

## See also

- **`establish-verification`** workflow — how the rows get produced and refreshed.
- **`claim-grounding`** skill — the same question asked of an asset's *claims* rather than the
  project's *work*.
