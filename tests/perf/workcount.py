"""Deterministic work count for a fixed workload: the "no slower than baseline" gate of #51.

Timing on shared CI runners is too noisy to gate on, so this counts work instead. It opens every
deck in the real-world corpus, reads it through the public API (`tests.unitutil.deckwalk.walk`)
and saves it, under cProfile, and counts the function calls made into `pptx` code.

The count moves with the regressions that matter: an extra lookup per shape, a list rebuilt in a
loop, a proxy built and thrown away. Repeated runs agree to within a call or two. It does not see
a change that makes the same calls faster or slower in C, such as `find()` vs `iterchildren()`;
the wall-clock baseline (`run_baseline.py`) covers that, as a report rather than a gate.

Call counts can differ slightly between Python versions, so the budget records the version it was
measured on, and `test_workcount.py` checks it only on that version (one CI cell).

    python -m tests.perf.workcount            # print the current count
    python -m tests.perf.workcount --update   # rewrite workcount_budget.json
"""

from __future__ import annotations

import cProfile
import io
import json
import os
import platform
import pstats
import sys
from pathlib import Path

HERE = Path(__file__).parent
CORPUS_DIR = HERE.parent / "test_files" / "real_world"
BUDGET_PATH = HERE / "workcount_budget.json"
#: The count may rise this far above the budget before the gate fails.
TOLERANCE = 0.02

_PPTX_DIR = os.sep + "pptx" + os.sep


def corpus_files() -> list[Path]:
    return sorted(CORPUS_DIR.glob("*.pptx"), key=lambda p: p.name.lower())


def _workload(files: list[Path]) -> None:
    from pptx import Presentation

    from tests.unitutil.deckwalk import walk

    for path in files:
        prs = Presentation(str(path))
        walk(prs)
        prs.save(io.BytesIO())


def measure() -> dict[str, int]:
    """Count the calls into `pptx` code made by one pass of the workload over the corpus."""
    files = corpus_files()
    # -- one pass first, so the counted pass sees warm caches (`qn()`, compiled XPath, the
    # -- per-class element maps) as a long-running batch job does --
    _workload(files)
    profile = cProfile.Profile()
    profile.enable()
    _workload(files)
    profile.disable()
    stats = pstats.Stats(profile).stats  # pyright: ignore[reportAttributeAccessIssue]
    pptx_calls = sum(
        primitive_calls
        for (filename, _, _), (primitive_calls, *_) in stats.items()
        if _PPTX_DIR in filename and "tests" + os.sep not in filename
    )
    return {"decks": len(files), "pptx_calls": pptx_calls}


def python_version() -> str:
    return "%d.%d" % sys.version_info[:2]


def load_budget() -> dict:
    return json.loads(BUDGET_PATH.read_text(encoding="utf-8"))


def write_budget(counts: dict[str, int]) -> None:
    import lxml.etree

    budget = {
        "python": python_version(),
        "implementation": platform.python_implementation(),
        "lxml": ".".join(str(n) for n in lxml.etree.LXML_VERSION[:3]),
        **counts,
    }
    BUDGET_PATH.write_text(json.dumps(budget, indent=2) + "\n", encoding="utf-8")


def main(argv: list[str]) -> int:
    counts = measure()
    print("python %s: %d decks, %d calls into pptx" % (python_version(), *counts.values()))
    if "--update" in argv:
        write_budget(counts)
        print("[write] %s" % BUDGET_PATH.relative_to(HERE.parent.parent))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
