"""The work-count gate (issue #51, Phase 5): the corpus workload must not do more work than budgeted.

Runs in the normal suite, but only on the Python version the budget was measured on (one CI cell);
see `workcount.py`. When the count drops, rerun `python -m tests.perf.workcount --update` and commit
the lower budget so it becomes the new ceiling.
"""

from __future__ import annotations

import pytest

from tests.perf import workcount

pytestmark = [
    pytest.mark.skipif(
        not workcount.corpus_files(), reason="real-world corpus not present (not in the sdist)"
    ),
    pytest.mark.skipif(
        workcount.python_version() != workcount.load_budget()["python"],
        reason="work-count budget is measured on Python %s" % workcount.load_budget()["python"],
    ),
]


def it_does_no_more_work_than_the_budget_allows():
    budget = workcount.load_budget()

    counts = workcount.measure()

    assert counts["decks"] == budget["decks"], "the corpus changed; rerun with --update"
    ceiling = int(budget["pptx_calls"] * (1 + workcount.TOLERANCE))
    assert counts["pptx_calls"] <= ceiling, (
        "the corpus workload now makes %d calls into pptx, over the budget of %d (+%d%% allowed)."
        " Find the regression, or if the extra work is intended, rerun"
        " `python -m tests.perf.workcount --update` and commit the new budget."
        % (counts["pptx_calls"], budget["pptx_calls"], workcount.TOLERANCE * 100)
    )
