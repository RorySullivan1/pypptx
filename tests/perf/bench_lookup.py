"""Verify the shape-name / shape-id indexes are O(1) amortized.

The indexes are `_BaseShapes._get_name_index` / `_get_id_index` in
`src/pptx/shapes/shapetree.py`. They are built lazily on first lookup and
cleared on any mutation (`_invalidate_shape_cache`); this measures, it does
not implement.

Strategy: build single-slide decks at three shape counts (10, 100, 1000).
Time the first lookup (cold — index build) and the median of subsequent
lookups (warm — dict hit). Assert that the warm-lookup time at n=1000 does
not exceed warm-lookup time at n=10 by more than a generous factor.
"""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from pptx import Presentation
from tests.perf import corpus

WARM_ITERATIONS = 500


def _time_warm_lookup(path: Path, target_name: str) -> float:
    prs = Presentation(str(path))
    shapes = prs.slides[0].shapes
    # Prime the index.
    shapes.get_by_name(target_name)
    t0 = time.perf_counter()
    for _ in range(WARM_ITERATIONS):
        shapes.get_by_name(target_name)
    return (time.perf_counter() - t0) / WARM_ITERATIONS


@pytest.mark.perf
def bench_lookup_linearity(tmp_path: Path) -> None:
    samples: dict[int, float] = {}
    for n in corpus.LOOKUP_COUNTS:
        deck = corpus.build_lookup_deck(n, tmp_path / f"lookup_{n}.pptx")
        # Look up a name in the middle of the document.
        target = f"box_{n // 2}"
        samples[n] = _time_warm_lookup(deck, target)
        print(f"\n[bench_lookup warm n={n}] {samples[n] * 1e6:.2f} us/lookup")

    smallest = samples[corpus.LOOKUP_COUNTS[0]]
    largest = samples[corpus.LOOKUP_COUNTS[-1]]
    growth = largest / max(smallest, 1e-9)
    print(
        f"\n[bench_lookup linearity] "
        f"n={corpus.LOOKUP_COUNTS[0]} -> n={corpus.LOOKUP_COUNTS[-1]}: "
        f"{growth:.2f}x"
    )
    # 100x size growth must not produce >5x warm-lookup time (O(1) amortized).
    assert growth < 5.0, (
        f"Warm lookup grew {growth:.2f}x as n went from "
        f"{corpus.LOOKUP_COUNTS[0]} to {corpus.LOOKUP_COUNTS[-1]}; "
        f"expected <5x for an O(1) index."
    )
