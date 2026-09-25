"""Cold-open wall-clock benchmark.

Times `Presentation(path)` (parses the package, loads relationships, mounts
custom element classes) for each synthetic deck size. The path-only variant
forces a fresh ZipFile + lxml parse on every iteration.
"""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from pptx import Presentation

ITERATIONS = 5


def _time_open(path: Path) -> float:
    """Return median wall-clock seconds across `ITERATIONS` cold opens."""
    samples = []
    for _ in range(ITERATIONS):
        t0 = time.perf_counter()
        Presentation(str(path))
        samples.append(time.perf_counter() - t0)
    samples.sort()
    return samples[len(samples) // 2]


@pytest.mark.perf
def bench_open_small(synth_small: Path) -> None:
    median = _time_open(synth_small)
    print(f"\n[bench_open small] median={median * 1000:.2f} ms")
    assert median < 5.0


@pytest.mark.perf
def bench_open_medium(synth_medium: Path) -> None:
    median = _time_open(synth_medium)
    print(f"\n[bench_open medium] median={median * 1000:.2f} ms")
    assert median < 10.0


@pytest.mark.perf
def bench_open_large(synth_large: Path) -> None:
    median = _time_open(synth_large)
    print(f"\n[bench_open large] median={median * 1000:.2f} ms")
    assert median < 30.0
