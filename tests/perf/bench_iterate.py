"""Iteration benchmarks.

Walks `slides -> shapes -> text_frame.paragraphs -> paragraph.runs` and
measures separately:

- shape iteration (which currently calls `list(self._iter_member_elms())`
  inside both `__iter__` and `__getitem__`, see
  `src/pptx/shapes/shapetree.py:92-113`).
- text iteration, which is the suspected hot spot: `TextFrame.paragraphs`
  and `_Paragraph.runs` are plain `@property` (not `@lazyproperty`) and
  rebuild fresh tuples of proxy objects on every access — see
  `src/pptx/text/text.py:182-188` and `src/pptx/text/text.py:870-873`.

Two access patterns are timed:
- "single pass": each property accessed once per element.
- "repeated access": each property accessed N=10 times per element, to
  amplify the cost of the missing cache.
"""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from pptx import Presentation

REPEAT_ACCESS = 10


def _time_shape_iter(prs) -> float:
    t0 = time.perf_counter()
    count = 0
    for slide in prs.slides:
        for _ in slide.shapes:
            count += 1
    return time.perf_counter() - t0


def _time_text_single(prs) -> float:
    t0 = time.perf_counter()
    chars = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    chars += len(run.text)
    return time.perf_counter() - t0


def _time_text_repeated(prs) -> float:
    """Access `.paragraphs` / `.runs` REPEAT_ACCESS times per text frame.

    If these were `@lazyproperty`, repeated accesses would be ~free. Because
    they are plain `@property`, this scales linearly with REPEAT_ACCESS.
    """
    t0 = time.perf_counter()
    chars = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            tf = shape.text_frame
            for _ in range(REPEAT_ACCESS):
                for paragraph in tf.paragraphs:
                    for run in paragraph.runs:
                        chars += len(run.text)
    return time.perf_counter() - t0


@pytest.mark.perf
def bench_iterate_shapes_medium(synth_medium: Path) -> None:
    prs = Presentation(str(synth_medium))
    elapsed = _time_shape_iter(prs)
    print(f"\n[bench_iterate shapes medium] {elapsed * 1000:.2f} ms")


@pytest.mark.perf
def bench_iterate_text_single_medium(synth_medium: Path) -> None:
    prs = Presentation(str(synth_medium))
    elapsed = _time_text_single(prs)
    print(f"\n[bench_iterate text single medium] {elapsed * 1000:.2f} ms")


@pytest.mark.perf
def bench_iterate_text_repeated_medium(synth_medium: Path) -> None:
    prs = Presentation(str(synth_medium))
    elapsed = _time_text_repeated(prs)
    ratio = elapsed / max(_time_text_single(prs), 1e-9)
    print(
        f"\n[bench_iterate text repeated medium] "
        f"{elapsed * 1000:.2f} ms  (ratio to single-pass = {ratio:.2f}x)"
    )


@pytest.mark.perf
def bench_iterate_text_repeated_large(synth_large: Path) -> None:
    prs = Presentation(str(synth_large))
    elapsed = _time_text_repeated(prs)
    print(f"\n[bench_iterate text repeated large] {elapsed * 1000:.2f} ms")
