"""Run the performance baseline outside of pytest.

Produces:
- `tests/perf/baseline.json` — machine-readable timings + memory peaks.
- Replaces the "Baseline" section at the end of `.claude/context/performance.md`
  with wall-clock, memory, and cProfile top-N entries.

Usage:
    python -m tests.perf.run_baseline
"""

from __future__ import annotations

import cProfile
import io
import json
import platform
import pstats
import sys
import time
import tracemalloc
from datetime import datetime, timezone
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches
from tests.perf import corpus

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
BASELINE_PATH = Path(__file__).resolve().parent / "baseline.json"
PERF_DOC_PATH = REPO_ROOT / ".claude" / "context" / "performance.md"
BASELINE_HEADING = "## Baseline"

PROFILE_TOP_N = 20
OPEN_ITERATIONS = 5
REPEAT_ACCESS = 10


def _median(samples: list[float]) -> float:
    samples = sorted(samples)
    return samples[len(samples) // 2]


def _build_corpus(workdir: Path) -> dict[str, Path]:
    decks: dict[str, Path] = {}
    for spec in (corpus.SMALL, corpus.MEDIUM, corpus.LARGE):
        t0 = time.perf_counter()
        decks[spec.name] = corpus.build_deck(spec, workdir / f"{spec.name}.pptx")
        print(
            f"[build] {spec.name:<6}  slides={spec.n_slides:<4}"
            f"  shapes/slide={spec.shapes_per_slide:<3}"
            f"  paragraphs/shape={spec.paragraphs_per_shape:<2}"
            f"  images={spec.n_images:<3}"
            f"  built in {time.perf_counter() - t0:.2f}s"
            f"  size={decks[spec.name].stat().st_size // 1024} KiB"
        )
    return decks


def _bench_open(deck: Path) -> float:
    samples = []
    for _ in range(OPEN_ITERATIONS):
        t0 = time.perf_counter()
        Presentation(str(deck))
        samples.append(time.perf_counter() - t0)
    return _median(samples)


def _bench_iterate_text(prs, repeats: int) -> float:
    t0 = time.perf_counter()
    chars = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            tf = shape.text_frame
            for _ in range(repeats):
                for paragraph in tf.paragraphs:
                    for run in paragraph.runs:
                        chars += len(run.text)
    return time.perf_counter() - t0


def _bench_save(deck: Path) -> float:
    prs = Presentation(str(deck))
    prs.slides[0].shapes.add_textbox(
        Inches(0.1), Inches(0.1), Inches(1.0), Inches(0.4)
    )
    buf = io.BytesIO()
    t0 = time.perf_counter()
    prs.save(buf)
    return time.perf_counter() - t0


def _peak_memory_open(deck: Path) -> tuple[int, int]:
    tracemalloc.start()
    Presentation(str(deck))
    current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return current, peak


def _profile_workload(deck: Path) -> str:
    """cProfile the full open+iterate+save cycle on `deck`. Returns top-N text."""
    prof = cProfile.Profile()
    prof.enable()
    prs = Presentation(str(deck))
    _bench_iterate_text(prs, REPEAT_ACCESS)
    prs.slides[0].shapes.add_textbox(
        Inches(0.1), Inches(0.1), Inches(1.0), Inches(0.4)
    )
    buf = io.BytesIO()
    prs.save(buf)
    prof.disable()

    out = io.StringIO()
    stats = pstats.Stats(prof, stream=out).strip_dirs().sort_stats("cumulative")
    stats.print_stats(PROFILE_TOP_N)
    return out.getvalue()


def _format_table(headers: list[str], rows: list[list[str]]) -> str:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def main() -> int:
    workdir = REPO_ROOT / "tests" / "perf" / "_corpus"
    workdir.mkdir(parents=True, exist_ok=True)

    print("=" * 72)
    print("Performance baseline — building corpus")
    print("=" * 72)
    decks = _build_corpus(workdir)

    print("\n" + "=" * 72)
    print("Wall-clock (median of N iterations)")
    print("=" * 72)
    timings: dict[str, dict[str, float]] = {}
    for name, deck in decks.items():
        open_s = _bench_open(deck)
        prs = Presentation(str(deck))
        iter_single_s = _bench_iterate_text(prs, 1)
        iter_repeat_s = _bench_iterate_text(prs, REPEAT_ACCESS)
        save_s = _bench_save(deck)
        timings[name] = {
            "open_median_s": open_s,
            "iter_text_single_s": iter_single_s,
            "iter_text_repeated_s": iter_repeat_s,
            "save_s": save_s,
        }
        print(
            f"[time] {name:<6}"
            f"  open={open_s * 1000:>7.2f} ms"
            f"  iter1x={iter_single_s * 1000:>7.2f} ms"
            f"  iter{REPEAT_ACCESS}x={iter_repeat_s * 1000:>8.2f} ms"
            f"  save={save_s * 1000:>7.2f} ms"
        )

    print("\n" + "=" * 72)
    print("Memory (tracemalloc on open)")
    print("=" * 72)
    memory: dict[str, dict[str, int]] = {}
    for name, deck in decks.items():
        current, peak = _peak_memory_open(deck)
        memory[name] = {"current_bytes": current, "peak_bytes": peak}
        print(
            f"[mem]  {name:<6}  current={current / 1024:>8.1f} KiB"
            f"  peak={peak / 1024:>8.1f} KiB"
        )

    print("\n" + "=" * 72)
    print(f"cProfile top {PROFILE_TOP_N} (workload: open + iterate + save, medium deck)")
    print("=" * 72)
    profile_text = _profile_workload(decks["medium"])
    print(profile_text)

    baseline = {
        "schema_version": 1,
        "captured_at": datetime.now(timezone.utc).isoformat(),
        "platform": {
            "python": sys.version.split()[0],
            "implementation": platform.python_implementation(),
            "system": platform.system(),
            "machine": platform.machine(),
        },
        "corpus": {
            spec.name: {
                "n_slides": spec.n_slides,
                "shapes_per_slide": spec.shapes_per_slide,
                "paragraphs_per_shape": spec.paragraphs_per_shape,
                "n_images": spec.n_images,
                "file_size_bytes": decks[spec.name].stat().st_size,
            }
            for spec in (corpus.SMALL, corpus.MEDIUM, corpus.LARGE)
        },
        "timings_seconds": timings,
        "memory_bytes_on_open": memory,
        "iter_repeat_access_count": REPEAT_ACCESS,
        "open_iterations": OPEN_ITERATIONS,
    }

    BASELINE_PATH.write_text(json.dumps(baseline, indent=2) + "\n")
    print(f"\n[write] {BASELINE_PATH.relative_to(REPO_ROOT)}")

    _write_perf_doc_section(baseline, profile_text)
    print(f"[write] {PERF_DOC_PATH.relative_to(REPO_ROOT)} (Baseline section)")
    return 0


def _write_perf_doc_section(baseline: dict, profile_text: str) -> None:
    timings = baseline["timings_seconds"]
    memory = baseline["memory_bytes_on_open"]
    corpus_info = baseline["corpus"]

    corpus_rows = [
        [
            name,
            str(info["n_slides"]),
            str(info["shapes_per_slide"]),
            str(info["paragraphs_per_shape"]),
            str(info["n_images"]),
            f"{info['file_size_bytes'] / 1024:.1f} KiB",
        ]
        for name, info in corpus_info.items()
    ]
    timing_rows = [
        [
            name,
            f"{timings[name]['open_median_s'] * 1000:.2f}",
            f"{timings[name]['iter_text_single_s'] * 1000:.2f}",
            f"{timings[name]['iter_text_repeated_s'] * 1000:.2f}",
            f"{timings[name]['save_s'] * 1000:.2f}",
        ]
        for name in corpus_info
    ]
    memory_rows = [
        [
            name,
            f"{memory[name]['current_bytes'] / 1024:.1f}",
            f"{memory[name]['peak_bytes'] / 1024:.1f}",
        ]
        for name in corpus_info
    ]

    section = f"""

## Baseline

Captured: `{baseline['captured_at']}`
Platform: `{baseline['platform']['implementation']} {baseline['platform']['python']}` on `{baseline['platform']['system']} / {baseline['platform']['machine']}`

### Corpus

{_format_table(
    ["fixture", "slides", "shapes/slide", "paragraphs/shape", "images", "size"],
    corpus_rows,
)}

### Wall-clock (ms)

Open is median of {baseline['open_iterations']} cold opens.
`iter1x` walks `slides -> shapes -> paragraphs -> runs` once per text frame.
`iter{baseline['iter_repeat_access_count']}x` repeats the inner walk
{baseline['iter_repeat_access_count']} times to amplify the cost of
`paragraphs`/`runs` rebuilding on every access
(`TextFrame.paragraphs`, `_Paragraph.runs` in `src/pptx/text/text.py`).

{_format_table(
    ["fixture", "open", "iter1x", f"iter{baseline['iter_repeat_access_count']}x", "save"],
    timing_rows,
)}

### Memory on open (KiB, tracemalloc)

{_format_table(["fixture", "current", "peak"], memory_rows)}

### cProfile top {PROFILE_TOP_N} (medium fixture, open + iterate + save)

```
{profile_text.strip()}
```

### Reproducing

```
python -m tests.perf.run_baseline
pytest -m perf
```
"""

    # -- replace any earlier Baseline section (always the last one) rather than appending --
    doc = PERF_DOC_PATH.read_text()
    heading_at = doc.find("\n" + BASELINE_HEADING + "\n")
    if heading_at != -1:
        doc = doc[:heading_at]
    PERF_DOC_PATH.write_text(doc.rstrip() + section)


if __name__ == "__main__":
    raise SystemExit(main())
