# Performance

## Status

**Stage 1: profiling harness baseline.** Preparatory infrastructure for the
`v0.5.0` milestone (see `../dev_map/v0.5.0.md`). No `src/pptx/` code changes
in Stage 1 — measurement only.

## Scope alignment

This document feeds `dev_map/v0.5.0.md` ("Performance & Developer Experience").
The v0.5.0 overview explicitly puts the following out of scope; this document
honors those boundaries:

- **No native (C-extension) optimization.** `lxml` is already C-backed; pure
  Python only.
- **No library-level multi-threading or async save.** The OPC zip format and
  `lxml` element trees are not thread-safe; concurrency belongs in user code.

A user-facing `concurrent.futures.ProcessPoolExecutor` recipe is a possible
future addition under `examples/` but is library-orthogonal — `pypptx` has
no module-level mutable state, so per-file fan-out is safe today.

## What already exists (verify, do not reimplement)

- **Shape name / ID indexes** at `src/pptx/shapes/shapetree.py:152-184`.
  Lazy build on first lookup, dict-cached, invalidated on mutation. Stage 1
  measures lookup linearity to confirm O(1) amortized behavior.
- **Lazy blob loading** at `src/pptx/opc/serialized.py:172-209`.
  `_ZipPkgReader.__getitem__` reads each part on demand via
  `zipfile.ZipFile.read`; only member names load eagerly. Stage 1 measures
  the per-deck memory footprint with `tracemalloc`.

## Confirmed Python-side hot spots (not addressed in Stage 1)

- `TextFrame.paragraphs` at `src/pptx/text/text.py:182-188` and
  `_Paragraph.runs` at `src/pptx/text/text.py:870-873` are plain `@property`,
  not `@lazyproperty`. Each access rebuilds a fresh tuple of proxy objects.
  The `iter1x` vs `iterNx` columns of the baseline isolate this cost.
- `_BaseShapes.__getitem__` / `__iter__` / `__len__` at
  `src/pptx/shapes/shapetree.py:92-113` rebuild `list(_iter_member_elms())`
  on each call.

These are candidate v0.5.0 work items, not Stage 1 deliverables.

## Stage 1 deliverables

- `tests/perf/corpus.py` — synthetic deck generator parametrized by slide
  count, shapes/slide, paragraphs/shape, and image count.
- `tests/perf/conftest.py` — session-scoped `synth_small`, `synth_medium`,
  `synth_large` fixtures.
- `tests/perf/bench_open.py` — cold-open wall-clock.
- `tests/perf/bench_iterate.py` — shape and text iteration; isolates the
  `paragraphs`/`runs` rebuild cost via single-pass vs repeated access.
- `tests/perf/bench_save.py` — save-to-`BytesIO` wall-clock.
- `tests/perf/bench_lookup.py` — index linearity assertion across
  n ∈ {10, 100, 1000} shape counts.
- `tests/perf/run_baseline.py` — standalone runner that writes
  `tests/perf/baseline.json` and appends the "Baseline (Stage 1)" section
  below.
- `pyproject.toml` — registers `@pytest.mark.perf` and adds
  `addopts = "-m 'not perf'"` so default `pytest` skips perf tests.
  Users opt in with `pytest -m perf`.

## Reproducing

```
pytest -m perf                       # opt-in pytest run
python -m tests.perf.run_baseline    # full baseline + cProfile report
```

The runner writes `tests/perf/baseline.json` and appends a
"Baseline (Stage 1)" section to this file containing the current numbers.

## Handoff to v0.5.0

The harness, corpus generator, and baseline JSON schema are designed to be
picked up unchanged by the `v0.5.0` branch when it opens. v0.3.0's
real-world corpus can be added as additional fixtures alongside the
synthetic ones with no change to the bench modules.

## Baseline (Stage 1)

Captured: `2026-05-16T16:16:07.694861+00:00`
Platform: `CPython 3.11.15` on `Linux / x86_64`

### Corpus

| fixture | slides | shapes/slide | paragraphs/shape | images | size |
| --- | --- | --- | --- | --- | --- |
| small | 5 | 5 | 2 | 0 | 31.6 KiB |
| medium | 30 | 15 | 4 | 5 | 71.2 KiB |
| large | 100 | 25 | 5 | 10 | 191.3 KiB |

### Wall-clock (ms)

Open is median of 5 cold opens.
`iter1x` walks `slides -> shapes -> paragraphs -> runs` once per text frame.
`iter10x` repeats the inner walk
10 times to amplify the cost of
`paragraphs`/`runs` rebuilding on every access
(`src/pptx/text/text.py:182-188`, `src/pptx/text/text.py:870-873`).

| fixture | open | iter1x | iter10x | save |
| --- | --- | --- | --- | --- |
| small | 4.77 | 2.21 | 6.82 | 5.55 |
| medium | 13.49 | 55.39 | 198.37 | 13.70 |
| large | 50.55 | 241.08 | 1314.62 | 44.42 |

### Memory on open (KiB, tracemalloc)

| fixture | current | peak |
| --- | --- | --- |
| small | 108.0 | 152.4 |
| medium | 105.7 | 221.9 |
| large | 457.2 | 457.8 |

### cProfile top 20 (medium fixture, open + iterate + save)

```
573149 function calls (571770 primitive calls) in 0.410 seconds

   Ordered by: cumulative time
   List reduced from 331 to 20 due to restriction <20>

   ncalls  tottime  percall  cumtime  percall filename:lineno(function)
        1    0.024    0.024    0.361    0.361 run_baseline.py:68(_bench_iterate_text)
    18000    0.021    0.000    0.142    0.000 text.py:870(runs)
    22610    0.067    0.000    0.121    0.000 xmlchemy.py:397(get_child_element_list)
    18000    0.007    0.000    0.108    0.000 text.py:984(text)
    18000    0.010    0.000    0.101    0.000 text.py:70(text)
    42987    0.034    0.000    0.096    0.000 ns.py:126(qn)
    18060    0.051    0.000    0.091    0.000 xmlchemy.py:491(get_child_element)
     4500    0.005    0.000    0.053    0.000 text.py:182(paragraphs)
    43041    0.025    0.000    0.030    0.000 ns.py:51(__init__)
    36000    0.011    0.000    0.027    0.000 text.py:873(<genexpr>)
        1    0.000    0.000    0.027    0.027 presentation.py:72(save)
        1    0.000    0.000    0.027    0.027 presentation.py:320(save)
        1    0.000    0.000    0.027    0.027 package.py:152(save)
2919/1807    0.003    0.000    0.026    0.000 util.py:166(__get__)
        1    0.000    0.000    0.025    0.025 serialized.py:70(write)
        1    0.000    0.000    0.025    0.025 serialized.py:81(_write)
      485    0.000    0.000    0.025    0.000 shapetree.py:101(__iter__)
      456    0.000    0.000    0.023    0.000 shapetree.py:777(_shape_factory)
      456    0.000    0.000    0.023    0.000 shapetree.py:994(SlideShapeFactory)
        1    0.000    0.000    0.022    0.022 api.py:22(Presentation)
```

### Reproducing

```
python -m tests.perf.run_baseline
pytest -m perf
```
