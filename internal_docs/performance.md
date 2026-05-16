# Performance: speedups for batch presentation processing

## Problem statement

`pypptx` is a pure-Python library, but it leans heavily on C-extension dependencies (`lxml`, `Pillow`, stdlib `zipfile`). Profiling reasoning (no measured corpus yet — see Phase 1) suggests roughly:

- **~75–80% of wall-clock time** in a typical read-modify-write workload is already in C (lxml parse/serialize, Pillow image decode, zip I/O).
- **~20–25%** is in pure-Python glue: element-proxy attribute access, namespace resolution, type conversion of XML attribute strings, relationship lookup.

For users processing **large amounts of presentations** the relevant bottleneck is *per-file wall-clock × N files*. That framing changes which optimizations matter.

## Is Cython a viable route?

Honest answer: **possible, but it is the lowest-leverage option on the table.** Three structural reasons:

1. The C-extension dependencies are already as fast as Cython would make them. You cannot Cython your way past `libxml2`.
2. The pure-Python descriptor most often cited as a candidate — `lazyproperty` at `src/pptx/util.py:108-214` — carries an explicit upstream note that it costs ~0.4 µs per access and is "probably not a rich target for optimization efforts" (`src/pptx/util.py:210-212`).
3. There is no C-extension build infrastructure yet. `pyproject.toml` is plain `setuptools` + `setuptools-scm`. Introducing Cython adds a compile-time toolchain dependency and complicates wheel builds for every supported platform.

If, after Phases 1–3 below, profiling shows pure-Python overhead is still ≥10% of total runtime, Cython is worth a focused spike on the two surfaces called out in Phase 4.

## Phased plan

### Phase 1 — Profile first (mandatory prerequisite)

No optimization without a benchmark. Deliverables:

- A `tests/perf/` benchmark harness that generates and processes a representative corpus (e.g. 100 decks × 50 slides × representative shape mix).
- `cProfile` baseline (`python -m cProfile -o batch.prof scripts/bench_batch.py`) plus a `pyinstrument` flame view.
- Documented baseline numbers in this file (wall-clock, peak RSS, top-20 by cumulative time).

Acceptance: numbers are reproducible on CI and on a developer laptop.

### Phase 2 — Process-level parallelism (highest ROI)

The library has **no module-level mutable state** and `lxml` custom-element classes pickle cleanly, so per-file work is safe to fan out across processes today, with no library changes required.

Deliverables:

- `examples/batch_process.py` demonstrating `concurrent.futures.ProcessPoolExecutor` over a directory of `.pptx` files.
- README section: "Processing many presentations".
- Phase 1 corpus re-run with N workers; capture wall-clock improvement.

Expected outcome: on a 4-core machine, ~3–3.5× wall-clock speedup over Phase 1 baseline.

### Phase 3 — Low-cost Python-level wins

Run in parallel with Phase 2 once Phase 1 numbers exist:

- **Add `__slots__`** to `ElementProxy`, `ParentedElementProxy`, `PartElementProxy` (`src/pptx/shared.py:13-83`). Proxy objects are instantiated thousands of times per deck; expected ~10–15% memory reduction and a small attribute-access speedup.
- **Compile and cache hot XPath expressions** via `lxml.etree.XPath(...)` instead of `element.xpath(str, namespaces=...)` per call. ~50 call sites under `src/pptx/oxml/`; lift the most-used ones to class-level compiled XPaths cached with `@lazyproperty`. Expected 20–40% local speedup on repeated queries.
- **Expose a `compression` option** on `Presentation.save()` so batch users can choose `ZIP_STORED` when CPU matters more than size. Touch site: `src/pptx/opc/serialized.py:256`.

Acceptance: full test suite green, no public-API breakage, Phase 1 benchmark shows measurable improvement (target ≥10% on the Python-side share).

### Phase 4 — Cython (conditional)

Only proceed if Phases 2–3 leave a residual pure-Python bottleneck ≥10% of total runtime. Candidate surfaces:

- `BaseFloatType.convert_from_xml` / `BaseIntType.convert_from_xml` at `src/pptx/oxml/simpletypes.py:66-97` — every numeric XML attribute access pays a Python `int()`/`float()`.
- `NamespacePrefixedTag` construction at `src/pptx/oxml/ns.py:45-98` — string-split and Clark-name formatting.

Expected per-site gain: 1.2–2×. Translation to total runtime: single-digit %. Build cost: real (compiler toolchain, wheel matrix, source-build fallback). The honest framing: do this only if the profile demands it.

## Critical files

| Path | Why it matters |
| --- | --- |
| `src/pptx/util.py:108-214` | `lazyproperty` descriptor; pervasive caching utility; reuse for XPath caching. |
| `src/pptx/shared.py:13-83` | Proxy bases — `__slots__` target. |
| `src/pptx/oxml/xmlchemy.py` | `xpath()` machinery and element-class registration; XPath compilation site. |
| `src/pptx/oxml/simpletypes.py:66-97` | Per-attribute type conversion; Cython candidate (Phase 4). |
| `src/pptx/oxml/ns.py:45-98` | Namespace/tag resolution; Cython candidate (Phase 4). |
| `src/pptx/opc/serialized.py:256` | ZIP compression flag; expose via API in Phase 3. |
| `pyproject.toml` | Will need `[build-system]` changes only if Phase 4 proceeds. |

## Existing utilities to reuse

- `lazyproperty` (`src/pptx/util.py:108`) — use to memoise compiled XPaths at class scope.
- `register_element_cls` (`src/pptx/oxml/ns.py`) — keep using; do not bypass lxml's element-class lookup with Python-side dispatch.

## Verification

1. `pytest` stays green after each phase.
2. Phase 1 produces a reproducible `cProfile` baseline checked into `tests/perf/`.
3. Phase 2: wall-clock on the Phase 1 corpus drops to ~25–35% of baseline on a 4-core host.
4. Phase 3: re-run profile; `tracemalloc` shows reduced proxy-object footprint; `pyinstrument` shows fewer ticks under `xpath`.
5. Phase 4 is gated on (3) leaving ≥10% residual pure-Python overhead; otherwise close as "not justified by data".
