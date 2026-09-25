# 2026-09-25 16:30 · issue-51-perf-caches

**Goal:** Explore issue #51 (batch performance) and ship the first, measured fix

## What happened
- Profiled a text-heavy batch job (100 slides × 25 text boxes, open → read text → edit runs → save) on main,
  reusing the stale `claude/cython-performance-exploration-VHSxn` corpus generator from scratch space.
- Findings: ~86% of profiled time is pypptx Python (the issue's "75–80% in C" premise doesn't hold for text walks);
  `BaseOxmlElement.xpath` recompiled XPath on every call (per-shape `ph` check alone 0.17 s/deck); `qn()` ran 92k
  times/deck. 4-process fan-out gives 3.4×; zlib is ~9 ms of a 50 ms save; `__slots__` would break `@lazyproperty`
  (uses instance `__dict__`, 199 usages).
- Shipped: `qn` wrapped in `functools.lru_cache`; `xmlchemy.compiled_xpath()` — per-thread, bounded (512, clear on
  full) cache of `etree.XPath` used by `BaseOxmlElement.xpath`. ~658 → ~400 ms/deck (−39%); with 4 processes the
  16-deck batch goes 10.6 s → 2.0 s.

## Open threads
- #51 remaining: rebase the perf harness (`tests/perf/`) for v0.5.0; `examples/batch_process.py`; update the issue's
  checklist (drop `__slots__`, compression kwarg; Cython not justified).
