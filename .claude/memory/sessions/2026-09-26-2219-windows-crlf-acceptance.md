# 2026-09-26 22:19 · windows-crlf-acceptance

**Goal:** Verify the 77 pulled commits' acceptance on Windows; fix the CRLF checkout that broke 51 tests

## What happened
- Fast-forwarded local `main` 77 commits (`bcb7789` → `0658fa5`, v0.3.0). Ran the v0.3.0 card's
  acceptance checks on this Windows machine: `pptx.__version__` and `pyproject.toml` both `0.3.0`;
  all 15 runnable `examples/` scripts exit 0 (`batch_process.py` needs dir args, skipped).
- First `pytest -q` here: **38 failed, 3605 passed, 14 errors** — not a regression. Root cause was
  the working tree, proven before any fix: `git ls-files --eol` → `i/lf w/crlf`, and
  `git show HEAD:<f> | od -c` vs worktree bytes → `\n` vs `\r\n`. Git for Windows carries
  `core.autocrlf=true` in its *system* config and the repo had no `.gitattributes`.
- Added `.gitattributes` (`* text=auto eol=lf` + explicit `binary` list), renormalized the working
  tree, fixed one non-portable test. `pytest -q` → **3657 passed, 0 failed**.
- Branch `claude/windows-eol-acceptance`, commit `ca1b627`, **PR #112** (open, not merged).

## Gotchas & dead ends
- **Why byte-exactness bites here:** `tests/unitutil/file.py` loaders open `"rb"`; `snippet_seq()`
  splits on `"\n\n"`, which under CRLF becomes `"\r\n\r\n"` and returns ONE giant snippet — that is
  why those 14 cases were *fixture setup errors*, not assertion failures. `tests/opc/test_serialized.py`
  SHA-1s a file out of `tests/test_files/expanded_pptx/`, so any CR changes the hash.
- **`dummy.mp4` and `cdw-logo.eps` were corrupt on disk** (43 vs 42 bytes; 22238 vs 21818 = 420
  injected CRs) — both are stored as *text* in the index, so git's heuristic mis-detects them. No
  test hashes either file, which is the only reason it went unnoticed. Now `binary`.
- **`git status` lies after an EOL conversion.** Its dirty check compares the *size recorded in the
  index*, not content, so it reported 480 modified files while `git diff` correctly showed none.
  `git update-index --refresh` prints "needs update" and exits 1; the fix is
  `git add --renormalize .` (same blobs, fresh stat data). Anyone with an existing Windows clone
  needs that one command after pulling #112.
- `git ls-files --eol` output is NOT whitespace-separable: the attr column contains spaces, the path
  follows a **tab**. Splitting on whitespace produced `OSError: Invalid argument` on the path.
- Isolating the two causes: extracting the same commit LF-only with
  `git -c core.autocrlf=false checkout-index -a --prefix=<tmp>/` and running pytest there gave
  3656 passed / 1 failed — separating the CRLF failures from the real test defect in one run.

## State at end
- `main` at `0658fa5` (synced). Working branch `claude/windows-eol-acceptance` pushed; **PR #112
  awaiting review** — nothing merged, `.meta/version` untouched (still v0.3.0 shipped, cursor v0.3.1).
- `pytest -q` is now genuinely `agent-runnable` on Windows: 3657 passed, 0 failed.
- No work started on the v0.3.1 cursor (example scripts).

## Open threads
- **PR #112 needs review/merge.** Until then a fresh Windows clone still fails 52 tests.
- **No CI runs the tests at all** — `.github/workflows/` holds only `epic-autoclose.yml` (ubuntu).
  Every "green at merge" claim in `.meta/version` came from an agent's Linux checkout, which is why
  this was invisible. A `pytest` matrix over ubuntu/windows/macOS would also satisfy the CI gate
  named in the issue #51 thread.
- `verification-surface.md` row for `pytest -q` should get a Windows `last-run` stamp once #112 merges.
- `_os_x_font_directories()` has never run on a real Mac; the test asserts POSIX semantics through
  an injected `posixpath` mock.
