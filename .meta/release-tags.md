# Release tags — commits that mark each version's completion

Sessions cannot push git tags, so versions are tagged by hand. Each row is the **merge commit
on `main`** where the version was complete: for versions that bumped the package, the merge
that first carries the new `version =` in `pyproject.toml`; for v0.2.0, which changed docs
only, the merge of its PR. Found by walking `main`'s first-parent history
(`git rev-list --first-parent main`) and reading `pyproject.toml` at each commit.

| Tag | Commit | Date | Merge | Package version at the commit | Notes |
|---|---|---|---|---|---|
| `v0.1.0` | `723ef5aa0d1973844a6a1975df7648db577174cf` | 2026-05-16 | PR #50 | 0.1.0 | First commit carrying 0.1.0 |
| `v0.2.0` | `49108ffc2cbbd6d523e4dd18a093ac8a086a8a03` | 2026-05-16 | PR #53 | 0.1.0 | Docs-only version; the package was never set to 0.2.0 (it went 0.1.0 → 0.3.0) |
| `v0.3.0` | `0658fa51370950394501dc251d4b92a3fe841d7c` | 2026-09-26 | PR #111 | 0.3.0 | Roadmap reconcile + package 0.3.0; the epic work it records merged in earlier PRs |
| `v0.3.1` | `99e2ef72a9fb3d4b29faa4138b9750244e207d23` | 2026-09-27 | PR #117 | 0.3.1 | Marks v0.3.1 shipped and bumps the package; the examples themselves merged in PR #116 (`1657988`) |
| `v0.3.2` | `cb8a179c3346cf75f1759c445df31f1ef244dd51` | 2026-09-30 | PR #121 | 0.3.2 | Marks v0.3.2 shipped and bumps the package; docs work merged in PR #120 (`33a3e02`). Also carries #114 and #115 |
| `v0.4.0` | `aceb359ebfd89df5f51e903b8eabde501554037c` | 2026-10-01 | PR #123 | 0.4.0 | **The ship commit itself, not a merge:** PR #123 also carries #51's P6 (`iterchildren` getters), which is not part of v0.4.0. This commit's parent is the v0.4.0 merge (`2c64b2d`, PR #122), so its tree is exactly v0.4.0 plus the ship bookkeeping |

## Commands

Annotated tags, then one push:

```bash
git fetch origin main
git tag -a v0.1.0 723ef5aa0d1973844a6a1975df7648db577174cf -m "v0.1.0"
git tag -a v0.2.0 49108ffc2cbbd6d523e4dd18a093ac8a086a8a03 -m "v0.2.0 (docs only; package 0.1.0)"
git tag -a v0.3.0 0658fa51370950394501dc251d4b92a3fe841d7c -m "v0.3.0"
git tag -a v0.3.1 99e2ef72a9fb3d4b29faa4138b9750244e207d23 -m "v0.3.1"
git tag -a v0.3.2 cb8a179c3346cf75f1759c445df31f1ef244dd51 -m "v0.3.2"
git tag -a v0.4.0 aceb359ebfd89df5f51e903b8eabde501554037c -m "v0.4.0"
git push origin v0.1.0 v0.2.0 v0.3.0 v0.3.1 v0.3.2 v0.4.0
```

A GitHub Release can then be drafted from each tag; the version cards under
`.meta/roadmap/stages/` hold the goals and acceptance for its notes.

## Keeping this current

When a version ships, add its row once the PR that marks it shipped has merged (its merge
commit is the one to tag). If that PR also carries later work, tag the ship commit itself instead,
as for v0.4.0, so the tag holds only the version. Once the tags exist on the remote, `git ls-remote --tags origin`
is the record and this file can go.
