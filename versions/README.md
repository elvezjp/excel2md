# versions/

[English](README.md) | [日本語](README_ja.md)

Frozen historical snapshots of past excel2md releases.

Each `vX.Y.Z/` (or `vX.Y/`) subdirectory holds the source tree that shipped with
that release. Directories under this folder are **not** built or published from
the current `main`; the active development source lives at the repository root
under `vX.Y.Z/` (matching `pyproject.toml`'s `version` field) and is the only
tree included in the wheel/sdist on PyPI.

## Why these are kept

- Historical reference and bisecting
- Auditability of past releases
- Following the existing versioning policy that froze old `v*/` directories

## Notes

- Snapshots are intentionally read-only. Bug fixes go into the active version
  directory, not here.
- The `pyproject.toml` excludes `versions/**` from the sdist, so these files do
  not ship to PyPI.
- For the published release history, see
  [CHANGELOG.md](../CHANGELOG.md).
