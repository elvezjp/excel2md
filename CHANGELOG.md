# Changelog

[English](./CHANGELOG.md) | [日本語](./CHANGELOG_ja.md)

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.2.1] - 2026-05-14

### Added
- **Library-level exception types: `ExcelConversionError` (base) and `WorkbookOpenError`** ([#36](https://github.com/elvezjp/excel2md/issues/36))
  - Exposed from both `excel2md` and `excel_to_md` so library consumers can catch conversion errors without inspecting strings

### Changed
- **`load_workbook_safe()` no longer calls `sys.exit(2)` on open failures; it now raises `WorkbookOpenError` instead** ([#36](https://github.com/elvezjp/excel2md/issues/36))
  - Previously, passing a bad path or corrupt bytes to `ExcelConverter` / `convert_to_markdown` terminated the caller's entire process — unacceptable for the PyPI-published reusable API (Pyodide / MCP servers / notebooks / web services)
  - The original openpyxl exception is preserved via `raise ... from e` (accessible as `__cause__`)
  - CLI behavior is unchanged: `cli.main()` now catches `ExcelConversionError`, prints `[ERROR] ...` to stderr, and exits with code 2 — same user experience as before
- Bumped `excel2md.__version__` to `"2.2.1"` (was still `"2.2.0"` after the v2.2.0 → v2.2.1 directory copy)

### Fixed
- **Fixed CSV markdown leaking content outside the declared print area when out-of-area images existed** ([#14](https://github.com/elvezjp/excel2md/issues/14))
  - `runner.run()` expanded `union_area` to cover every image position reported by `extract_images_from_sheet`, so an image placed outside the print area dragged the CSV output range with it and pulled in unrelated cell values
  - The expansion contradicted both the v2.1.0 spec (§"印刷領域内のみが変換対象となる") and the v1.8-and-earlier behavior
  - Removed the expansion block; `extract_print_area_for_csv()` iterates strictly within the declared print area and out-of-area `cell_to_image` entries are now ignored
  - Note: `extract_images_from_sheet` is still invoked, so image files outside the print area may still be written to disk. That side effect is tracked separately and out of scope for this fix.

### Changed
- Added new active development directory `v2.2.1/`. `v2.2.0/` is preserved as a frozen snapshot at the v2.2.0 release point, per the repository versioning policy (existing `v*/` directories are frozen; new versions go into a new directory)
- **Repository layout: moved past version directories (`v1.7`, `v1.8`, `v2.0`, `v2.0.1`, `v2.1.0`, `v2.1.1`, `v2.2.0`) under `versions/`** ([#11](https://github.com/elvezjp/excel2md/issues/11))
  - The active development directory (`v2.2.1/`) remains at the repository root
  - `pyproject.toml`'s sdist `exclude` was consolidated to `versions/**`; the wheel/sdist contents shipped to PyPI are unchanged
  - All moves are recorded as git renames, preserving history
- Added `versions/README.md` describing the purpose of the directory and that its contents are not published to PyPI

### Tests
- Added `TestIssue14CsvPrintAreaRespect` to `tests/test_runner_regression.py` covering an A1:B2 print area with an out-of-area image at (5, 5) and out-of-area data cells — asserts the declared range stays `A1:B2` and that neither the outside image link nor the out-of-area cell values leak into the CSV output
- **Added test coverage for `.xlsm` (macro-enabled workbook) support** ([#43](https://github.com/elvezjp/excel2md/issues/43))
  - Added `v2.2.1/tests/fixtures/test_macro.xlsm` (basic table + a sheet containing VBA macros)
  - Added `v2.2.1/tests/test_xlsm_support.py` — regression-checks that the openpyxl default (`keep_vba=False`) discards the VBA binary, that the project's `workbook_loader.load_workbook_safe` opens `.xlsm`, that both path and bytes inputs convert successfully, that the basic table parity-matches an equivalent `.xlsx`, and that no macro-induced side effects appear during conversion

### Documentation
- Documented `.xlsm` support scope (read-only, VBA discarded, macros not executed) and the new fixture in `v2.2.1/spec.md` §3.1.1 / §11.3 ([#43](https://github.com/elvezjp/excel2md/issues/43))
- Clarified in `SECURITY.md` / `SECURITY_ja.md` that `.xlsm` files are loaded via `openpyxl.load_workbook(..., keep_vba=False)`, so the VBA binary is discarded and auto-run macros such as `Auto_Open` / `Workbook_Open` are never triggered ([#43](https://github.com/elvezjp/excel2md/issues/43))
- Added v2.2.1 example outputs under `docs/examples/v2.2.1/` (samples for `test_standard.xlsx`, `test_mermaid.xlsx`, and `test_macro.xlsm`) and updated `docs/examples/README.md` with the `.xlsm` generation commands

## [2.2.0] - 2026-05-13

### Added
- **Library API surface**: `ConversionConfig`, `ExcelConverter`, and `convert_to_markdown` are now exposed from both `excel2md` and `excel_to_md` ([#16](https://github.com/elvezjp/excel2md/issues/16))
  - `convert_to_markdown(data: bytes | str | Path, **opts) -> dict` — one-shot convenience for non-CLI callers (Pyodide / MCP servers / notebooks / web services)
  - `ConversionConfig` — type-hinted dataclass mirroring CLI options, with `from_args()` and `to_opts_dict()` helpers
  - `ExcelConverter` — reusable class wrapping config + bytes-in/dict-out conversion
- Pure Python implementation works inside Pyodide via `micropip.install('excel2md')`; no native dependencies
- GitHub Actions Trusted Publisher workflows for PyPI and TestPyPI publishing ([#19](https://github.com/elvezjp/excel2md/issues/19))
  - `.github/workflows/publish.yml` — tag-triggered production release
  - `.github/workflows/publish-testpypi.yml` — manual TestPyPI rehearsal
  - See [docs/20260513_pypi_trusted_publisher_setup.md](docs/20260513_pypi_trusted_publisher_setup.md) for administrator setup
- **First public release on PyPI as `excel2md`**
- Registered the standalone verification CLI as `excel2md-verify` under `project.scripts` ([#38](https://github.com/elvezjp/excel2md/issues/38))

### Fixed
- **Fixed missing `verify_csv_markdown` module that silently disabled CSV Markdown verification metadata** ([#38](https://github.com/elvezjp/excel2md/issues/38))
  - `csv_export.py` imported the sibling file `v2.2.0/verify_csv_markdown.py` via a `sys.path` hack, but the module was not included in the wheel/sdist
  - In development it only resolved because Python automatically prepends the cwd to `sys.path` when running scripts; after `pip install` users hit `[WARN] Failed to append verification metadata: No module named 'verify_csv_markdown'` and the verification metadata block was dropped
  - Moved `verify_csv_markdown.py` into `v2.2.0/excel2md/` and switched to a relative import (`from .verify_csv_markdown import ...`)

### Changed
- `runner.run()` internally normalizes both `argparse.Namespace` (CLI) and `ConversionConfig` (library) inputs. CLI behavior is unchanged; the inline options dictionary was replaced with `ConversionConfig.to_opts_dict()`, with a roundtrip test guaranteeing exact equivalence
- `pyproject.toml` `authors` now includes `email = "info@elvez.co.jp"` for PyPI metadata
- Added new active development directory `v2.2.0/`. `v2.1.1/` is preserved as a frozen snapshot at the v2.1.1 release point (commit 034fa57), per the repository versioning policy (existing `v*/` directories are frozen; new versions go into a new directory)

### Notes
- `v2.1.1/` was assigned an internal version number only and was never published to PyPI; `v2.2.0` is the first PyPI release
- The frozen `v2.1.1/` snapshot does **not** contain the library API (`ConversionConfig` / `ExcelConverter` / `convert_to_markdown`) introduced in this release — those are only present in `v2.2.0/`

## [2.1.1] - 2026-05-11

### Fixed
- **Restored backward-compatible re-exports of `is_code_block` and `build_code_block_from_rows`** ([#15](https://github.com/elvezjp/excel2md/issues/15))
  - Both symbols moved into `excel2md.table_formatting` in v2.0 and stopped being importable from the top-level `excel2md` / `excel_to_md` modules
  - Re-exported from `excel2md/__init__.py` and `excel_to_md.py` to restore the v1.8 public API surface
- **Fixed inconsistent return arity from `extract_table()` on truncation path** ([#24](https://github.com/elvezjp/excel2md/issues/24))
  - The `max_cells_per_table` truncation branch returned a 3-tuple while `runner.run()` unpacked a 4-tuple, raising `ValueError`
  - Truncation path now returns the 4-tuple `(md_rows, note_refs, True, table_title)`
- **Fixed duplicated footnote numbering across multiple tables** ([#25](https://github.com/elvezjp/excel2md/issues/25))
  - `runner.run()` passed the same `global_footnote_start` to every table, causing `[^1]` to be re-issued and references to become ambiguous
  - Now advances the start counter by `len(note_refs)` after each table, so `footnote_scope=book` numbers sequentially across the workbook and `footnote_scope=sheet` resets per sheet
- **Fixed sheet-scope footnote definitions being dropped in non-`split-by-sheet` mode**
  - When `footnote_scope=sheet` was used without `--split-by-sheet`, per-sheet footnote definitions were not emitted
  - Now emits them at the end of each sheet's section as expected

### Tests
- Added `tests/test_public_api.py` to lock in the v1.x re-export surface
- Added `tests/test_runner_regression.py` covering the truncation and footnote-numbering regressions

### Notes
- `v2.1.0/` is preserved as a frozen release snapshot. All fixes above live under `v2.1.1/`.

## [2.1.0] - 2026-04-17

### Changed
- **Raised minimum supported Python version to 3.10**
  - Python 3.9 has reached end-of-life (2025-10) and is no longer supported
  - `requires-python` updated to `>=3.10`
  - CI matrix updated to test against minimum (3.10) and latest (3.14) Python versions

### Security
- **Updated pytest to 9.0.3** ([CVE-2025-71176](https://github.com/advisories/GHSA-6w46-j5rx-g56g))
  - Fixes vulnerable tmpdir handling in pytest
- **Updated Pygments to 2.20.0** ([CVE-2026-4539](https://github.com/advisories/GHSA-5239-wwwm-4pmq))
  - Fixes ReDoS caused by inefficient regex for GUID matching

### Documentation
- Updated spec.md / spec_appendix.md headers to v2.1
- Updated README.md / README_ja.md Python badge and path references

## [2.0.1] - 2026-04-16

### Fixed
- **Fixed missing `is_code_block` import in mermaid_generator.py** ([#13](https://github.com/elvezjp/excel2md/issues/13))
  - Fixed `NameError` occurring in heuristic detection mode
  - Added `from .table_formatting import is_code_block`

- **Resolved duplicate `import re`** ([#13](https://github.com/elvezjp/excel2md/issues/13))
  - Removed duplicate `import re` and `import re as _re` (leftover from v1.8 migration)
  - Unified to `_re`

### Documentation
- Fixed module dependency diagram in specification (spec.md)
- Added code block exclusion note to heuristic detection mode conditions in spec.md

## [2.0.0] - 2026-01-26

### Changed
- **Modularized codebase**
  - Split single implementation file into feature-based modules
  - Restructured as `excel2md/` package

### Documentation
- Reorganized specification structure
- Separated details and supplements into appendix

### Tests
- Added module-level test suites

### Compatibility
- Maintained feature compatibility with v1.8

## [1.8.0] - 2026-01-24

### Added
- **Image extraction**
  - Automatically extracts images from Excel files as external files
  - Image filename format: `{sheet_name}_img_{number}.{extension}`
  - Save location: subdirectory based on Markdown filename
  - Supported formats: PNG, JPEG, GIF
  - Automatic image format detection (format attribute or magic byte detection)
  - Automatic cell position identification (TwoCellAnchor, OneCellAnchor support)
  - Automatic Markdown link generation: `![alt text](relative_path)`
  - Image links work in CSV Markdown mode
  - Uses cell value as alt text (generates cell reference if empty)
  - Graceful error handling (skips image and continues)

### Tests
- Added comprehensive unit tests for image extraction (18 test cases)
  - Image format detection tests (PNG, JPEG, GIF)
  - Anchor position extraction tests (TwoCellAnchor, OneCellAnchor)
  - Integration tests with CSV extraction
  - Error handling and edge case tests
  - Integration tests using actual openpyxl worksheets

### Documentation
- Added image extraction feature description to README.md
  - Added image extraction to feature list
  - Added usage examples and output examples
  - Added detailed behavior description of image extraction
- Added technical specifications to spec.md (v1.7)
  - Added section 7.8: Image Extraction and Markdown Link Generation
  - Documented image processing flow, format detection, and error handling

### Code Quality
- Compliant with PEP 8 style guidelines
- Added comprehensive docstrings (PEP 257 compliant)
- Added detailed inline comments for complex logic
- Used more descriptive variable names (ext -> file_extension, etc.)

## [1.7.0] - 2025-12-25

### Added
- **Mermaid output support in CSV Markdown**
  - `--mermaid-enabled` option now works with CSV Markdown
  - Only supported for `mermaid_detect_mode="shapes"` (extracts flowcharts from Excel shapes)
  - `column_headers` / `heuristic` modes are not supported in CSV Markdown (outputs WARN log and skips)
  - Outputs Mermaid code block immediately after each sheet's CSV block

- **Description section exclusion option**
  - Added `--csv-include-description` / `--no-csv-include-description` options
  - Allows exclusion of CSV Markdown description section
  - Supports token count reduction when converting and combining multiple files
  - Default is `true` (outputs description section as before)

### Changed
- Maintained backward compatibility with v1.6

## [1.6.0] - 2025-11-18

### Added
- **Hyperlink plain text output mode (inline_plain)**
  - Added `--hyperlink-mode inline_plain` option
  - Outputs cell hyperlinks in plain text format: `display text (URL)`
  - For internal links: `display text (-> location)`
  - Explicitly displays link information without Markdown syntax

- **Split by sheet output**
  - Added `--split-by-sheet` option
  - Outputs each sheet as an individual Markdown file
  - Filename format: `{output_filename}_{sheet_name}.md`
  - Each sheet file includes sheet name, spec version, and source filename
  - Uses independent footnote numbering per sheet

### Changed
- Maintained backward compatibility with v1.5

## [1.5.0] - 2025-11-11

### Added
- **CSV Markdown output (enabled by default)**
  - Filename format: `{basename}_csv.md`
  - Each sheet's print area as CSV code blocks
  - Auto-generated summary section and validation metadata section
  - Converts cell line breaks to spaces (guarantees 1 record = 1 line)
  - Outputs display text only for hyperlinks

- **Batch processing support**
  - Updated `batch_test.py` for v1.5
  - Added CSV Markdown output statistics display

- **New options**
  - `--csv-markdown-enabled` / `--no-csv-markdown-enabled`: Enable/disable CSV Markdown output
  - `--csv-output-dir`: Output directory for CSV Markdown
  - `--csv-include-metadata` / `--no-csv-include-metadata`: Include validation metadata
  - `--csv-apply-merge-policy` / `--no-csv-apply-merge-policy`: Apply merge_policy during CSV extraction
  - `--csv-normalize-values` / `--no-csv-normalize-values`: Apply numeric normalization to CSV values

### Changed
- Maintained backward compatibility with v1.4

## [1.4.0] - 2025-11-08

### Added
- **Mermaid flowchart conversion**
  - Column name-based detection: Detects `From` / `To` / `Label` columns for flowchart generation
  - Heuristic detection: Automatic detection from table structure
  - Shape detection: Extracts flowcharts from Excel DrawingML shapes
  - Automatic node ID generation, duplicate edge removal, subgraph support

- **New options**
  - `--mermaid-enabled`: Enable Mermaid flowchart conversion
  - `--mermaid-detect-mode`: Detection mode (`shapes` / `column_headers` / `heuristic`)
  - `--mermaid-direction`: Flowchart direction (`TD` / `LR` / `BT` / `RL`)
  - `--mermaid-keep-source-table`: Output original table alongside Mermaid

## [1.3.0] - 2025-11-08

### Added
- **Core implementation**
  - Maximal rectangle decomposition algorithm (histogram method + carving method)
  - Print area and empty cell detection
  - Merged cell and empty detection
  - Markdown output (table format)
  - Hyperlink processing (footnote format)
  - Performance optimization and limits

- **Basic options**
  - `-o`, `--output`: Output file path
  - `--header-detection`: Treat first row as header
  - `--align-detection`: Right-align numeric columns (80% rule)
  - `--no-print-area-mode`: Behavior when print area not set
  - `--max-cells-per-table`: Maximum cells per table
  - `--markdown-escape-level`: Markdown symbol escape level
  - `--hyperlink-mode`: Hyperlink output method
  - `--footnote-scope`: Footnote numbering scope

### Technical Details
- Supports Python 3.9 and above
- Uses openpyxl 3.1.5+ as dependency
- Safe file reading with `read_only=True, data_only=True` mode

## Links

- [Repository](https://github.com/elvezjp/excel2md)
- [Issues](https://github.com/elvezjp/excel2md/issues)

---

## Version Comparison

| Version | Key Features |
|---------|-------------|
| 2.1.1   | Bug fixes: v1.x re-exports (#15), truncation arity (#24), footnote numbering (#25) |
| 2.1.0   | Raised minimum Python to 3.10, security updates (pytest, Pygments) |
| 2.0.1   | Bug fix in mermaid_generator.py (missing import, duplicate resolution) |
| 2.0.0   | Codebase modularization |
| 1.8.0   | Image extraction (extract images from Excel as external files) |
| 1.7.0   | CSV Markdown mode extensions (Mermaid output, description exclusion) |
| 1.6.0   | Hyperlink plain text output, split by sheet output |
| 1.5.0   | CSV Markdown output |
| 1.4.0   | Mermaid flowchart conversion |
| 1.3.0   | Core implementation |
