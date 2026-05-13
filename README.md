# excel2md

[English](https://github.com/elvezjp/excel2md/blob/main/README.md) | [日本語](https://github.com/elvezjp/excel2md/blob/main/README_ja.md)

[![Elvez](https://img.shields.io/badge/Elvez-Product-3F61A7?style=flat-square)](https://elvez.co.jp/)
[![IXV Ecosystem](https://img.shields.io/badge/IXV-Ecosystem-3F61A7?style=flat-square)](https://elvez.co.jp/ixv/)
[![PyPI version](https://img.shields.io/pypi/v/excel2md?style=flat-square)](https://pypi.org/project/excel2md/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.10+-blue?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Stars](https://img.shields.io/github/stars/elvezjp/excel2md?style=social)](https://github.com/elvezjp/excel2md/stargazers)

Excel to Markdown converter. Reads Excel workbooks (.xlsx/.xlsm) and automatically generates Markdown format output.

## Features

- **Smart Table Detection**: Automatically detects Excel print areas and converts them to Markdown tables
- **CSV Markdown Output**: Exports entire sheets in CSV format with validation metadata
- **Image Extraction**: Extracts images from Excel files and outputs them as Markdown image links
- **Mermaid Flowcharts**: Generates Mermaid diagrams from Excel shapes and tables
- **Hyperlink Support**: Multiple output modes (inline, footnote, plain text)
- **Split by Sheet**: Generate individual files per sheet
- **Customizable**: Detailed settings for formatting, alignment, and data processing

## Use Cases

- **Document Generation**: Convert Excel specifications to Markdown
- **AI/LLM Processing**: CSV markdown format optimized for token efficiency
- **Flowchart Extraction**: Extract diagrams from Excel shapes
- **Data Migration**: Export Excel data to portable Markdown format
- **Version Control**: Track Excel changes in text-based format

## Documentation

- [CHANGELOG.md](https://github.com/elvezjp/excel2md/blob/main/CHANGELOG.md) - Version history
- [CONTRIBUTING.md](https://github.com/elvezjp/excel2md/blob/main/CONTRIBUTING.md) - Contribution guidelines
- [SECURITY.md](https://github.com/elvezjp/excel2md/blob/main/SECURITY.md) - Security policy and best practices
- [v2.2.0/spec.md](https://github.com/elvezjp/excel2md/blob/main/v2.2.0/spec.md) - Technical specification (v2.2.0, latest)
- [v2.1.0/spec.md](https://github.com/elvezjp/excel2md/blob/main/v2.1.0/spec.md) - Technical specification (v2.1.0, frozen snapshot)
- [v1.8/spec.md](https://github.com/elvezjp/excel2md/blob/main/v1.8/spec.md) - Technical specification (v1.8)

## Installation

Requires Python 3.10 or higher.

```bash
pip install excel2md
# or with uv
uv add excel2md
```

After installation, the `excel2md` command is available on your `PATH`.

## Usage

```bash
excel2md input.xlsx
```
This generates:
- `input_csv.md`: CSV markdown format (default)
- `input_images/`: Image directory (if images exist)

**Note**
- Output filenames and directories are based on input filename (e.g., `input.xlsx` → `input_csv.md`, `input_images/`)
- Output is saved in the same directory as input file (use `--csv-output-dir` to change)

### Common Examples

**Convert with Mermaid flowchart support:**
```bash
excel2md input.xlsx --mermaid-enabled
```

**Generate individual files per sheet:**
```bash
excel2md input.xlsx --split-by-sheet
```

**Specify CSV markdown output directory:**
```bash
excel2md input.xlsx --csv-output-dir ./output
# CSV markdown: ./output/input_csv.md
# Images: ./output/input_images/
```

**Output standard Markdown only (no CSV output):**
```bash
excel2md input.xlsx -o output.md --no-csv-markdown-enabled
```

**Plain text hyperlinks (no Markdown syntax):**
```bash
excel2md input.xlsx --hyperlink-mode inline_plain
```

**Reduce token count (exclude CSV summary section):**
```bash
excel2md input.xlsx --no-csv-include-description
```

## Use as a Library

`excel2md` is also usable as a Python library.

```python
from excel2md import convert_to_markdown

# Pass a path, or raw xlsx bytes (handy for Pyodide / web uploads)
result = convert_to_markdown("input.xlsx", csv_markdown_enabled=False)

print(result["markdown"])      # Generated Markdown string
print(result["output_path"])   # Where the .md file was written
```

CLI options map 1:1 to keyword arguments (e.g. `mermaid_enabled=True`, `split_by_sheet=True`). For multiple conversions sharing the same configuration, use `ConversionConfig` + `ExcelConverter` directly.


### From source

```bash
git clone https://github.com/elvezjp/excel2md.git
cd excel2md
uv sync
```

See [CONTRIBUTING.md](https://github.com/elvezjp/excel2md/blob/main/CONTRIBUTING.md) for the full developer setup.


## Key Options

### Output Control

| Option | Default | Description |
|--------|---------|-------------|
| `--split-by-sheet` | false | Generate individual files per sheet |
| `--csv-markdown-enabled` | true | Enable CSV markdown output |
| `--csv-output-dir` | Same as input | Output directory for CSV markdown and images |
| `--csv-include-description` | true | Include summary section in CSV output |
| `--csv-include-metadata` | true | Include validation metadata in CSV output |
| `--image-extraction` | true | Enable image extraction |
| `-o`, `--output` | - | Output file path for standard Markdown |

### Hyperlink Formats

| Mode | Description | Output Example |
|------|-------------|----------------|
| `inline` | Markdown format | `[text](URL)` |
| `inline_plain` | Plain text format | `text (URL)` |
| `footnote` | Footnote format | `[text][^1]` + `[^1]: URL` |
| `text_only` | Display text only | `text` |
| `both` | Inline + footnote | Both formats |

### Mermaid Flowcharts

| Option | Default | Description |
|--------|---------|-------------|
| `--mermaid-enabled` | false | Enable Mermaid conversion |
| `--mermaid-detect-mode` | shapes | Detection mode: `shapes`, `column_headers`, `heuristic` |
| `--mermaid-direction` | TD | Flowchart direction: `TD`, `LR`, `BT`, `RL` |
| `--mermaid-keep-source-table` | true | Output original table along with Mermaid |

### Table Processing

| Option | Default | Description |
|--------|---------|-------------|
| `--header-detection` | first_row | Treat first row as header |
| `--align-detection` | numbers_right | Right-align numeric columns |
| `--max-cells-per-table` | 200000 | Maximum cells per table |
| `--no-print-area-mode` | used_range | Behavior when print area not set |


### Advanced Options

List all options:

```bash
excel2md --help
```

Key advanced options:
- Cell merge policy
- Date/number format control
- Whitespace handling
- Markdown escape level
- Hidden row/column policy
- Locale-specific formatting


## Output Examples

Real input / output samples (including images) live under [docs/examples/](https://github.com/elvezjp/excel2md/tree/main/docs/examples). Each version directory contains:

- Input `.xlsx` files
- `output-default/` — default mode (CSV markdown + image extraction)
- `output-markdown/` — standard Markdown mode (`--no-csv-markdown-enabled`)
- `output-mermaid/` — Mermaid flowchart enabled (`--mermaid-enabled`)

The regeneration commands for each pattern are documented in [docs/examples/README.md](https://github.com/elvezjp/excel2md/blob/main/docs/examples/README.md).


## Directory Structure

```
excel2md/
├── v2.2.0/                     # Latest version
│   ├── excel_to_md.py          # Entry point
│   ├── excel2md/               # Main package
│   ├── tests/                  # Test suite
│   ├── spec.md                 # Specification
│   └── spec_appendix.md        # Specification appendix
├── v2.1.1/                     # Previous version (frozen snapshot, not published to PyPI)
├── v2.1.0/                     # Previous version (frozen snapshot)
├── v2.0.1/                     # Previous version
├── v2.0/                       # Previous version
├── v1.8/                       # Legacy version
│   ├── excel_to_md.py          # Main conversion program
│   ├── spec.md                 # Specification
│   └── tests/                  # Test suite
├── v1.7/                       # Legacy version
│   ├── excel_to_md.py          # Main conversion program
│   ├── spec.md                 # Specification
│   └── tests/                  # Test suite
├── docs/                   # Documentation
├── pyproject.toml          # Project metadata
├── LICENSE                 # MIT License
├── README.md / _ja.md     # README (English / Japanese)
├── CONTRIBUTING.md / _ja.md # Contribution guide (English / Japanese)
├── SECURITY.md / _ja.md   # Security policy (English / Japanese)
└── CHANGELOG.md / _ja.md  # Version history (English / Japanese)
```

## Security

For security concerns, please see [SECURITY.md](https://github.com/elvezjp/excel2md/blob/main/SECURITY.md).

**Key security notes:**
- Only process Excel files from trusted sources
- Use `read_only=True` mode to prevent file modification
- Excel macros are not executed
- Sanitize Markdown output to prevent injection

## Contributing

Contributions are welcome! See [CONTRIBUTING.md](https://github.com/elvezjp/excel2md/blob/main/CONTRIBUTING.md) for details.

- Report bugs via [GitHub Issues](https://github.com/elvezjp/excel2md/issues)
- Submit pull requests for improvements
- Follow existing code style
- Add tests for new features

## Changelog

See [CHANGELOG.md](https://github.com/elvezjp/excel2md/blob/main/CHANGELOG.md) for details.

## Background

This tool was created during the development of **IXV**, an AI development support tool targeting Japanese development documents and specifications.

IXV addresses challenges in understanding, structuring, and utilizing Japanese documents in system development. This repository publicly shares a portion of that work.

## License

MIT License - See [LICENSE](https://github.com/elvezjp/excel2md/blob/main/LICENSE) for details.

## Contact

- **Email**: info@elvez.co.jp
- **Company**: Elvez, Inc.
