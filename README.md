# excel2md

[English](./README.md) | [日本語](./README_ja.md)

[![Elvez](https://img.shields.io/badge/Elvez-Product-3F61A7?style=flat-square)](https://elvez.co.jp/)
[![IXV Ecosystem](https://img.shields.io/badge/IXV-Ecosystem-3F61A7?style=flat-square)](https://elvez.co.jp/ixv/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.9+-blue?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Stars](https://img.shields.io/github/stars/elvezjp/excel2md?style=social)](https://github.com/elvezjp/excel2md/stargazers)

Excel to Markdown converter. Reads Excel workbooks (.xlsx/.xlsm) and automatically generates Markdown format output.

## Features

- **Smart Table Detection**: Automatically detects Excel print areas and converts them to Markdown tables
- **CSV Markdown Output**: Exports entire sheets in CSV format with validation metadata
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

- [CHANGELOG.md](CHANGELOG.md) - Version history
- [CONTRIBUTING.md](CONTRIBUTING.md) - Contribution guidelines
- [SECURITY.md](SECURITY.md) - Security policy and best practices
- [v1.7/spec.md](v1.7/spec.md) - Technical specification

## Setup

### Requirements

- Python 3.9 or higher
- [uv](https://docs.astral.sh/uv/) package manager

### Install Dependencies

```bash
# Install uv (if not already installed)
# Details: https://docs.astral.sh/uv/getting-started/installation/
curl -LsSf https://astral.sh/uv/install.sh | sh

uv sync
```

## Usage

```bash
uv run python v1.7/excel_to_md.py input.xlsx -o output.md
```

This generates:
- `output.md` - Standard Markdown table format
- `input_csv.md` - CSV markdown format (enabled by default)

### Common Examples

**Convert with Mermaid flowchart support:**
```bash
uv run python v1.7/excel_to_md.py input.xlsx -o output.md --mermaid-enabled
```

**Generate individual files per sheet:**
```bash
uv run python v1.7/excel_to_md.py input.xlsx -o output.md --split-by-sheet
```

**Output standard Markdown only (no CSV output):**
```bash
uv run python v1.7/excel_to_md.py input.xlsx -o output.md --no-csv-markdown-enabled
```

**Plain text hyperlinks (no Markdown syntax):**
```bash
uv run python v1.7/excel_to_md.py input.xlsx -o output.md --hyperlink-mode inline_plain
```

**Reduce token count (exclude CSV summary section):**
```bash
uv run python v1.7/excel_to_md.py input.xlsx -o output.md --no-csv-include-description
```

## Key Options

### Output Control

| Option | Default | Description |
|--------|---------|-------------|
| `-o`, `--output` | - | Output file path |
| `--split-by-sheet` | false | Generate individual files per sheet |
| `--csv-markdown-enabled` | true | Enable CSV markdown output |
| `--csv-include-description` | true | Include summary section in CSV output |
| `--csv-include-metadata` | true | Include validation metadata in CSV output |

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

## Output Examples

### Standard Markdown Output

```markdown
# Conversion Result: sample.xlsx

- Spec Version: 1.7
- Sheet Count: 2
- Sheet List: Sheet1, Summary

---

## Sheet1

### Table 1 (A1:C4)
| Item | Quantity | Notes |
| --- | ---: | --- |
| Apple | 10 | [Supplier](https://example.com)[^1] |
| Orange | 5 |  |

[^1]: https://example.com
```

### CSV Markdown Output

````markdown
# CSV Output: sample.xlsx

## Summary

### File Information
- Original Excel filename: sample.xlsx
- Sheet count: 2
- Generated at: 2025-01-05 10:00:00

### About This File
This CSV markdown file is designed to help AI understand Excel content...

---

## Sheet1

```csv
Item,Quantity,Notes
Apple,10,Supplier
Orange,5,
```

---

## Validation Metadata

- **Generated at**: 2025-01-05 10:00:00
- **Original Excel file**: sample.xlsx
- **Validation status**: OK
````

## Advanced Options

List all options:

```bash
uv run python v1.7/excel_to_md.py --help
```

Key advanced options:
- Cell merge policy
- Date/number format control
- Whitespace handling
- Markdown escape level
- Hidden row/column policy
- Locale-specific formatting

## Directory Structure

```
excel2md/
├── v1.7/
│   ├── excel_to_md.py      # Main conversion program (latest version)
│   ├── spec.md             # Specification
│   └── tests/              # Test suite
├── pyproject.toml          # Project metadata
├── LICENSE                 # MIT License
├── README.md               # This file
├── CONTRIBUTING.md         # Contribution guide
├── SECURITY.md             # Security policy
└── CHANGELOG.md            # Version history
```

## Security

For security concerns, please see [SECURITY.md](SECURITY.md).

**Key security notes:**
- Only process Excel files from trusted sources
- Use `read_only=True` mode to prevent file modification
- Excel macros are not executed
- Sanitize Markdown output to prevent injection

## Contributing

Contributions are welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for details.

- Report bugs via [GitHub Issues](https://github.com/elvez/excel2md/issues)
- Submit pull requests for improvements
- Follow existing code style
- Add tests for new features

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for details.

## Background

This tool was created during the development of **IXV**, an AI development support tool targeting Japanese development documents and specifications.

IXV addresses challenges in understanding, structuring, and utilizing Japanese documents in system development. This repository publicly shares a portion of that work.

## License

MIT License - See [LICENSE](LICENSE) for details.

## Contact

- **Email**: info@elvez.co.jp
- **Company**: Elvez, Inc.
