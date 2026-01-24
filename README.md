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
- **Image Extraction** (v1.8): Extracts images from Excel files and outputs them as Markdown image links
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
- [v1.8/spec.md](v1.8/spec.md) - Technical specification (v1.8 with image extraction)
- [v1.7/spec.md](v1.7/spec.md) - Technical specification (v1.7)

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
uv run python v1.8/excel_to_md.py input.xlsx
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
uv run python v1.8/excel_to_md.py input.xlsx --mermaid-enabled
```

**Generate individual files per sheet:**
```bash
uv run python v1.8/excel_to_md.py input.xlsx --split-by-sheet
```

**Specify CSV markdown output directory:**
```bash
uv run python v1.8/excel_to_md.py input.xlsx --csv-output-dir ./output
# CSV markdown: ./output/input_csv.md
# Images: ./output/input_images/
```

**Output standard Markdown only (no CSV output):**
```bash
uv run python v1.8/excel_to_md.py input.xlsx -o output.md --no-csv-markdown-enabled
```

**Plain text hyperlinks (no Markdown syntax):**
```bash
uv run python v1.8/excel_to_md.py input.xlsx --hyperlink-mode inline_plain
```

**Reduce token count (exclude CSV summary section):**
```bash
uv run python v1.8/excel_to_md.py input.xlsx --no-csv-include-description
```

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

### Image Extraction

Images in Excel files are automatically processed:

1. **Automatic Extraction**: Images from each sheet are saved as external files
   - Filename format: `{sheet_name}_img_{number}.{extension}`
   - Example: `Sheet1_img_1.png`, `Sheet1_img_2.jpg`

2. **Save Location**: Output to same directory as CSV markdown
   - Directory name: `{input_filename}_images/`
   - Example: `input.xlsx` → `input_images/` directory
   - Use `--csv-output-dir` option to change output location

3. **Markdown Links**: Generates Markdown image links for cells with images
   - Format: `![alt text](relative_path)`
   - Uses cell value as alt text if available
   - Auto-generates alt text like `Image at A1` if cell is empty

4. **Supported Formats**: PNG, JPEG, GIF

**Example:**

If a company logo image is at cell position (B2):
- Image file: saved as `input_images/Sheet1_img_1.png`
- CSV output: `![Company Logo](input_images/Sheet1_img_1.png)`
- Cell text "Company Logo" is used as alt text

## Advanced Options

List all options:

```bash
uv run python v1.8/excel_to_md.py --help
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
├── v1.8/                       # Latest version
│   ├── excel_to_md.py          # Main conversion program
│   ├── spec.md                 # Specification
│   └── tests/                  # Test suite
├── v1.7/                       # Previous version
│   ├── excel_to_md.py          # Main conversion program
│   ├── spec.md                 # Specification
│   └── tests/                  # Test suite
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
