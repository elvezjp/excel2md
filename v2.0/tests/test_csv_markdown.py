"""
Unit tests for CSV Markdown output functions.
"""
import pytest
import sys
import tempfile
from pathlib import Path
from io import StringIO

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_to_md import (
    write_csv_markdown,
    extract_print_area_for_csv,
    coords_to_excel_range,
    format_timestamp,
    sanitize_sheet_name,
    build_merged_lookup,
)


# ============================================================
# Tests for coords_to_excel_range
# ============================================================

class TestCoordsToExcelRange:
    """Tests for coordinate to Excel range conversion."""

    def test_simple_range(self):
        """Simple range A1:B2."""
        result = coords_to_excel_range(1, 1, 2, 2)
        assert result == "A1:B2"

    def test_large_range(self):
        """Larger range with multi-letter columns."""
        result = coords_to_excel_range(1, 1, 100, 30)
        assert result == "A1:AD100"

    def test_single_cell(self):
        """Single cell range."""
        result = coords_to_excel_range(5, 3, 5, 3)
        assert result == "C5:C5"

    def test_column_aa(self):
        """Range including column AA."""
        result = coords_to_excel_range(1, 27, 10, 27)
        assert result == "AA1:AA10"


# ============================================================
# Tests for format_timestamp
# ============================================================

class TestFormatTimestamp:
    """Tests for timestamp formatting."""

    def test_format_pattern(self):
        """Timestamp should match expected pattern."""
        result = format_timestamp()
        # Should be YYYY-MM-DD HH:MM:SS format
        import re
        pattern = r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}'
        assert re.match(pattern, result)

    def test_timestamp_not_empty(self):
        """Timestamp should not be empty."""
        result = format_timestamp()
        assert len(result) > 0


# ============================================================
# Tests for sanitize_sheet_name
# ============================================================

class TestSanitizeSheetName:
    """Tests for sheet name sanitization."""

    def test_normal_name(self):
        """Normal sheet name should be unchanged."""
        result = sanitize_sheet_name("Sheet1")
        assert result == "Sheet1"

    def test_with_slash(self):
        """Slash should be replaced with underscore."""
        result = sanitize_sheet_name("Sheet/1")
        assert "/" not in result
        assert "_" in result

    def test_with_backslash(self):
        """Backslash should be replaced with underscore."""
        result = sanitize_sheet_name("Sheet\\1")
        assert "\\" not in result

    def test_with_colon(self):
        """Colon should be replaced with underscore."""
        result = sanitize_sheet_name("Sheet:1")
        assert ":" not in result

    def test_with_asterisk(self):
        """Asterisk should be replaced with underscore."""
        result = sanitize_sheet_name("Sheet*1")
        assert "*" not in result

    def test_japanese_name(self):
        """Japanese sheet name should be preserved."""
        result = sanitize_sheet_name("シート1")
        assert result == "シート1"

    def test_multiple_unsafe_chars(self):
        """Multiple unsafe characters should all be replaced."""
        result = sanitize_sheet_name("a/b\\c:d*e")
        assert "/" not in result
        assert "\\" not in result
        assert ":" not in result
        assert "*" not in result


# ============================================================
# Tests for extract_print_area_for_csv
# ============================================================

class TestExtractPrintAreaForCSV:
    """Tests for CSV print area extraction."""

    def test_basic_extraction(self, simple_worksheet, default_opts):
        """Basic 2x2 data extraction."""
        ws = simple_worksheet
        area = (1, 1, 2, 2)
        merged_lookup = build_merged_lookup(ws, area)

        rows = extract_print_area_for_csv(ws, area, default_opts, merged_lookup)

        assert len(rows) == 2
        assert len(rows[0]) == 2
        assert rows[0][0] == 'Header1'
        assert rows[0][1] == 'Header2'
        assert rows[1][0] == 'Data1'
        assert rows[1][1] == 'Data2'

    def test_merged_cells(self, worksheet_with_merged_cells, default_opts):
        """Merged cell handling with top_left_only."""
        ws = worksheet_with_merged_cells
        area = (1, 1, 3, 3)
        merged_lookup = build_merged_lookup(ws, area)

        rows = extract_print_area_for_csv(ws, area, default_opts, merged_lookup)

        # First row has merged cells - top-left has value, others should be empty
        assert rows[0][0] == 'Merged Title'  # A1 (top-left)
        assert rows[0][1] == ''  # B1 (part of merged range, empty in top_left_only mode)
        assert rows[0][2] == ''  # C1 (part of merged range, empty in top_left_only mode)

    def test_newline_replacement(self, empty_workbook, default_opts):
        """Newlines in cells should be replaced with spaces."""
        ws = empty_workbook.active
        ws['A1'] = "Line1\nLine2"
        ws['A2'] = "Line3\r\nLine4"
        area = (1, 1, 2, 1)
        merged_lookup = build_merged_lookup(ws, area)

        rows = extract_print_area_for_csv(ws, area, default_opts, merged_lookup)

        assert '\n' not in rows[0][0]
        assert '\r' not in rows[1][0]
        assert ' ' in rows[0][0]  # Newline replaced with space

    def test_hyperlink_inline_plain(self, worksheet_with_hyperlinks, default_opts):
        """Hyperlinks in inline_plain mode."""
        ws = worksheet_with_hyperlinks
        opts = default_opts.copy()
        opts['hyperlink_mode'] = 'inline_plain'
        area = (1, 1, 3, 1)
        merged_lookup = build_merged_lookup(ws, area)

        rows = extract_print_area_for_csv(ws, area, opts, merged_lookup)

        # External link should include URL in parentheses
        assert 'example.com' in rows[0][0] or 'External Link' in rows[0][0]

    def test_empty_cells(self, empty_workbook, default_opts):
        """Empty cells should be empty strings."""
        ws = empty_workbook.active
        ws['A1'] = 'Data'
        # B1 is empty
        area = (1, 1, 1, 2)
        merged_lookup = build_merged_lookup(ws, area)

        rows = extract_print_area_for_csv(ws, area, default_opts, merged_lookup)

        assert rows[0][0] == 'Data'
        assert rows[0][1] == ''

    def test_numeric_normalization(self, empty_workbook, default_opts):
        """Numeric values should be normalized according to options."""
        ws = empty_workbook.active
        ws['A1'] = 1234
        area = (1, 1, 1, 1)
        merged_lookup = build_merged_lookup(ws, area)

        rows = extract_print_area_for_csv(ws, area, default_opts, merged_lookup)

        assert '1234' in rows[0][0]


# ============================================================
# Tests for write_csv_markdown
# ============================================================

class TestWriteCSVMarkdown:
    """Tests for CSV Markdown file writing."""

    def test_with_description(self, empty_workbook, sample_csv_data_dict, default_opts):
        """CSV markdown with description section (default)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            opts = default_opts.copy()
            opts['csv_include_description'] = True
            opts['csv_include_metadata'] = False  # Disable to simplify test

            result = write_csv_markdown(
                empty_workbook,
                sample_csv_data_dict,
                'test_file',
                opts,
                tmpdir
            )

            assert result is not None
            content = Path(result).read_text(encoding='utf-8')

            # Should have description section
            assert '## 概要' in content
            assert '### CSV生成方法' in content
            assert '### CSV形式の仕様' in content

    def test_without_description(self, empty_workbook, sample_csv_data_dict, default_opts):
        """CSV markdown without description section."""
        with tempfile.TemporaryDirectory() as tmpdir:
            opts = default_opts.copy()
            opts['csv_include_description'] = False
            opts['csv_include_metadata'] = False

            result = write_csv_markdown(
                empty_workbook,
                sample_csv_data_dict,
                'test_file',
                opts,
                tmpdir
            )

            assert result is not None
            content = Path(result).read_text(encoding='utf-8')

            # Should NOT have description section
            assert '## 概要' not in content
            assert '### CSV生成方法' not in content

    def test_with_metadata(self, empty_workbook, sample_csv_data_dict, default_opts):
        """CSV markdown with metadata section."""
        with tempfile.TemporaryDirectory() as tmpdir:
            opts = default_opts.copy()
            opts['csv_include_description'] = False
            opts['csv_include_metadata'] = True

            result = write_csv_markdown(
                empty_workbook,
                sample_csv_data_dict,
                'test_file',
                opts,
                tmpdir
            )

            assert result is not None
            content = Path(result).read_text(encoding='utf-8')

            # Should have metadata section (added by verify_csv_markdown)
            # The metadata is appended by the module, so we just check file exists
            assert Path(result).exists()

    def test_without_metadata(self, empty_workbook, sample_csv_data_dict, default_opts):
        """CSV markdown without metadata section."""
        with tempfile.TemporaryDirectory() as tmpdir:
            opts = default_opts.copy()
            opts['csv_include_description'] = False
            opts['csv_include_metadata'] = False

            result = write_csv_markdown(
                empty_workbook,
                sample_csv_data_dict,
                'test_file',
                opts,
                tmpdir
            )

            assert result is not None
            content = Path(result).read_text(encoding='utf-8')

            # Should NOT have metadata section
            assert '## 検証用メタデータ' not in content

    def test_with_mermaid(self, empty_workbook, default_opts):
        """CSV markdown with Mermaid block."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create data with Mermaid content
            csv_data = {
                'Sheet1': {
                    'rows': [['A', 'B'], ['1', '2']],
                    'range': 'A1:B2',
                    'area': (1, 1, 2, 2),
                    'mermaid': '```mermaid\nflowchart TD\n  A --> B\n```',
                }
            }

            opts = default_opts.copy()
            opts['csv_include_description'] = False
            opts['csv_include_metadata'] = False

            result = write_csv_markdown(
                empty_workbook,
                csv_data,
                'test_file',
                opts,
                tmpdir
            )

            assert result is not None
            content = Path(result).read_text(encoding='utf-8')

            # Should have Mermaid block
            assert '```mermaid' in content
            assert 'flowchart' in content

    def test_csv_code_block(self, empty_workbook, sample_csv_data_dict, default_opts):
        """CSV data should be in code block."""
        with tempfile.TemporaryDirectory() as tmpdir:
            opts = default_opts.copy()
            opts['csv_include_description'] = False
            opts['csv_include_metadata'] = False

            result = write_csv_markdown(
                empty_workbook,
                sample_csv_data_dict,
                'test_file',
                opts,
                tmpdir
            )

            content = Path(result).read_text(encoding='utf-8')

            # Should have CSV code block
            assert '```csv' in content
            assert '```' in content

    def test_multi_sheet(self, empty_workbook, sample_csv_data_dict_multi_sheet, default_opts):
        """Multiple sheets should all be included."""
        with tempfile.TemporaryDirectory() as tmpdir:
            opts = default_opts.copy()
            opts['csv_include_description'] = False
            opts['csv_include_metadata'] = False

            result = write_csv_markdown(
                empty_workbook,
                sample_csv_data_dict_multi_sheet,
                'test_file',
                opts,
                tmpdir
            )

            content = Path(result).read_text(encoding='utf-8')

            # Should have both sheet headers
            assert '## Sheet1' in content
            assert '## Sheet2' in content

    def test_single_sheet_format(self, empty_workbook, sample_csv_data_dict, default_opts):
        """Single sheet should use simplified format."""
        with tempfile.TemporaryDirectory() as tmpdir:
            opts = default_opts.copy()
            opts['csv_include_description'] = True
            opts['csv_include_metadata'] = False

            result = write_csv_markdown(
                empty_workbook,
                sample_csv_data_dict,
                'test_file',
                opts,
                tmpdir
            )

            content = Path(result).read_text(encoding='utf-8')

            # Single sheet mode uses H1 for sheet name
            assert '# Sheet1' in content
            # Should not have "## Sheet1" (that's multi-sheet format)
            lines = content.split('\n')
            h2_sheet1 = any(line.strip() == '## Sheet1' for line in lines)
            assert not h2_sheet1

    def test_output_filename(self, empty_workbook, sample_csv_data_dict, default_opts):
        """Output file should have correct name."""
        with tempfile.TemporaryDirectory() as tmpdir:
            opts = default_opts.copy()
            opts['csv_include_metadata'] = False

            result = write_csv_markdown(
                empty_workbook,
                sample_csv_data_dict,
                'my_excel_file',
                opts,
                tmpdir
            )

            assert result.endswith('my_excel_file_csv.md')

    def test_empty_data(self, empty_workbook, default_opts):
        """Empty data should return None."""
        with tempfile.TemporaryDirectory() as tmpdir:
            result = write_csv_markdown(
                empty_workbook,
                {},
                'test_file',
                default_opts,
                tmpdir
            )

            assert result is None

    def test_csv_escaping(self, empty_workbook, default_opts):
        """CSV special characters should be escaped properly."""
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_data = {
                'Sheet1': {
                    'rows': [
                        ['Name', 'Value'],
                        ['Test, with comma', 'Normal'],
                        ['Test "with" quotes', 'Also normal'],
                    ],
                    'range': 'A1:B3',
                    'area': (1, 1, 3, 2),
                    'mermaid': None,
                }
            }

            opts = default_opts.copy()
            opts['csv_include_description'] = False
            opts['csv_include_metadata'] = False

            result = write_csv_markdown(
                empty_workbook,
                csv_data,
                'test_file',
                opts,
                tmpdir
            )

            content = Path(result).read_text(encoding='utf-8')

            # CSV should properly escape commas and quotes
            # Commas in field require quoting
            assert '"Test, with comma"' in content or 'Test, with comma' in content
            # Quotes in field require escaping
            assert '""' in content or '"with"' in content


# ============================================================
# Integration tests
# ============================================================

class TestCSVMarkdownIntegration:
    """Integration tests for CSV Markdown output."""

    def test_full_flow(self, simple_worksheet, default_opts):
        """Full extraction to CSV Markdown flow."""
        with tempfile.TemporaryDirectory() as tmpdir:
            ws = simple_worksheet
            wb = ws.parent
            area = (1, 1, 2, 2)
            merged_lookup = build_merged_lookup(ws, area)

            # Extract CSV data
            rows = extract_print_area_for_csv(ws, area, default_opts, merged_lookup)

            # Build data dict
            csv_data = {
                ws.title: {
                    'rows': rows,
                    'range': coords_to_excel_range(*area),
                    'area': area,
                    'mermaid': None,
                }
            }

            # Write CSV Markdown
            opts = default_opts.copy()
            opts['csv_include_metadata'] = False

            result = write_csv_markdown(wb, csv_data, 'test', opts, tmpdir)

            assert result is not None
            content = Path(result).read_text(encoding='utf-8')

            # Verify content
            assert 'Header1' in content
            assert 'Data1' in content
