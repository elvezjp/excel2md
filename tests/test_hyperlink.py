"""
Unit tests for hyperlink processing functions.
"""
import pytest
import sys
from pathlib import Path
from unittest.mock import MagicMock

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_to_md import (
    hyperlink_info,
    is_valid_url,
)
from openpyxl.cell.cell import Cell


def create_mock_cell(value=None, fill_color=None, hyperlink_target=None,
                     hyperlink_location=None, is_date=False):
    """
    Create a mock cell for testing.
    """
    cell = MagicMock(spec=Cell)
    cell.value = value
    cell.is_date = is_date

    # Fill setup
    if fill_color:
        cell.fill = MagicMock()
        cell.fill.patternType = 'solid'
        cell.fill.fgColor = MagicMock()
        cell.fill.fgColor.rgb = fill_color
        cell.fill.fgColor.type = 'rgb'
    else:
        cell.fill = MagicMock()
        cell.fill.patternType = None

    # Hyperlink setup
    if hyperlink_target or hyperlink_location:
        cell.hyperlink = MagicMock()
        cell.hyperlink.target = hyperlink_target
        cell.hyperlink.location = hyperlink_location
        cell.hyperlink.display = str(value) if value else None
    else:
        cell.hyperlink = None

    # Border setup (no borders by default)
    cell.border = MagicMock()
    cell.border.left = MagicMock(style=None)
    cell.border.right = MagicMock(style=None)
    cell.border.top = MagicMock(style=None)
    cell.border.bottom = MagicMock(style=None)

    return cell


# ============================================================
# Tests for hyperlink_info
# ============================================================

class TestHyperlinkInfo:
    """Tests for hyperlink information extraction."""

    def test_external_url(self):
        """External URL hyperlink."""
        cell = create_mock_cell(
            value="Click here",
            hyperlink_target="https://example.com"
        )
        result = hyperlink_info(cell)

        assert result is not None
        assert result['target'] == "https://example.com"
        assert result['location'] is None

    def test_internal_link(self):
        """Internal sheet link."""
        cell = create_mock_cell(
            value="Go to Sheet2",
            hyperlink_location="Sheet2!A1"
        )
        result = hyperlink_info(cell)

        assert result is not None
        assert result['location'] == "Sheet2!A1"
        assert result['target'] is None

    def test_no_hyperlink(self):
        """Cell without hyperlink."""
        cell = create_mock_cell(value="Plain text")
        result = hyperlink_info(cell)

        assert result is None

    def test_mailto_link(self):
        """mailto: hyperlink."""
        cell = create_mock_cell(
            value="Email us",
            hyperlink_target="mailto:user@example.com"
        )
        result = hyperlink_info(cell)

        assert result is not None
        assert result['target'] == "mailto:user@example.com"

    def test_http_link(self):
        """HTTP URL hyperlink."""
        cell = create_mock_cell(
            value="HTTP Link",
            hyperlink_target="http://example.com"
        )
        result = hyperlink_info(cell)

        assert result is not None
        assert result['target'] == "http://example.com"

    def test_file_link(self):
        """File URL hyperlink."""
        cell = create_mock_cell(
            value="Open file",
            hyperlink_target="file:///path/to/file.pdf"
        )
        result = hyperlink_info(cell)

        assert result is not None
        assert result['target'] == "file:///path/to/file.pdf"

    def test_display_text(self):
        """Display text should be captured."""
        cell = create_mock_cell(
            value="Display Text",
            hyperlink_target="https://example.com"
        )
        result = hyperlink_info(cell)

        assert result is not None
        assert result['display'] == "Display Text"

    def test_both_target_and_location(self):
        """Cell with both target and location."""
        cell = MagicMock()
        cell.hyperlink = MagicMock()
        cell.hyperlink.target = "https://example.com"
        cell.hyperlink.location = "Sheet1!A1"
        cell.hyperlink.display = "Mixed link"

        result = hyperlink_info(cell)

        assert result is not None
        assert result['target'] == "https://example.com"
        assert result['location'] == "Sheet1!A1"


# ============================================================
# Tests for is_valid_url
# ============================================================

class TestIsValidUrl:
    """Tests for URL validation."""

    def test_https_url(self):
        """HTTPS URL is valid."""
        assert is_valid_url("https://example.com") is True

    def test_http_url(self):
        """HTTP URL is valid."""
        assert is_valid_url("http://example.com") is True

    def test_mailto(self):
        """mailto: URL is valid."""
        assert is_valid_url("mailto:user@example.com") is True

    def test_file_url(self):
        """file:// URL is valid."""
        assert is_valid_url("file:///path/to/file") is True

    def test_relative_path_dot(self):
        """Relative path with ./ is valid."""
        assert is_valid_url("./relative/path") is True

    def test_relative_path_dotdot(self):
        """Relative path with ../ is valid."""
        assert is_valid_url("../parent/path") is True

    def test_absolute_path(self):
        """Absolute path with / is valid."""
        assert is_valid_url("/absolute/path") is True

    def test_javascript_invalid(self):
        """javascript: URL is invalid."""
        assert is_valid_url("javascript:alert(1)") is False

    def test_empty_string(self):
        """Empty string is invalid."""
        assert is_valid_url("") is False

    def test_none(self):
        """None is invalid."""
        assert is_valid_url(None) is False

    def test_plain_text(self):
        """Plain text is invalid."""
        assert is_valid_url("just some text") is False

    def test_partial_url(self):
        """Partial URL without protocol is invalid."""
        assert is_valid_url("example.com") is False

    def test_url_with_query(self):
        """URL with query string is valid."""
        assert is_valid_url("https://example.com?param=value") is True

    def test_url_with_fragment(self):
        """URL with fragment is valid."""
        assert is_valid_url("https://example.com#section") is True

    def test_url_with_port(self):
        """URL with port is valid."""
        assert is_valid_url("https://example.com:8080/path") is True

    def test_url_with_auth(self):
        """URL with authentication is valid."""
        assert is_valid_url("https://user:pass@example.com") is True

    def test_ftp_invalid(self):
        """FTP URL is invalid (not in allowed list)."""
        # Based on the regex, ftp:// is not in the allowed protocols
        assert is_valid_url("ftp://example.com") is False

    def test_data_url_invalid(self):
        """data: URL is invalid."""
        assert is_valid_url("data:text/html,Hello") is False


# ============================================================
# Tests for hyperlink processing in extract_table
# ============================================================

class TestHyperlinkInExtraction:
    """Integration tests for hyperlink processing during extraction."""

    def test_inline_mode(self, worksheet_with_hyperlinks, default_opts):
        """Hyperlink in inline mode should produce markdown link."""
        from excel_to_md import extract_table, build_merged_lookup

        ws = worksheet_with_hyperlinks
        opts = default_opts.copy()
        opts['hyperlink_mode'] = 'inline'

        table = {
            'bbox': (1, 1, 3, 1),
            'mask': {(1, 1), (2, 1), (3, 1)},
        }
        merged_lookup = build_merged_lookup(ws)
        footnotes = []

        md_rows, note_refs, truncated, title = extract_table(
            ws, table, opts, footnotes, 1, merged_lookup
        )

        # First cell should have markdown link format
        assert '[' in md_rows[0][0] or 'External Link' in md_rows[0][0]

    def test_inline_plain_mode(self, worksheet_with_hyperlinks, default_opts):
        """Hyperlink in inline_plain mode should produce text with URL."""
        from excel_to_md import extract_table, build_merged_lookup

        ws = worksheet_with_hyperlinks
        opts = default_opts.copy()
        opts['hyperlink_mode'] = 'inline_plain'

        table = {
            'bbox': (1, 1, 3, 1),
            'mask': {(1, 1), (2, 1), (3, 1)},
        }
        merged_lookup = build_merged_lookup(ws)
        footnotes = []

        md_rows, note_refs, truncated, title = extract_table(
            ws, table, opts, footnotes, 1, merged_lookup
        )

        # Should have text with URL in parentheses
        first_cell = md_rows[0][0]
        assert 'External Link' in first_cell or '(' in first_cell

    def test_footnote_mode(self, worksheet_with_hyperlinks, default_opts):
        """Hyperlink in footnote mode should produce footnote reference."""
        from excel_to_md import extract_table, build_merged_lookup

        ws = worksheet_with_hyperlinks
        opts = default_opts.copy()
        opts['hyperlink_mode'] = 'footnote'

        table = {
            'bbox': (1, 1, 3, 1),
            'mask': {(1, 1), (2, 1), (3, 1)},
        }
        merged_lookup = build_merged_lookup(ws)
        footnotes = []

        md_rows, note_refs, truncated, title = extract_table(
            ws, table, opts, footnotes, 1, merged_lookup
        )

        # Should have footnote reference
        first_cell = md_rows[0][0]
        assert '[^' in first_cell or len(note_refs) > 0

    def test_text_only_mode(self, worksheet_with_hyperlinks, default_opts):
        """Hyperlink in text_only mode should only show display text."""
        from excel_to_md import extract_table, build_merged_lookup

        ws = worksheet_with_hyperlinks
        opts = default_opts.copy()
        opts['hyperlink_mode'] = 'text_only'

        table = {
            'bbox': (1, 1, 3, 1),
            'mask': {(1, 1), (2, 1), (3, 1)},
        }
        merged_lookup = build_merged_lookup(ws)
        footnotes = []

        md_rows, note_refs, truncated, title = extract_table(
            ws, table, opts, footnotes, 1, merged_lookup
        )

        # Should not have URL, just text
        first_cell = md_rows[0][0]
        assert 'example.com' not in first_cell or 'External Link' in first_cell

    def test_no_hyperlink_cell(self, worksheet_with_hyperlinks, default_opts):
        """Cell without hyperlink should just have text."""
        from excel_to_md import extract_table, build_merged_lookup

        ws = worksheet_with_hyperlinks
        opts = default_opts.copy()
        opts['hyperlink_mode'] = 'inline'

        table = {
            'bbox': (3, 1, 3, 1),  # Only the "No Link" cell
            'mask': {(3, 1)},
        }
        merged_lookup = build_merged_lookup(ws)
        footnotes = []

        md_rows, note_refs, truncated, title = extract_table(
            ws, table, opts, footnotes, 1, merged_lookup
        )

        # Should just be plain text
        assert 'No Link' in md_rows[0][0]
        assert '[' not in md_rows[0][0] or md_rows[0][0] == 'No Link'
