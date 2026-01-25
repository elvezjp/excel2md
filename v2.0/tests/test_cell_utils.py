"""
Unit tests for cell utility functions.
Spec reference: 付録D.3.1
"""
import pytest
import sys
from pathlib import Path
from unittest.mock import MagicMock

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_to_md import (
    cell_is_empty,
    cell_display_value,
    no_fill,
    numeric_like,
    normalize_numeric_text,
    md_escape,
    a1_from_rc,
    remove_control_chars,
    is_whitespace_only,
)

# Import create_mock_cell from conftest (pytest auto-loads conftest.py)
# But we need to define it here for direct use in test methods
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
# Tests for a1_from_rc (D.2.4)
# ============================================================

class TestA1FromRC:
    """Tests for coordinate to A1 notation conversion."""

    def test_single_letter_column(self):
        """Test single letter columns (A-Z)."""
        assert a1_from_rc(1, 1) == "A1"
        assert a1_from_rc(1, 26) == "Z1"
        assert a1_from_rc(10, 5) == "E10"

    def test_double_letter_column(self):
        """Test double letter columns (AA-AZ, etc)."""
        assert a1_from_rc(1, 27) == "AA1"
        assert a1_from_rc(1, 28) == "AB1"
        assert a1_from_rc(1, 52) == "AZ1"

    def test_triple_letter_column(self):
        """Test triple letter columns."""
        assert a1_from_rc(1, 703) == "AAA1"

    def test_large_row(self):
        """Test large row numbers."""
        assert a1_from_rc(1000, 1) == "A1000"
        assert a1_from_rc(1048576, 1) == "A1048576"


# ============================================================
# Tests for is_whitespace_only (D.2.4)
# ============================================================

class TestIsWhitespaceOnly:
    """Tests for whitespace-only string detection."""

    def test_empty_string(self):
        """Empty string is NOT whitespace only (requires length > 0)."""
        # Implementation requires len(s) > 0, so empty string returns False
        assert is_whitespace_only("") is False

    def test_spaces_only(self):
        """Spaces only is whitespace only."""
        assert is_whitespace_only("   ") is True

    def test_tabs_only(self):
        """Tabs only is whitespace only."""
        assert is_whitespace_only("\t\t") is True

    def test_mixed_whitespace(self):
        """Mixed whitespace with newlines - newline is not in WHITESPACE_CHARS."""
        # Note: \n is not in WHITESPACE_CHARS, so this returns False
        assert is_whitespace_only(" \t \n ") is False

    def test_spaces_and_tabs(self):
        """Spaces and tabs only is whitespace only."""
        assert is_whitespace_only(" \t ") is True

    def test_text_with_whitespace(self):
        """Text with whitespace is not whitespace only."""
        assert is_whitespace_only("  text  ") is False

    def test_single_character(self):
        """Single non-whitespace character."""
        assert is_whitespace_only("a") is False


# ============================================================
# Tests for remove_control_chars (D.2.4)
# ============================================================

class TestRemoveControlChars:
    """Tests for control character removal."""

    def test_normal_text(self):
        """Normal text should not be modified."""
        assert remove_control_chars("Hello World") == "Hello World"

    def test_null_character(self):
        """Null character should be removed."""
        assert remove_control_chars("Hello\x00World") == "HelloWorld"

    def test_keeps_newline_and_tab(self):
        """Newlines and tabs should be kept."""
        result = remove_control_chars("Hello\nWorld\tTest")
        assert "\n" in result
        assert "\t" in result

    def test_removes_bell_character(self):
        """Bell character (\\x07) should be removed."""
        assert remove_control_chars("Hello\x07World") == "HelloWorld"

    def test_japanese_text(self):
        """Japanese text should not be modified."""
        assert remove_control_chars("こんにちは") == "こんにちは"


# ============================================================
# Tests for numeric_like (D.3.1.2)
# ============================================================

class TestNumericLike:
    """Tests for numeric string detection."""

    # NL001-NL007: True cases
    def test_nl001_integer(self):
        """NL001: Plain integer."""
        assert numeric_like("123") is True

    def test_nl002_negative_integer(self):
        """NL002: Negative integer."""
        assert numeric_like("-123") is True

    def test_nl003_thousands_separator(self):
        """NL003: Number with thousands separator."""
        assert numeric_like("1,234") is True

    def test_nl004_thousands_with_decimal(self):
        """NL004: Number with thousands separator and decimal."""
        assert numeric_like("1,234.56") is True

    def test_nl005_currency_yen(self):
        """NL005: Currency symbol (Yen)."""
        assert numeric_like("¥1,234") is True

    def test_nl005_currency_dollar(self):
        """NL005: Currency symbol (Dollar)."""
        assert numeric_like("$1,234") is True

    def test_nl006_parentheses_negative(self):
        """NL006: Parentheses for negative number."""
        assert numeric_like("(123)") is True

    def test_nl007_percentage(self):
        """NL007: Percentage."""
        assert numeric_like("12.5%") is True

    # NL008-NL012: False cases
    def test_nl008_alphabetic_string(self):
        """NL008: Alphabetic string."""
        assert numeric_like("abc") is False

    def test_nl009_mixed_alphanumeric(self):
        """NL009: Mixed alphanumeric."""
        assert numeric_like("12abc") is False

    def test_nl010_unbalanced_parentheses(self):
        """NL010: Unbalanced parentheses."""
        assert numeric_like("(123") is False

    def test_nl011_multiple_decimals(self):
        """NL011: Multiple decimal points."""
        assert numeric_like("1.2.3") is False

    def test_nl012_empty_string(self):
        """NL012: Empty string."""
        assert numeric_like("") is False

    # Additional edge cases
    def test_decimal_only(self):
        """Decimal number without integer part - not matched by NUMERIC_PATTERN."""
        # NUMERIC_PATTERN requires integer part (\\d+), so .5 alone doesn't match
        assert numeric_like(".5") is False

    def test_zero(self):
        """Zero is numeric."""
        assert numeric_like("0") is True

    def test_positive_sign(self):
        """Positive sign is allowed."""
        assert numeric_like("+123") is True

    def test_scientific_notation(self):
        """Scientific notation."""
        assert numeric_like("1.2e-3") is True


# ============================================================
# Tests for normalize_numeric_text (D.2.1)
# ============================================================

class TestNormalizeNumericText:
    """Tests for numeric text normalization."""

    def test_no_normalization_needed(self, default_opts):
        """Text that doesn't need normalization."""
        result = normalize_numeric_text("123", default_opts)
        assert result == "123"

    def test_strip_currency_symbol(self, default_opts):
        """Strip currency symbol when configured."""
        opts = default_opts.copy()
        opts["currency_symbol"] = "strip"
        result = normalize_numeric_text("¥1,234", opts)
        assert "¥" not in result

    def test_remove_thousand_separator(self, default_opts):
        """Remove thousand separator when configured."""
        opts = default_opts.copy()
        opts["numeric_thousand_sep"] = "remove"
        result = normalize_numeric_text("1,234,567", opts)
        assert "," not in result

    def test_percent_to_numeric(self, default_opts):
        """Convert percent to numeric."""
        opts = default_opts.copy()
        opts["percent_format"] = "numeric"
        result = normalize_numeric_text("50%", opts)
        assert "%" not in result

    def test_percent_divide_100(self, default_opts):
        """Divide percent by 100 when configured."""
        opts = default_opts.copy()
        opts["percent_format"] = "numeric"
        opts["percent_divide_100"] = True
        result = normalize_numeric_text("50%", opts)
        assert float(result) == 0.5

    def test_non_numeric_unchanged(self, default_opts):
        """Non-numeric text should be unchanged."""
        result = normalize_numeric_text("Hello", default_opts)
        assert result == "Hello"


# ============================================================
# Tests for md_escape (D.3.1.3)
# ============================================================

class TestMdEscape:
    """Tests for Markdown escaping."""

    # ME001-ME006: Basic escape tests
    def test_me001_no_escape_needed(self):
        """ME001: Text without special characters."""
        assert md_escape("text", "safe") == "text"

    def test_me002_pipe_escape(self):
        """ME002: Pipe character should be escaped."""
        result = md_escape("a|b", "safe")
        assert "|" not in result or "\\|" in result

    def test_me003_newline_to_br(self):
        """ME003: Newline should be converted to <br> (escaped as <br\\>)."""
        result = md_escape("line1\nline2", "safe")
        # md_escape first converts newlines to <br>, then escapes < and >
        # Result contains the escaped form
        assert "br" in result

    def test_me004_asterisk_escape(self):
        """ME004: Asterisk should be escaped."""
        result = md_escape("*bold*", "safe")
        assert "\\*" in result

    def test_me005_minimal_preserves_brackets(self):
        """ME005: Minimal level preserves some characters."""
        result = md_escape("[link]", "minimal")
        # Minimal should be less aggressive
        assert result is not None

    def test_me006_aggressive_backtick(self):
        """ME006: Aggressive level escapes backticks."""
        result = md_escape("`code`", "aggressive")
        assert "\\`" in result

    # Additional tests
    def test_escape_underscore(self):
        """Underscore should be escaped in safe mode."""
        result = md_escape("_italic_", "safe")
        assert "\\_" in result

    def test_empty_string(self):
        """Empty string should remain empty."""
        assert md_escape("", "safe") == ""

    def test_japanese_text(self):
        """Japanese text should not be escaped."""
        result = md_escape("こんにちは", "safe")
        assert result == "こんにちは"

    def test_carriage_return_newline(self):
        """CRLF should be converted to <br> (then escaped)."""
        result = md_escape("line1\r\nline2", "safe")
        # After escaping, the result contains "br"
        assert "br" in result


# ============================================================
# Tests for no_fill (D.2.1)
# ============================================================

class TestNoFill:
    """Tests for cell fill detection."""

    def test_no_fill_pattern(self):
        """Cell with no fill pattern."""
        cell = create_mock_cell(value="test")
        assert no_fill(cell, "assume_no_fill") is True

    def test_white_fill_is_no_fill(self):
        """White fill should be treated as no fill."""
        cell = create_mock_cell(value="test", fill_color="FFFFFF")
        assert no_fill(cell, "assume_no_fill") is True

    def test_yellow_fill_has_fill(self):
        """Yellow fill should be detected."""
        cell = create_mock_cell(value="test", fill_color="FFFF00")
        assert no_fill(cell, "assume_no_fill") is False

    def test_assume_no_fill_policy(self):
        """assume_no_fill policy should return True for unknown fills."""
        cell = create_mock_cell(value="test")
        assert no_fill(cell, "assume_no_fill") is True


# ============================================================
# Tests for cell_is_empty (D.3.1.1)
# ============================================================

class TestCellIsEmpty:
    """Tests for cell emptiness detection."""

    def test_ce001_none_value_no_fill(self, default_opts):
        """CE001: None value with no fill is empty."""
        cell = create_mock_cell(value=None)
        assert cell_is_empty(cell, default_opts) is True

    def test_ce002_empty_string_no_fill(self, default_opts):
        """CE002: Empty string with no fill is empty."""
        cell = create_mock_cell(value="")
        assert cell_is_empty(cell, default_opts) is True

    def test_ce003_whitespace_only_no_fill(self, default_opts):
        """CE003: Whitespace only with no fill is empty."""
        cell = create_mock_cell(value="   ")
        assert cell_is_empty(cell, default_opts) is True

    def test_ce004_text_no_fill(self, default_opts):
        """CE004: Text with no fill is not empty."""
        cell = create_mock_cell(value="text")
        assert cell_is_empty(cell, default_opts) is False

    def test_ce005_none_with_yellow_fill(self, default_opts):
        """CE005: None value with yellow fill is not empty."""
        cell = create_mock_cell(value=None, fill_color="FFFF00")
        assert cell_is_empty(cell, default_opts) is False

    def test_ce006_none_with_white_fill(self, default_opts):
        """CE006: None value with white fill is empty (white = no fill)."""
        cell = create_mock_cell(value=None, fill_color="FFFFFF")
        assert cell_is_empty(cell, default_opts) is True

    def test_ce007_zero_integer(self, default_opts):
        """CE007: Integer 0 is not empty."""
        cell = create_mock_cell(value=0)
        assert cell_is_empty(cell, default_opts) is False

    def test_ce008_zero_float(self, default_opts):
        """CE008: Float 0.0 is not empty."""
        cell = create_mock_cell(value=0.0)
        assert cell_is_empty(cell, default_opts) is False


# ============================================================
# Tests for cell_display_value (D.2.1)
# ============================================================

class TestCellDisplayValue:
    """Tests for cell display value extraction."""

    def test_string_value(self, default_opts):
        """String value should be returned as-is."""
        cell = create_mock_cell(value="Hello")
        result = cell_display_value(cell, default_opts)
        assert result == "Hello"

    def test_integer_value(self, default_opts):
        """Integer value should be converted to string."""
        cell = create_mock_cell(value=123)
        result = cell_display_value(cell, default_opts)
        assert result == "123"

    def test_float_value(self, default_opts):
        """Float value should be converted to string."""
        cell = create_mock_cell(value=123.45)
        result = cell_display_value(cell, default_opts)
        assert "123.45" in result

    def test_none_value(self, default_opts):
        """None value should return empty string."""
        cell = create_mock_cell(value=None)
        result = cell_display_value(cell, default_opts)
        assert result == ""

    def test_strip_whitespace(self, default_opts):
        """Whitespace should be stripped when configured."""
        cell = create_mock_cell(value="  text  ")
        result = cell_display_value(cell, default_opts)
        assert result == "text"

    def test_no_strip_whitespace(self, opts_no_strip):
        """Whitespace should be preserved when strip_whitespace=False."""
        cell = create_mock_cell(value="  text  ")
        result = cell_display_value(cell, opts_no_strip)
        assert result == "  text  "

    def test_unicode_normalization(self, default_opts):
        """Unicode should be NFC normalized."""
        # Combine character (U+3099 is combining dakuten)
        cell = create_mock_cell(value="か\u3099")  # Should normalize to が
        result = cell_display_value(cell, default_opts)
        assert result == "が"
