"""Regression tests for the public API surface.

Issue #15: ``is_code_block`` and ``build_code_block_from_rows`` were importable
from the top-level ``excel_to_md`` module in v1.8 but were dropped from the
v2.x public surface. They must remain importable from both ``excel2md`` and
``excel_to_md`` for backward compatibility.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))


_JAVA_ROWS = [
    ["public class Example {"],
    ["    private int value;"],
    ["}"],
]


def test_is_code_block_importable_from_excel2md_package():
    from excel2md import is_code_block

    assert callable(is_code_block)
    assert is_code_block(_JAVA_ROWS) is True
    assert is_code_block([["Name", "Value"], ["Item", "1"]]) is False


def test_build_code_block_from_rows_importable_from_excel2md_package():
    from excel2md import build_code_block_from_rows

    assert callable(build_code_block_from_rows)
    out = build_code_block_from_rows(_JAVA_ROWS)
    assert out is not None
    assert "public class Example" in out


def test_is_code_block_importable_from_excel_to_md_facade():
    from excel_to_md import is_code_block

    assert callable(is_code_block)


def test_build_code_block_from_rows_importable_from_excel_to_md_facade():
    from excel_to_md import build_code_block_from_rows

    assert callable(build_code_block_from_rows)


def test_same_callable_across_paths():
    """Re-exports must point at the canonical implementation, not copies."""
    from excel2md import is_code_block as via_pkg
    from excel_to_md import is_code_block as via_facade
    from excel2md.table_formatting import is_code_block as canonical

    assert via_pkg is canonical
    assert via_facade is canonical
