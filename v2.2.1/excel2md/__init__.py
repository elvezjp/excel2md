# -*- coding: utf-8 -*-
"""excel2md package (v2.2.1)."""

__version__ = "2.2.1"

# ライブラリとしての公開 API (issue #16)
from .config import ConversionConfig
from .converter import ExcelConverter, convert_to_markdown

# Library-level exception types (issue #36)
from .exceptions import ExcelConversionError, WorkbookOpenError

# Backward-compatible re-exports for the v1.x public API surface
# (see issue #15 — these symbols moved into submodules in v2.x and stopped
# being importable from ``excel2md`` / ``excel_to_md``).
from .table_formatting import is_code_block, build_code_block_from_rows

__all__ = [
    "__version__",
    "ConversionConfig",
    "ExcelConverter",
    "convert_to_markdown",
    "ExcelConversionError",
    "WorkbookOpenError",
    "is_code_block",
    "build_code_block_from_rows",
]
