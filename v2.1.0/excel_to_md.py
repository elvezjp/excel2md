# -*- coding: utf-8 -*-
"""Excel -> Markdown Converter.

仕様書参照: §3.1.2 主要設定オプション、§4 処理フロー
"""

from excel2md.cli import build_argparser, main
from excel2md.runner import run
from excel2md import __version__ as VERSION
from excel2md.output import warn, info
from excel2md.cell_utils import (
    remove_control_chars,
    is_whitespace_only,
    md_escape,
    no_fill,
    excel_is_date,
    format_value,
    cell_display_value,
    cell_is_empty,
    numeric_like,
    normalize_numeric_text,
    hyperlink_info,
    is_valid_url,
    has_border,
)
from excel2md.workbook_loader import load_workbook_safe, a1_from_rc, parse_dimension, get_print_areas
from excel2md.mermaid_generator import (
    _v14_normalize_header_name,
    _v14_sanitize_node_id,
    _v14_resolve_columns_by_name,
    _v14_extract_shapes_to_mermaid,
    _v14_infer_edges,
    is_flow_table,
    build_mermaid,
)
from excel2md.image_extraction import extract_images_from_xlsx_drawing, extract_images_from_sheet, sanitize_sheet_name
from excel2md.table_detection import (
    collect_hidden,
    build_nonempty_grid,
    enumerate_histogram_rectangles,
    carve_rectangles,
    bfs_components,
    rectangles_for_component,
    union_rects,
    grid_to_tables,
    build_merged_lookup,
)
from excel2md.table_extraction import detect_table_title, extract_table, dispatch_table_output
from excel2md.table_formatting import (
    is_source_code,
    detect_code_language,
    format_table_as_text_or_nested,
    choose_header_row_heuristic,
    detect_right_align,
    make_markdown_table,
)
from excel2md.csv_export import coords_to_excel_range, format_timestamp, write_csv_markdown, extract_print_area_for_csv

__all__ = [
    "VERSION",
    "run",
    "build_argparser",
    "main",
    "warn",
    "info",
    "remove_control_chars",
    "is_whitespace_only",
    "md_escape",
    "no_fill",
    "excel_is_date",
    "format_value",
    "cell_display_value",
    "cell_is_empty",
    "numeric_like",
    "normalize_numeric_text",
    "hyperlink_info",
    "is_valid_url",
    "load_workbook_safe",
    "a1_from_rc",
    "parse_dimension",
    "get_print_areas",
    "_v14_normalize_header_name",
    "_v14_sanitize_node_id",
    "_v14_resolve_columns_by_name",
    "_v14_extract_shapes_to_mermaid",
    "_v14_infer_edges",
    "is_flow_table",
    "build_mermaid",
    "extract_images_from_xlsx_drawing",
    "extract_images_from_sheet",
    "sanitize_sheet_name",
    "collect_hidden",
    "build_nonempty_grid",
    "enumerate_histogram_rectangles",
    "carve_rectangles",
    "bfs_components",
    "rectangles_for_component",
    "union_rects",
    "grid_to_tables",
    "build_merged_lookup",
    "detect_table_title",
    "extract_table",
    "has_border",
    "dispatch_table_output",
    "is_source_code",
    "detect_code_language",
    "format_table_as_text_or_nested",
    "choose_header_row_heuristic",
    "detect_right_align",
    "make_markdown_table",
    "coords_to_excel_range",
    "format_timestamp",
    "write_csv_markdown",
    "extract_print_area_for_csv",
]

if __name__ == "__main__":
    main()
