# -*- coding: utf-8 -*-
"""Run orchestration for Excel to Markdown conversion.

仕様書参照: §4 処理フロー
"""

from pathlib import Path
from typing import List, Tuple, Optional

from . import __version__ as VERSION
from .output import warn, info
from .workbook_loader import load_workbook_safe, get_print_areas
from .mermaid_generator import _v14_extract_shapes_to_mermaid
from .table_detection import build_merged_lookup, grid_to_tables, union_rects
from .table_extraction import extract_table, dispatch_table_output
from .table_formatting import make_markdown_table
from .image_extraction import extract_images_from_sheet
from .csv_export import coords_to_excel_range, write_csv_markdown, extract_print_area_for_csv

def run(input_path: str, output_path: Optional[str], args):
    wb = load_workbook_safe(input_path, read_only=args.read_only)
    sheets = wb.sheetnames

    split_by_sheet = getattr(args, "split_by_sheet", False)

    # For split_by_sheet mode, use a dict to store md_lines for each sheet
    if split_by_sheet:
        sheet_md_dict = {}  # {sheet_name: [md_lines]}
        sheet_footnotes_dict = {}  # {sheet_name: [(idx, text)]}
    else:
        md_lines = []
        md_lines.append(f"# 変換結果: {Path(input_path).name}")
        md_lines.append("")
        md_lines.append(f"- 仕様バージョン: {VERSION}")
        md_lines.append(f"- シート数: {len(sheets)}")
        md_lines.append(f"- シート一覧: {', '.join(sheets)}")
        md_lines.append("\n---\n")

    footnotes: List[Tuple[int,str]] = []
    global_footnote_start = 1
    sheet_counter = 0

    opts = {
        "no_print_area_mode": args.no_print_area_mode,
        "value_mode": args.value_mode,
        "merge_policy": args.merge_policy,
        "hyperlink_mode": args.hyperlink_mode,
        "header_detection": (args.header_detection == "first_row"),
        "hidden_policy": args.hidden_policy,
        "strip_whitespace": args.strip_whitespace,
        "escape_pipes": args.escape_pipes,
        "date_format_override": args.date_format_override,
        "date_default_format": args.date_default_format,
        "numeric_thousand_sep": args.numeric_thousand_sep,
        "percent_format": args.percent_format,
        "currency_symbol": args.currency_symbol,
        "percent_divide_100": args.percent_divide_100,
        "readonly_fill_policy": getattr(args, "readonly_fill_policy", "assume_no_fill"),
        "align_detection": (args.align_detection == "numbers_right"),
        "numbers_right_threshold": args.numbers_right_threshold,
        "max_sheet_count": args.max_sheet_count,
        "max_cells_per_table": args.max_cells_per_table,
        "sort_tables": args.sort_tables,
        "footnote_scope": args.footnote_scope,
        "locale": args.locale,
        "markdown_escape_level": args.markdown_escape_level,
        "mermaid_enabled": args.mermaid_enabled,
        "mermaid_detect_mode": args.mermaid_detect_mode,
        "mermaid_diagram_type": getattr(args, "mermaid_diagram_type", "flowchart"),
        "mermaid_direction": args.mermaid_direction,
        "mermaid_keep_source_table": getattr(args, "mermaid_keep_source_table", True),
        "mermaid_dedupe_edges": getattr(args, "mermaid_dedupe_edges", True),
        "mermaid_node_id_policy": getattr(args, "mermaid_node_id_policy", "auto"),
        "mermaid_group_column_behavior": getattr(args, "mermaid_group_column_behavior", "subgraph"),
        "mermaid_columns": (lambda s: {
            "from": (s.split(",")[0].strip() if len(s.split(","))>0 else "From"),
            "to": (s.split(",")[1].strip() if len(s.split(","))>1 else "To"),
            "label": (s.split(",")[2].strip() if len(s.split(","))>2 else "Label"),
            "group": (s.split(",")[3].strip() if len(s.split(","))>3 else None),
            "note": (s.split(",")[4].strip() if len(s.split(","))>4 else None),
        })(args.mermaid_columns),
        "mermaid_heuristic_min_rows": args.mermaid_heuristic_min_rows,
        "mermaid_heuristic_arrow_ratio": args.mermaid_heuristic_arrow_ratio,
        "mermaid_heuristic_len_median_ratio_min": args.mermaid_heuristic_len_median_ratio_min,
        "mermaid_heuristic_len_median_ratio_max": args.mermaid_heuristic_len_median_ratio_max,
        "dispatch_skip_code_and_mermaid_on_fallback": getattr(args, "dispatch_skip_code_and_mermaid_on_fallback", True),

        "detect_dates": True,
        "prefer_excel_display": args.prefer_excel_display,

        # CSV markdown output options
        "csv_output_dir": getattr(args, "csv_output_dir", None),
        "csv_apply_merge_policy": getattr(args, "csv_apply_merge_policy", True),
        "csv_normalize_values": getattr(args, "csv_normalize_values", True),
        "csv_markdown_enabled": getattr(args, "csv_markdown_enabled", True),
        "csv_include_metadata": getattr(args, "csv_include_metadata", True),
        "csv_include_description": getattr(args, "csv_include_description", True),

        # Image extraction options
        "image_extraction": getattr(args, "image_extraction", True),
    }

    # Prepare for CSV markdown output
    csv_output_dir = opts.get("csv_output_dir") or str(Path(input_path).parent)
    csv_basename = Path(input_path).stem
    csv_markdown_data = {}  # For CSV markdown output (dict of sheet_name: rows)

    for sname in sheets:
        sheet_counter += 1
        ws = wb[sname]
        if getattr(getattr(ws, "protection", None), "sheet", False):
            info(f"Sheet '{sname}' is protected (read-only); proceeding with read-only extraction.")

        # Initialize current_md_lines for this sheet
        if split_by_sheet:
            current_md_lines = []
            current_md_lines.append(f"# {sname}")
            current_md_lines.append("")
            current_md_lines.append(f"- 仕様バージョン: {VERSION}")
            current_md_lines.append(f"- 元ファイル: {Path(input_path).name}")
            current_md_lines.append("\n---\n")
            sheet_md_dict[sname] = current_md_lines
            sheet_footnotes_dict[sname] = []
            if opts["footnote_scope"] == "book":
                # For split_by_sheet mode, treat each sheet as independent (sheet scope)
                footnotes = []
                global_footnote_start = 1
        else:
            current_md_lines = md_lines

        if opts["max_sheet_count"] and sheet_counter > opts["max_sheet_count"]:
            current_md_lines.append(f"## {sname}\n（シート数上限によりスキップ）\n\n---\n")
            continue

        if not split_by_sheet:
            current_md_lines.append(f"## {sname}\n")

        # シェイプ（図形）からのMermaid検出（シート単位で実行）
        shapes_mermaid = None
        if opts.get("mermaid_enabled", False) and opts.get("mermaid_detect_mode") == "shapes":
            shapes_mermaid = _v14_extract_shapes_to_mermaid(input_path, ws, opts)
            if shapes_mermaid:
                current_md_lines.append(shapes_mermaid + "\n")
                current_md_lines.append("\n---\n")
        areas = get_print_areas(ws, opts["no_print_area_mode"])
        if not areas:
            current_md_lines.append("（テーブルなし）\n\n---\n")
            continue

        unioned = union_rects(areas)

        if opts["footnote_scope"] == "sheet":
            footnotes = []
            global_footnote_start = 1

        table_id = 0

        for union_area in unioned:
            merged_lookup = build_merged_lookup(ws, union_area)
            tables = grid_to_tables(ws, union_area, hidden_policy=opts["hidden_policy"], opts=opts)
            if not tables:
                continue

            # Process tables
            for tbl in tables:
                table_id += 1

                md_rows, note_refs, truncated, table_title = extract_table(ws, tbl, opts, footnotes, global_footnote_start, merged_lookup, print_area=union_area)

                if table_title:
                    current_md_lines.append(f"### {table_title}")
                else:
                    current_md_lines.append(f"### Table {table_id}")
                for (n, txt) in note_refs:
                    footnotes.append((n, txt))

                if not md_rows:
                    current_md_lines.append("（テーブルなし）\n")
                    continue

                format_type, formatted_output = dispatch_table_output(ws, tbl, md_rows, opts, merged_lookup, xlsx_path=input_path)

                if format_type == "text":
                    current_md_lines.append(formatted_output + "\n")
                elif format_type == "nested":
                    current_md_lines.append(formatted_output + "\n")
                elif format_type == "code":
                    current_md_lines.append(formatted_output + "\n")
                elif format_type == "mermaid":
                    current_md_lines.append(formatted_output + "\n")
                elif format_type == "empty":
                    current_md_lines.append("\n")
                else:
                    # Note: dispatch_table_output() already handles table formatting,
                    # but this branch handles cases where format_type is "table" from dispatch
                    hdr = opts.get("header_detection", True)
                    table_md = make_markdown_table(
                        md_rows,
                        header_detection=hdr,
                        align_detect=opts["align_detection"],
                        align_threshold=opts["numbers_right_threshold"],
                    )
                    current_md_lines.append(table_md + "\n")
                    if truncated:
                        current_md_lines.append("_※ このテーブルは max_cells_per_table 制限により途中で打ち切られました。_\n")

            current_md_lines.append("\n---\n")

        if opts.get("csv_markdown_enabled", True):
            cell_to_image = {}
            if opts.get("image_extraction", True):
                cell_to_image = extract_images_from_sheet(ws, Path(csv_output_dir), sname, csv_basename, opts, xlsx_path=input_path)

            # Collect CSV data for markdown output
            for union_area in unioned:
                merged_lookup = build_merged_lookup(ws, union_area)
                try:
                    csv_rows = extract_print_area_for_csv(ws, union_area, opts, merged_lookup, cell_to_image)
                    if csv_rows:
                        # Store both CSV rows and Excel range info
                        excel_range = coords_to_excel_range(*union_area)
                        csv_markdown_data[sname] = {
                            "rows": csv_rows,
                            "range": excel_range,
                            "area": union_area,
                            "mermaid": None  # Will be populated if mermaid_enabled=true and mermaid_detect_mode="shapes"
                        }
                except Exception as e:
                    warn(f"CSV data extraction failed for sheet '{sname}': {e}")

            # Extract Mermaid for CSV markdown if mermaid_enabled=true
            if opts.get("mermaid_enabled", False) and sname in csv_markdown_data:
                detect_mode = opts.get("mermaid_detect_mode", "shapes")
                if detect_mode == "shapes":
                    try:
                        csv_mermaid = _v14_extract_shapes_to_mermaid(input_path, ws, opts)
                        if csv_mermaid:
                            csv_markdown_data[sname]["mermaid"] = csv_mermaid
                    except Exception as e:
                        warn(f"Mermaid extraction for CSV markdown failed for sheet '{sname}': {e}")
                elif detect_mode in ("column_headers", "heuristic"):
                    # column_headers and heuristic modes are not supported for CSV markdown
                    # because CSV markdown does not split tables
                    warn(f"mermaid_detect_mode='{detect_mode}' is not supported for CSV markdown output (only 'shapes' is supported). Mermaid output will be skipped for sheet '{sname}'.")

        # For split_by_sheet mode, save footnotes for this sheet and add to file
        if split_by_sheet:
            sheet_footnotes_dict[sname] = list(footnotes)
            if footnotes and opts["hyperlink_mode"] in ("footnote", "both"):
                footnotes_sorted = sorted(set(footnotes), key=lambda x: x[0])
                current_md_lines.append("\n")
                for idx, txt in footnotes_sorted:
                    current_md_lines.append(f"[^{idx}]: {txt}")

    # For non-split mode, add footnotes at the end
    if not split_by_sheet:
        if footnotes and opts["hyperlink_mode"] in ("footnote", "both"):
            footnotes_sorted = sorted(set(footnotes), key=lambda x: x[0])
            md_lines.append("\n")
            for idx, txt in footnotes_sorted:
                md_lines.append(f"[^{idx}]: {txt}")

    # Write output file(s)
    # csv_markdown_enabled=true: CSV markdown only
    # csv_markdown_enabled=false: Regular markdown only
    if opts.get("csv_markdown_enabled", True):
        # CSV markdown output mode
        if csv_markdown_data:
            try:
                if split_by_sheet:
                    # Split by sheet: write each sheet to separate CSV markdown file
                    output_dir = Path(output_path).parent if output_path else Path(input_path).parent
                    output_basename = Path(output_path).stem if output_path else Path(input_path).stem
                    output_files = []

                    for sname in sheets:
                        if sname not in csv_markdown_data:
                            continue
                        # Sanitize sheet name for filename
                        safe_sheet_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in sname)
                        # Write single-sheet CSV markdown
                        single_sheet_data = {sname: csv_markdown_data[sname]}
                        csv_file = write_csv_markdown(
                            wb, single_sheet_data,
                            f"{output_basename}_{safe_sheet_name}",
                            opts, csv_output_dir or str(output_dir)
                        )
                        if csv_file:
                            output_files.append(csv_file)

                    return "\n".join([f"シートごとに分割してCSVマークダウン出力しました:"] + output_files)
                else:
                    # Normal mode: write all sheets to single CSV markdown file
                    csv_file = write_csv_markdown(wb, csv_markdown_data, csv_basename, opts, csv_output_dir)
                    return csv_file or "CSV markdown output completed"
            except Exception as e:
                warn(f"CSV markdown output failed: {e}")
                return None
        else:
            warn("No CSV data to output")
            return None

    # Regular markdown output mode (csv_markdown_enabled=false)
    if split_by_sheet:
        # Split by sheet mode: write each sheet to a separate file
        output_dir = Path(output_path).parent if output_path else Path(input_path).parent
        output_basename = Path(output_path).stem if output_path else Path(input_path).stem
        output_files = []

        for sname in sheets:
            if sname not in sheet_md_dict:
                continue
            # Sanitize sheet name for filename
            safe_sheet_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in sname)
            sheet_output_path = output_dir / f"{output_basename}_{safe_sheet_name}.md"
            Path(sheet_output_path).write_text("\n".join(sheet_md_dict[sname]), encoding="utf-8")
            output_files.append(str(sheet_output_path))

        return "\n".join([f"シートごとに分割して出力しました:"] + output_files)
    else:
        # Normal mode: write all sheets to a single file
        if not output_path:
            output_path = str(Path(input_path).with_suffix(".md"))
        Path(output_path).write_text("\n".join(md_lines), encoding="utf-8")
        return output_path
