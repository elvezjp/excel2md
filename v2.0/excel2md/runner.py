# -*- coding: utf-8 -*-
"""メイン処理オーケストレーション

処理フロー全体の制御、各モジュールの呼び出し順序管理を担当する。
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
    """Excel→Markdown変換のメイン処理を実行する。"""
    # ワークブック読み込み
    wb = load_workbook_safe(input_path, read_only=args.read_only)
    sheets = wb.sheetnames

    split_by_sheet = getattr(args, "split_by_sheet", False)

    # split_by_sheetモード: シートごとにMarkdown行と脚注を管理
    if split_by_sheet:
        sheet_md_dict = {}
        sheet_footnotes_dict = {}
    else:
        md_lines = []
        md_lines.append(f"# 変換結果: {Path(input_path).name}")
        md_lines.append("")
        md_lines.append(f"- 仕様バージョン: {VERSION}")
        md_lines.append(f"- シート数: {len(sheets)}")
        md_lines.append(f"- シート一覧: {', '.join(sheets)}")
        md_lines.append("\n---\n")

    # 脚注管理
    footnotes: List[Tuple[int,str]] = []
    global_footnote_start = 1
    sheet_counter = 0

    # オプション辞書の構築
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

        # CSV Markdown出力オプション
        "csv_output_dir": getattr(args, "csv_output_dir", None),
        "csv_apply_merge_policy": getattr(args, "csv_apply_merge_policy", True),
        "csv_normalize_values": getattr(args, "csv_normalize_values", True),
        "csv_markdown_enabled": getattr(args, "csv_markdown_enabled", True),
        "csv_include_metadata": getattr(args, "csv_include_metadata", True),
        "csv_include_description": getattr(args, "csv_include_description", True),

        # 画像抽出オプション
        "image_extraction": getattr(args, "image_extraction", True),
    }

    # CSV Markdown出力の準備
    csv_output_dir = opts.get("csv_output_dir") or str(Path(input_path).parent)
    csv_basename = Path(input_path).stem
    csv_markdown_data = {}

    # シート単位ループ
    for sname in sheets:
        sheet_counter += 1
        ws = wb[sname]

        # 保護状態チェック
        if getattr(getattr(ws, "protection", None), "sheet", False):
            info(f"Sheet '{sname}' is protected (read-only); proceeding with read-only extraction.")

        # シートごとのMarkdown行を初期化
        if split_by_sheet:
            current_md_lines = []
            current_md_lines.append(f"# {sname}")
            current_md_lines.append("")
            current_md_lines.append(f"- 仕様バージョン: {VERSION}")
            current_md_lines.append(f"- 元ファイル: {Path(input_path).name}")
            current_md_lines.append("\n---\n")
            sheet_md_dict[sname] = current_md_lines
            sheet_footnotes_dict[sname] = []
            # split_by_sheetモードではシート単位で脚注を独立管理
            if opts["footnote_scope"] == "book":
                footnotes = []
                global_footnote_start = 1
        else:
            current_md_lines = md_lines

        if opts["max_sheet_count"] and sheet_counter > opts["max_sheet_count"]:
            current_md_lines.append(f"## {sname}\n（シート数上限によりスキップ）\n\n---\n")
            continue

        if not split_by_sheet:
            current_md_lines.append(f"## {sname}\n")

        # shapes検出モード時のMermaid生成
        shapes_mermaid = None
        if opts.get("mermaid_enabled", False) and opts.get("mermaid_detect_mode") == "shapes":
            shapes_mermaid = _v14_extract_shapes_to_mermaid(input_path, ws, opts)
            if shapes_mermaid:
                current_md_lines.append(shapes_mermaid + "\n")
                current_md_lines.append("\n---\n")

        # 印刷領域取得
        areas = get_print_areas(ws, opts["no_print_area_mode"])
        if not areas:
            current_md_lines.append("（テーブルなし）\n\n---\n")
            continue

        # 矩形和集合計算
        unioned = union_rects(areas)

        # 脚注スコープ処理
        if opts["footnote_scope"] == "sheet":
            footnotes = []
            global_footnote_start = 1

        table_id = 0

        # 矩形・テーブル単位ループ
        for union_area in unioned:
            # 結合セルマップ作成
            merged_lookup = build_merged_lookup(ws, union_area)

            # テーブル分割検出
            tables = grid_to_tables(ws, union_area, hidden_policy=opts["hidden_policy"], opts=opts)
            if not tables:
                continue

            # 各テーブル単位ループ
            for tbl in tables:
                table_id += 1

                # テーブル抽出
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

                # テーブル形式判定・出力
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
                    # 通常テーブル形式
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

        # CSV Markdown出力処理
        if opts.get("csv_markdown_enabled", True):
            # 画像抽出
            cell_to_image = {}
            if opts.get("image_extraction", True):
                cell_to_image = extract_images_from_sheet(ws, Path(csv_output_dir), sname, csv_basename, opts, xlsx_path=input_path)

            # CSVデータ収集
            for union_area in unioned:
                # 画像位置を含むように範囲を拡張
                if cell_to_image:
                    min_r, min_c, max_r, max_c = union_area
                    for (img_row, img_col) in cell_to_image.keys():
                        min_r = min(min_r, img_row)
                        min_c = min(min_c, img_col)
                        max_r = max(max_r, img_row)
                        max_c = max(max_c, img_col)
                    union_area = (min_r, min_c, max_r, max_c)

                merged_lookup = build_merged_lookup(ws, union_area)
                try:
                    csv_rows = extract_print_area_for_csv(ws, union_area, opts, merged_lookup, cell_to_image)
                    if csv_rows:
                        excel_range = coords_to_excel_range(*union_area)
                        csv_markdown_data[sname] = {
                            "rows": csv_rows,
                            "range": excel_range,
                            "area": union_area,
                            "mermaid": None,
                        }
                except Exception as e:
                    warn(f"CSV data extraction failed for sheet '{sname}': {e}")

            # CSV Markdown用Mermaid抽出
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
                    # CSV Markdownではcolumn_headers/heuristicモード非対応（テーブル分割なしのため）
                    warn(f"mermaid_detect_mode='{detect_mode}' is not supported for CSV markdown output (only 'shapes' is supported). Mermaid output will be skipped for sheet '{sname}'.")

        # split_by_sheetモード: シートごとの脚注を保存・出力
        if split_by_sheet:
            sheet_footnotes_dict[sname] = list(footnotes)
            if footnotes and opts["hyperlink_mode"] in ("footnote", "both"):
                footnotes_sorted = sorted(set(footnotes), key=lambda x: x[0])
                current_md_lines.append("\n")
                for idx, txt in footnotes_sorted:
                    current_md_lines.append(f"[^{idx}]: {txt}")

    # 通常モード: ドキュメント末尾に脚注を追加
    if not split_by_sheet:
        if footnotes and opts["hyperlink_mode"] in ("footnote", "both"):
            footnotes_sorted = sorted(set(footnotes), key=lambda x: x[0])
            md_lines.append("\n")
            for idx, txt in footnotes_sorted:
                md_lines.append(f"[^{idx}]: {txt}")

    # 出力ファイル書き込み
    if opts.get("csv_markdown_enabled", True):
        # CSV Markdown出力モード
        if csv_markdown_data:
            try:
                if split_by_sheet:
                    # シートごと分割出力
                    output_dir = Path(output_path).parent if output_path else Path(input_path).parent
                    output_basename = Path(output_path).stem if output_path else Path(input_path).stem
                    output_files = []

                    for sname in sheets:
                        if sname not in csv_markdown_data:
                            continue
                        # シート名をファイル名用にサニタイズ
                        safe_sheet_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in sname)
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
                    # 通常モード: 単一ファイルに全シート出力
                    csv_file = write_csv_markdown(wb, csv_markdown_data, csv_basename, opts, csv_output_dir)
                    return csv_file or "CSV markdown output completed"
            except Exception as e:
                warn(f"CSV markdown output failed: {e}")
                return None
        else:
            warn("No CSV data to output")
            return None

    # 通常Markdown出力モード
    if split_by_sheet:
        # シートごと分割出力
        output_dir = Path(output_path).parent if output_path else Path(input_path).parent
        output_basename = Path(output_path).stem if output_path else Path(input_path).stem
        output_files = []

        for sname in sheets:
            if sname not in sheet_md_dict:
                continue
            # シート名をファイル名用にサニタイズ
            safe_sheet_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in sname)
            sheet_output_path = output_dir / f"{output_basename}_{safe_sheet_name}.md"
            Path(sheet_output_path).write_text("\n".join(sheet_md_dict[sname]), encoding="utf-8")
            output_files.append(str(sheet_output_path))

        return "\n".join([f"シートごとに分割して出力しました:"] + output_files)
    else:
        # 通常モード: 単一ファイルに出力
        if not output_path:
            output_path = str(Path(input_path).with_suffix(".md"))
        Path(output_path).write_text("\n".join(md_lines), encoding="utf-8")
        return output_path
