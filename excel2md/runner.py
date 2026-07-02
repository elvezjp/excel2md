# -*- coding: utf-8 -*-
"""メイン処理オーケストレーション

処理フロー全体の制御、各モジュールの呼び出し順序管理を担当する。
"""

from pathlib import Path
from typing import List, Tuple, Optional

from . import __version__ as VERSION
from .config import ConversionConfig
from .output import warn, info
from .workbook_loader import load_workbook_safe, get_print_areas
from .mermaid_generator import _v14_extract_shapes_to_mermaid
from .table_detection import build_merged_lookup, grid_to_tables, union_rects
from .table_extraction import extract_table, dispatch_table_output
from .table_formatting import make_markdown_table
from .image_extraction import extract_images_from_sheet
from .csv_export import coords_to_excel_range, write_csv_markdown, extract_print_area_for_csv

def run(input_path: str, output_path: Optional[str], args):
    """Excel→Markdown変換のメイン処理を実行する。

    ``args`` には argparse.Namespace (CLI 経由) または ConversionConfig
    (ライブラリ呼び出し / ExcelConverter 経由) のいずれも渡せる。両者は
    同じ属性インタフェースを持つが、内部処理は ConversionConfig に
    正規化して扱う (Issue #16 Phase 4)。
    """
    config = args if isinstance(args, ConversionConfig) else ConversionConfig.from_args(args)

    # ワークブック読み込み
    wb = load_workbook_safe(input_path, read_only=config.read_only)
    sheets = wb.sheetnames

    split_by_sheet = config.split_by_sheet

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
    # ConversionConfig.to_opts_dict() がかつての inline 辞書と完全に同等の
    # 形を返すため、ここでは委譲するだけで振る舞いは変わらない。
    # test_config.py の test_to_opts_dict_matches_runner_inline_dict_* が
    # この同等性を回帰テストとして担保している。
    opts = config.to_opts_dict()

    # CSV Markdown出力の準備
    csv_output_dir = opts.get("csv_output_dir") or str(Path(input_path).parent)
    csv_basename = Path(input_path).stem
    csv_markdown_data = {}

    # CSV Markdown出力が有効な場合、通常Markdownは最終的に出力されず破棄されるため、
    # 通常Markdown用の組み立て（テーブル検出・抽出・形式判定・整形・脚注管理）を
    # シートループごとスキップする (Issue #26)
    emit_normal_md = not opts.get("csv_markdown_enabled", True)

    # シート単位ループ
    for sname in sheets:
        sheet_counter += 1
        ws = wb[sname]

        # 保護状態チェック
        if getattr(getattr(ws, "protection", None), "sheet", False):
            info(f"Sheet '{sname}' is protected (read-only); proceeding with read-only extraction.")

        # シートごとのMarkdown行を初期化（通常Markdown出力時のみ）
        if emit_normal_md:
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
            if emit_normal_md:
                current_md_lines.append(f"## {sname}\n（シート数上限によりスキップ）\n\n---\n")
            continue

        if emit_normal_md and not split_by_sheet:
            current_md_lines.append(f"## {sname}\n")

        # shapes検出モード時のMermaid生成
        # 通常Markdown用とCSV Markdown用で同一の結果になるため、シートごとに
        # 1回だけ抽出して両方で使い回す (Issue #26)
        shapes_mermaid = None
        shapes_mermaid_ready = False
        if emit_normal_md and opts.get("mermaid_enabled", False) and opts.get("mermaid_detect_mode") == "shapes":
            shapes_mermaid = _v14_extract_shapes_to_mermaid(input_path, ws, opts)
            shapes_mermaid_ready = True
            if shapes_mermaid:
                current_md_lines.append(shapes_mermaid + "\n")
                current_md_lines.append("\n---\n")

        # 印刷領域取得
        areas = get_print_areas(ws, opts["no_print_area_mode"])
        if not areas:
            if emit_normal_md:
                current_md_lines.append("（テーブルなし）\n\n---\n")
            continue

        # 矩形和集合計算
        unioned = union_rects(areas)

        # 通常Markdown用のテーブル検出・抽出・整形
        if emit_normal_md:
            # 脚注スコープ処理
            if opts["footnote_scope"] == "sheet":
                footnotes = []
                global_footnote_start = 1

            table_id = 0

            # 矩形・テーブル単位ループ
            for union_area in unioned:
                # 結合セルマップ作成
                merged_lookup = build_merged_lookup(ws, union_area)

                # テーブル分割検出（結合セルマップは構築済みのものを再利用）
                tables = grid_to_tables(ws, union_area, hidden_policy=opts["hidden_policy"], opts=opts, merged_lookup=merged_lookup)
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
                    global_footnote_start += len(note_refs)

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
            # 仕様 §「印刷領域内のみが変換対象となる」を満たすため、画像位置で
            # union_area を拡張しない。印刷領域外の画像は extract_print_area_for_csv
            # 側の範囲反復で自然にフィルタされる。
            for union_area in unioned:
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

            # CSV Markdown用Mermaid抽出（通常Markdown用に抽出済みならその結果を再利用）
            if opts.get("mermaid_enabled", False) and sname in csv_markdown_data:
                detect_mode = opts.get("mermaid_detect_mode", "shapes")
                if detect_mode == "shapes":
                    if not shapes_mermaid_ready:
                        shapes_mermaid = _v14_extract_shapes_to_mermaid(input_path, ws, opts)
                        shapes_mermaid_ready = True
                    if shapes_mermaid:
                        csv_markdown_data[sname]["mermaid"] = shapes_mermaid
                elif detect_mode in ("column_headers", "heuristic"):
                    # CSV Markdownではcolumn_headers/heuristicモード非対応（テーブル分割なしのため）
                    warn(f"mermaid_detect_mode='{detect_mode}' is not supported for CSV markdown output (only 'shapes' is supported). Mermaid output will be skipped for sheet '{sname}'.")

        # シート単位スコープの脚注を保存・出力（通常Markdown出力時のみ）
        if emit_normal_md and (split_by_sheet or opts["footnote_scope"] == "sheet"):
            if split_by_sheet:
                sheet_footnotes_dict[sname] = list(footnotes)
            if footnotes and opts["hyperlink_mode"] in ("footnote", "both"):
                footnotes_sorted = sorted(set(footnotes), key=lambda x: x[0])
                current_md_lines.append("\n")
                for idx, txt in footnotes_sorted:
                    current_md_lines.append(f"[^{idx}]: {txt}")

    # 通常モード: ドキュメント末尾に脚注を追加
    if emit_normal_md and not split_by_sheet:
        if opts["footnote_scope"] != "sheet" and footnotes and opts["hyperlink_mode"] in ("footnote", "both"):
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
