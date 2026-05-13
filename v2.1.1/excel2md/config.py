# -*- coding: utf-8 -*-
"""ConversionConfig — Python ネイティブな変換設定を表現する dataclass。

Issue #16 (excel2md をライブラリとして再利用可能にするリファクタリング) の Phase 1。
CLI 以外の呼び出し元（Web デモ、MCP サーバー、ノートブック等）が
argparse.Namespace を組み立てずに済むようにする。

このコミット時点では `runner.run()` はまだこの dataclass を使っていない。
PR D で runner 側を ConversionConfig ベースに置換する際、`to_opts_dict()` の
出力が現在 `runner.run()` で組み立てている opts 辞書と **完全に一致** することを
test_config.py の roundtrip テストで担保している。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Optional


@dataclass
class ConversionConfig:
    """Excel→Markdown 変換の設定を保持する。

    フィールドは現在の CLI (cli.py の build_argparser) と 1:1 で対応し、
    既定値も argparse と同一。CLI の文字列（"first_row"/"numbers_right" 等）は
    そのまま文字列として保持し、内部の bool 化や辞書展開は to_opts_dict() に
    一元化する。
    """

    # ----- General -----
    read_only: bool = False
    split_by_sheet: bool = False

    # ----- Print area / value handling -----
    no_print_area_mode: str = "used_range"          # used_range / entire_sheet_range / skip_sheet
    value_mode: str = "display"                     # display / formula / both
    merge_policy: str = "top_left_only"             # expand / repeat / warn / top_left_only
    hyperlink_mode: str = "footnote"                # inline / inline_plain / footnote / both / text_only
    header_detection: str = "first_row"             # none / first_row / heuristic (string; runner expects bool)
    hidden_policy: str = "ignore"                   # ignore / include / exclude

    # ----- Whitespace / escaping -----
    strip_whitespace: bool = True
    escape_pipes: bool = True

    # ----- Date / number formatting -----
    date_format_override: Optional[str] = None
    date_default_format: str = "YYYY-MM-DD"
    numeric_thousand_sep: str = "keep"              # keep / remove
    percent_format: str = "keep"                    # keep / numeric
    currency_symbol: str = "keep"                   # keep / strip
    percent_divide_100: bool = False

    # ----- Misc -----
    readonly_fill_policy: str = "assume_no_fill"
    align_detection: str = "numbers_right"          # none / numbers_right (string; runner expects bool)
    numbers_right_threshold: float = 0.8
    max_sheet_count: int = 0                        # 0 = unlimited
    max_cells_per_table: int = 200000
    sort_tables: str = "document_order"
    footnote_scope: str = "book"                    # book / sheet
    locale: str = "ja-JP"
    markdown_escape_level: str = "safe"             # safe / minimal / aggressive
    prefer_excel_display: bool = True

    # ----- Mermaid -----
    mermaid_enabled: bool = False
    mermaid_detect_mode: str = "shapes"             # none / column_headers / heuristic / shapes
    mermaid_diagram_type: str = "flowchart"
    mermaid_direction: str = "TD"
    mermaid_keep_source_table: bool = True
    mermaid_dedupe_edges: bool = True
    mermaid_node_id_policy: str = "auto"
    mermaid_group_column_behavior: str = "subgraph"
    mermaid_columns: str = "From,To,Label,Group,Note"  # CSV string; expanded to dict in to_opts_dict()
    mermaid_heuristic_min_rows: int = 3
    mermaid_heuristic_arrow_ratio: float = 0.3
    mermaid_heuristic_len_median_ratio_min: float = 0.4
    mermaid_heuristic_len_median_ratio_max: float = 2.5
    dispatch_skip_code_and_mermaid_on_fallback: bool = True

    # ----- CSV markdown -----
    csv_output_dir: Optional[str] = None
    csv_apply_merge_policy: bool = True
    csv_normalize_values: bool = True
    csv_markdown_enabled: bool = True
    csv_include_metadata: bool = True
    csv_include_description: bool = True

    # ----- Image extraction -----
    image_extraction: bool = True

    # =============================================================
    # Factories / converters
    # =============================================================

    @classmethod
    def from_args(cls, args: Any) -> "ConversionConfig":
        """argparse.Namespace から ConversionConfig を構築する。

        Namespace に存在しないフィールドはこの dataclass の既定値で埋める
        （getattr で fallback）。これにより、Namespace 側で `dest` が省かれている
        オプションがあっても落ちない。
        """
        defaults = cls()
        kwargs = {
            name: getattr(args, name, getattr(defaults, name))
            for name in cls.__dataclass_fields__
        }
        return cls(**kwargs)

    def to_opts_dict(self) -> Dict[str, Any]:
        """runner.run() が期待する opts 辞書を返す。

        現状の runner.run() における opts 辞書構築箇所 (v2.1.1/excel2md/runner.py
        の lines 47-105 相当) と **完全に一致** する辞書を返す。PR D で runner を
        ConversionConfig ベースに置換する際、この一致が振る舞いの不変性を担保する。
        test_config.py の test_to_opts_dict_matches_runner_inline_dict_* がこの
        一致を検証する回帰テスト。
        """
        return {
            "no_print_area_mode": self.no_print_area_mode,
            "value_mode": self.value_mode,
            "merge_policy": self.merge_policy,
            "hyperlink_mode": self.hyperlink_mode,
            "header_detection": (self.header_detection == "first_row"),
            "hidden_policy": self.hidden_policy,
            "strip_whitespace": self.strip_whitespace,
            "escape_pipes": self.escape_pipes,
            "date_format_override": self.date_format_override,
            "date_default_format": self.date_default_format,
            "numeric_thousand_sep": self.numeric_thousand_sep,
            "percent_format": self.percent_format,
            "currency_symbol": self.currency_symbol,
            "percent_divide_100": self.percent_divide_100,
            "readonly_fill_policy": self.readonly_fill_policy,
            "align_detection": (self.align_detection == "numbers_right"),
            "numbers_right_threshold": self.numbers_right_threshold,
            "max_sheet_count": self.max_sheet_count,
            "max_cells_per_table": self.max_cells_per_table,
            "sort_tables": self.sort_tables,
            "footnote_scope": self.footnote_scope,
            "locale": self.locale,
            "markdown_escape_level": self.markdown_escape_level,
            "mermaid_enabled": self.mermaid_enabled,
            "mermaid_detect_mode": self.mermaid_detect_mode,
            "mermaid_diagram_type": self.mermaid_diagram_type,
            "mermaid_direction": self.mermaid_direction,
            "mermaid_keep_source_table": self.mermaid_keep_source_table,
            "mermaid_dedupe_edges": self.mermaid_dedupe_edges,
            "mermaid_node_id_policy": self.mermaid_node_id_policy,
            "mermaid_group_column_behavior": self.mermaid_group_column_behavior,
            "mermaid_columns": self._parse_mermaid_columns(),
            "mermaid_heuristic_min_rows": self.mermaid_heuristic_min_rows,
            "mermaid_heuristic_arrow_ratio": self.mermaid_heuristic_arrow_ratio,
            "mermaid_heuristic_len_median_ratio_min": self.mermaid_heuristic_len_median_ratio_min,
            "mermaid_heuristic_len_median_ratio_max": self.mermaid_heuristic_len_median_ratio_max,
            "dispatch_skip_code_and_mermaid_on_fallback": self.dispatch_skip_code_and_mermaid_on_fallback,
            "detect_dates": True,  # runner.run() でハードコードされている
            "prefer_excel_display": self.prefer_excel_display,
            "csv_output_dir": self.csv_output_dir,
            "csv_apply_merge_policy": self.csv_apply_merge_policy,
            "csv_normalize_values": self.csv_normalize_values,
            "csv_markdown_enabled": self.csv_markdown_enabled,
            "csv_include_metadata": self.csv_include_metadata,
            "csv_include_description": self.csv_include_description,
            "image_extraction": self.image_extraction,
        }

    def _parse_mermaid_columns(self) -> Dict[str, Optional[str]]:
        parts = self.mermaid_columns.split(",")
        return {
            "from": parts[0].strip() if len(parts) > 0 else "From",
            "to": parts[1].strip() if len(parts) > 1 else "To",
            "label": parts[2].strip() if len(parts) > 2 else "Label",
            "group": parts[3].strip() if len(parts) > 3 else None,
            "note": parts[4].strip() if len(parts) > 4 else None,
        }
