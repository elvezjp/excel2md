"""Tests for ConversionConfig (issue #16, Phase 1).

The critical test is ``test_to_opts_dict_matches_runner_inline_dict_*`` — it
verifies that the dict produced by ConversionConfig is exactly the same as
the dict that ``runner.run()`` currently constructs inline. This roundtrip
guarantee is what makes the eventual runner refactor a no-op for behavior.
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel2md.config import ConversionConfig
from excel2md.cli import build_argparser


# =============================================================
# Snapshot of the inline opts-dict construction currently inside
# runner.run() (v2.1.1/excel2md/runner.py lines 47-105 as of 2026-05-13).
# Kept verbatim here as the ground-truth baseline.
# =============================================================

def _build_opts_dict_inline(args):
    return {
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
            "from": (s.split(",")[0].strip() if len(s.split(",")) > 0 else "From"),
            "to": (s.split(",")[1].strip() if len(s.split(",")) > 1 else "To"),
            "label": (s.split(",")[2].strip() if len(s.split(",")) > 2 else "Label"),
            "group": (s.split(",")[3].strip() if len(s.split(",")) > 3 else None),
            "note": (s.split(",")[4].strip() if len(s.split(",")) > 4 else None),
        })(args.mermaid_columns),
        "mermaid_heuristic_min_rows": args.mermaid_heuristic_min_rows,
        "mermaid_heuristic_arrow_ratio": args.mermaid_heuristic_arrow_ratio,
        "mermaid_heuristic_len_median_ratio_min": args.mermaid_heuristic_len_median_ratio_min,
        "mermaid_heuristic_len_median_ratio_max": args.mermaid_heuristic_len_median_ratio_max,
        "dispatch_skip_code_and_mermaid_on_fallback": getattr(args, "dispatch_skip_code_and_mermaid_on_fallback", True),
        "detect_dates": True,
        "prefer_excel_display": args.prefer_excel_display,
        "csv_output_dir": getattr(args, "csv_output_dir", None),
        "csv_apply_merge_policy": getattr(args, "csv_apply_merge_policy", True),
        "csv_normalize_values": getattr(args, "csv_normalize_values", True),
        "csv_markdown_enabled": getattr(args, "csv_markdown_enabled", True),
        "csv_include_metadata": getattr(args, "csv_include_metadata", True),
        "csv_include_description": getattr(args, "csv_include_description", True),
        "image_extraction": getattr(args, "image_extraction", True),
    }


# =============================================================
# Construction
# =============================================================

class TestConstruction:
    def test_default_values_match_cli(self):
        """ConversionConfig() の既定値が CLI 既定値と一致すること。"""
        cfg = ConversionConfig()
        args = build_argparser().parse_args(["dummy.xlsx"])
        assert cfg.hyperlink_mode == args.hyperlink_mode
        assert cfg.max_cells_per_table == args.max_cells_per_table
        assert cfg.csv_markdown_enabled == args.csv_markdown_enabled
        assert cfg.merge_policy == args.merge_policy
        assert cfg.footnote_scope == args.footnote_scope

    def test_keyword_construction(self):
        cfg = ConversionConfig(hyperlink_mode="inline", max_cells_per_table=10)
        assert cfg.hyperlink_mode == "inline"
        assert cfg.max_cells_per_table == 10
        # 他の値は既定のまま
        assert cfg.footnote_scope == "book"


# =============================================================
# from_args
# =============================================================

class TestFromArgs:
    def test_passes_all_fields_through(self):
        args = build_argparser().parse_args(["dummy.xlsx"])
        cfg = ConversionConfig.from_args(args)
        # スポットチェック
        for fname in ("hyperlink_mode", "max_cells_per_table", "csv_markdown_enabled",
                      "merge_policy", "mermaid_columns", "footnote_scope"):
            assert getattr(cfg, fname) == getattr(args, fname), fname

    def test_namespace_without_optional_fields_uses_defaults(self):
        """getattr のフォールバックが効くこと。"""
        class Bare:
            pass
        bare = Bare()
        # 必須属性だけ盛る
        bare.no_print_area_mode = "used_range"
        bare.value_mode = "display"
        bare.merge_policy = "top_left_only"
        bare.hyperlink_mode = "footnote"
        bare.header_detection = "first_row"
        bare.hidden_policy = "ignore"
        bare.strip_whitespace = True
        bare.escape_pipes = True
        bare.date_format_override = None
        bare.date_default_format = "YYYY-MM-DD"
        bare.numeric_thousand_sep = "keep"
        bare.percent_format = "keep"
        bare.currency_symbol = "keep"
        bare.percent_divide_100 = False
        bare.align_detection = "numbers_right"
        bare.numbers_right_threshold = 0.8
        bare.max_sheet_count = 0
        bare.max_cells_per_table = 200000
        bare.sort_tables = "document_order"
        bare.footnote_scope = "book"
        bare.locale = "ja-JP"
        bare.markdown_escape_level = "safe"
        bare.mermaid_enabled = False
        bare.mermaid_detect_mode = "shapes"
        bare.mermaid_direction = "TD"
        bare.mermaid_columns = "From,To,Label,Group,Note"
        bare.mermaid_heuristic_min_rows = 3
        bare.mermaid_heuristic_arrow_ratio = 0.3
        bare.mermaid_heuristic_len_median_ratio_min = 0.4
        bare.mermaid_heuristic_len_median_ratio_max = 2.5
        bare.prefer_excel_display = True
        bare.read_only = False
        # readonly_fill_policy, mermaid_diagram_type など、CLI に明示が無いものは
        # 省略して getattr のフォールバックを試す
        cfg = ConversionConfig.from_args(bare)
        assert cfg.readonly_fill_policy == "assume_no_fill"
        assert cfg.mermaid_diagram_type == "flowchart"
        assert cfg.csv_markdown_enabled is True


# =============================================================
# to_opts_dict — runner.run() inline 辞書との完全一致
# =============================================================

class TestOptsDictMatchesRunnerInline:
    """これが PR D を機械的置換にする保証。"""

    def test_defaults(self):
        args = build_argparser().parse_args(["dummy.xlsx"])
        cfg = ConversionConfig.from_args(args)
        assert cfg.to_opts_dict() == _build_opts_dict_inline(args)

    def test_custom_args(self):
        args = build_argparser().parse_args([
            "dummy.xlsx",
            "--hyperlink-mode", "inline",
            "--max-cells-per-table", "50",
            "--mermaid-enabled",
            "--footnote-scope", "sheet",
            "--no-strip-whitespace",
            "--mermaid-columns", "Src,Dst,Edge,Cat,Note",
            "--markdown-escape-level", "aggressive",
            "--no-csv-include-metadata",
        ])
        cfg = ConversionConfig.from_args(args)
        assert cfg.to_opts_dict() == _build_opts_dict_inline(args)

    @pytest.mark.parametrize("flag,value", [
        ("--value-mode", "formula"),
        ("--merge-policy", "expand"),
        ("--header-detection", "none"),
        ("--hidden-policy", "exclude"),
        ("--numeric-thousand-sep", "remove"),
        ("--percent-format", "numeric"),
        ("--currency-symbol", "strip"),
        ("--align-detection", "none"),
        ("--mermaid-detect-mode", "heuristic"),
        ("--mermaid-direction", "LR"),
        ("--mermaid-node-id-policy", "explicit"),
    ])
    def test_individual_choice_args(self, flag, value):
        args = build_argparser().parse_args(["dummy.xlsx", flag, value])
        cfg = ConversionConfig.from_args(args)
        assert cfg.to_opts_dict() == _build_opts_dict_inline(args)


# =============================================================
# 個別変換ルール (header_detection / align_detection / mermaid_columns)
# =============================================================

class TestStringToBoolCoercion:
    def test_header_detection_first_row(self):
        cfg = ConversionConfig(header_detection="first_row")
        assert cfg.header_detection == "first_row"
        assert cfg.to_opts_dict()["header_detection"] is True

    def test_header_detection_none(self):
        cfg = ConversionConfig(header_detection="none")
        assert cfg.to_opts_dict()["header_detection"] is False

    def test_header_detection_heuristic_is_not_true(self):
        """heuristic は CLI の選択肢にあるが、runner では first_row のみが True 扱い。"""
        cfg = ConversionConfig(header_detection="heuristic")
        assert cfg.to_opts_dict()["header_detection"] is False

    def test_align_detection_numbers_right(self):
        cfg = ConversionConfig(align_detection="numbers_right")
        assert cfg.to_opts_dict()["align_detection"] is True

    def test_align_detection_none(self):
        cfg = ConversionConfig(align_detection="none")
        assert cfg.to_opts_dict()["align_detection"] is False


class TestMermaidColumnsParsing:
    def test_all_five_fields(self):
        cfg = ConversionConfig(mermaid_columns="A,B,C,D,E")
        assert cfg.mermaid_columns == "A,B,C,D,E"  # 生文字列は保持される
        opts = cfg.to_opts_dict()
        assert opts["mermaid_columns"] == {
            "from": "A", "to": "B", "label": "C", "group": "D", "note": "E",
        }

    def test_three_fields_group_and_note_default_to_none(self):
        cfg = ConversionConfig(mermaid_columns="X,Y,Z")
        opts = cfg.to_opts_dict()
        assert opts["mermaid_columns"]["from"] == "X"
        assert opts["mermaid_columns"]["to"] == "Y"
        assert opts["mermaid_columns"]["label"] == "Z"
        assert opts["mermaid_columns"]["group"] is None
        assert opts["mermaid_columns"]["note"] is None

    def test_whitespace_trimmed(self):
        cfg = ConversionConfig(mermaid_columns="A , B , C")
        opts = cfg.to_opts_dict()
        assert opts["mermaid_columns"]["from"] == "A"
        assert opts["mermaid_columns"]["to"] == "B"
        assert opts["mermaid_columns"]["label"] == "C"

    def test_default_matches_cli(self):
        """既定値が CLI 側の既定 "From,To,Label,Group,Note" と同じ展開になる。"""
        cfg = ConversionConfig()
        opts = cfg.to_opts_dict()
        assert opts["mermaid_columns"] == {
            "from": "From", "to": "To", "label": "Label",
            "group": "Group", "note": "Note",
        }


# =============================================================
# detect_dates は runner.run でハードコードされている
# =============================================================

def test_detect_dates_hardcoded_true():
    cfg = ConversionConfig()
    assert cfg.to_opts_dict()["detect_dates"] is True
