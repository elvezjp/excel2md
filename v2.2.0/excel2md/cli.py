# -*- coding: utf-8 -*-
"""CLI entrypoints for excel2md.

仕様書参照: §3.1.2 主要設定オプション
"""

import argparse

from .runner import run
from . import __version__ as VERSION

def build_argparser():
    p = argparse.ArgumentParser(description=f"Excel -> Markdown converter (Spec v{VERSION})")
    p.add_argument("input", help="Path to .xlsx/.xlsm file")
    p.add_argument("-o", "--output", help="Path to output .md file")
    # Spec defaults
    p.add_argument("--no-print-area-mode", choices=["used_range","entire_sheet_range","skip_sheet"], default="used_range")
    p.add_argument("--value-mode", choices=["display","formula","both"], default="display")
    p.add_argument("--merge-policy", choices=["expand","repeat","warn","top_left_only"], default="top_left_only")
    p.add_argument("--hyperlink-mode", choices=["inline","inline_plain","footnote","both","text_only"], default="footnote")
    p.add_argument("--header-detection", choices=["none","first_row","heuristic"], default="first_row")
    p.add_argument("--mermaid-enabled", action="store_true", default=False)
    p.add_argument("--mermaid-detect-mode", choices=["none","column_headers","heuristic","shapes"], default="shapes")
    p.add_argument("--mermaid-diagram-type", choices=["flowchart","sequence","state"], default="flowchart")
    p.add_argument("--mermaid-direction", choices=["TD","LR","BT","RL"], default="TD")
    p.add_argument("--mermaid-keep-source-table", action="store_true", default=True)
    p.add_argument("--no-mermaid-keep-source-table", dest="mermaid_keep_source_table", action="store_false")
    p.add_argument("--mermaid-dedupe-edges", action="store_true", default=True)
    p.add_argument("--no-mermaid-dedupe-edges", dest="mermaid_dedupe_edges", action="store_false")
    p.add_argument("--mermaid-node-id-policy", choices=["auto","shape_id","explicit"], default="auto")
    p.add_argument("--mermaid-group-column-behavior", choices=["subgraph","ignore"], default="subgraph")
    p.add_argument("--mermaid-columns", default="From,To,Label,Group,Note")
    p.add_argument("--mermaid-heuristic-min-rows", type=int, default=3)
    p.add_argument("--mermaid-heuristic-arrow-ratio", type=float, default=0.3)
    p.add_argument("--mermaid-heuristic-len-median-ratio-min", type=float, default=0.4)
    p.add_argument("--mermaid-heuristic-len-median-ratio-max", type=float, default=2.5)
    p.add_argument("--dispatch-skip-code-and-mermaid-on-fallback", action="store_true", default=True)
    p.add_argument("--no-dispatch-skip-code-and-mermaid-on-fallback", dest="dispatch_skip_code_and_mermaid_on_fallback", action="store_false")

    p.add_argument("--hidden-policy", choices=["ignore","include","exclude"], default="ignore")
    p.add_argument("--strip-whitespace", action="store_true", default=True)
    p.add_argument("--no-strip-whitespace", dest="strip_whitespace", action="store_false")
    p.add_argument("--escape-pipes", action="store_true", default=True)
    p.add_argument("--no-escape-pipes", dest="escape_pipes", action="store_false")
    p.add_argument("--date-format-override", default=None)
    p.add_argument("--date-default-format", default="YYYY-MM-DD")
    p.add_argument("--numeric-thousand-sep", choices=["keep","remove"], default="keep")
    p.add_argument("--percent-format", choices=["keep","numeric"], default="keep")
    p.add_argument("--percent-divide-100", action="store_true", help="When percent-format=numeric, divide by 100 (e.g., 12%% -> 0.12)")
    p.add_argument("--currency-symbol", choices=["keep","strip"], default="keep")
    p.add_argument("--align-detection", choices=["none","numbers_right"], default="numbers_right")
    p.add_argument("--numbers-right-threshold", type=float, default=0.8)
    p.add_argument("--max-sheet-count", type=int, default=0)
    p.add_argument("--max-cells-per-table", type=int, default=200000)
    p.add_argument("--sort-tables", choices=["document_order"], default="document_order")
    p.add_argument("--footnote-scope", choices=["book","sheet"], default="book")
    p.add_argument("--locale", default="ja-JP")
    p.add_argument("--markdown-escape-level", choices=["safe","minimal","aggressive"], default="safe")
    p.add_argument("--read-only", action="store_true", help="Use openpyxl read_only=True (styles may be limited)")
    p.add_argument("--prefer-excel-display", action="store_true", default=True, help="Prefer Excel displayed value over formatted output when possible")

    # CSV Markdown output options
    p.add_argument("--csv-output-dir", default=None, help="CSV markdown output directory (default: same as input file)")
    p.add_argument("--csv-apply-merge-policy", action="store_true", default=True, help="Apply merge_policy to CSV data extraction")
    p.add_argument("--no-csv-apply-merge-policy", dest="csv_apply_merge_policy", action="store_false")
    p.add_argument("--csv-normalize-values", action="store_true", default=True, help="Apply numeric normalization to CSV values")
    p.add_argument("--no-csv-normalize-values", dest="csv_normalize_values", action="store_false")
    p.add_argument("--csv-markdown-enabled", action="store_true", default=True, help="Enable CSV markdown output (default: True)")
    p.add_argument("--no-csv-markdown-enabled", dest="csv_markdown_enabled", action="store_false")
    p.add_argument("--csv-include-metadata", action="store_true", default=True, help="Include validation metadata in CSV markdown output (default: True)")
    p.add_argument("--no-csv-include-metadata", dest="csv_include_metadata", action="store_false")
    p.add_argument("--csv-include-description", action="store_true", default=True, help="Include description section in CSV markdown output (default: True)")
    p.add_argument("--no-csv-include-description", dest="csv_include_description", action="store_false")

    # Image extraction options
    p.add_argument("--image-extraction", action="store_true", default=True, help="Enable image extraction from Excel (default: True)")
    p.add_argument("--no-image-extraction", dest="image_extraction", action="store_false")

    # Sheet output options
    p.add_argument("--split-by-sheet", action="store_true", default=False, help="Split output by sheet (generate separate file for each sheet)")

    return p

def main(argv=None):
    parser = build_argparser()
    args = parser.parse_args(argv)
    out = run(args.input, args.output, args)
    print(out)

if __name__ == "__main__":
    main()
