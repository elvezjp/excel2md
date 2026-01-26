# -*- coding: utf-8 -*-
"""Table extraction utilities.

仕様書参照: §5 セル・テーブル処理規則
"""

from typing import Dict, List, Set, Tuple

from .cell_utils import md_escape, cell_display_value, cell_is_empty, hyperlink_info, is_valid_url, normalize_numeric_text, has_border
from .output import warn
from .workbook_loader import a1_from_rc
from .mermaid_generator import is_flow_table, build_mermaid
from .table_formatting import make_markdown_table, format_table_as_text_or_nested, build_code_block_from_rows

def detect_table_title(ws, table, merged_lookup, opts, print_area=None):
    """Detect if table has a title from a large merged cell at the top-left.

    Args:
        ws: Worksheet object
        table: Table dict with 'bbox' and 'mask'
        merged_lookup: Merged cell lookup (only includes cells within print area)
        opts: Options dictionary
        print_area: Optional print area tuple (r0, c0, r1, c1) to ensure cells outside are excluded
    """
    min_row, min_col, max_row, max_col = table["bbox"]
    mask: Set[Tuple[int,int]] = table["mask"]

    if print_area is not None:
        area_r0, area_c0, area_r1, area_c1 = print_area
        min_row = max(min_row, area_r0)
        min_col = max(min_col, area_c0)
        max_row = min(max_row, area_r1)
        max_col = min(max_col, area_c1)

    for r in range(min_row, min(min_row + 3, max_row + 1)):
        for c in range(min_col, min(min_col + 10, max_col + 1)):
            if (r, c) not in mask:
                continue
            if print_area is not None:
                area_r0, area_c0, area_r1, area_c1 = print_area
                if r < area_r0 or r > area_r1 or c < area_c0 or c > area_c1:
                    continue
            tl = merged_lookup.get((r, c))
            if tl and tl == (r, c):  # This is a top-left of a merged cell
                # Get the merged range
                merged_cells = getattr(ws, "merged_cells", None)
                if merged_cells:
                    for rng in merged_cells.ranges:
                        if rng.min_row == r and rng.min_col == c:
                            if print_area is not None:
                                area_r0, area_c0, area_r1, area_c1 = print_area
                                if r < area_r0 or c < area_c0:
                                    continue
                                span_cols = min(rng.max_col, area_c1) - c + 1
                            else:
                                span_cols = rng.max_col - rng.min_col + 1
                            if span_cols >= 3 and r <= min_row + 2:
                                cell = ws.cell(row=r, column=c)
                                text = cell_display_value(cell, opts)
                                if text and text.strip():
                                    if print_area is not None:
                                        area_r0, area_c0, area_r1, area_c1 = print_area
                                        exclude_cols = set(range(c, min(rng.max_col, area_c1) + 1))
                                    else:
                                        exclude_cols = set(range(c, rng.max_col + 1))
                                    return text.strip(), exclude_cols
    return None, set()

def extract_table(ws, table, opts, footnotes, footnote_index_start, merged_lookup, print_area=None):
    """Extract Markdown rows for a logical table (possibly multiple rects).

    Args:
        ws: Worksheet object
        table: Table dict with 'bbox' and 'mask'
        opts: Options dictionary
        footnotes: Footnotes list
        footnote_index_start: Starting footnote index
        merged_lookup: Merged cell lookup (only includes cells within print area)
        print_area: Optional print area tuple (r0, c0, r1, c1) to ensure cells outside are excluded
    """
    min_row, min_col, max_row, max_col = table["bbox"]
    mask: Set[Tuple[int,int]] = table["mask"]

    if print_area is not None:
        area_r0, area_c0, area_r1, area_c1 = print_area
        # Filter mask to only include cells within print area
        mask = {(r, c) for (r, c) in mask if area_r0 <= r <= area_r1 and area_c0 <= c <= area_c1}
        # Adjust bbox to be within print area
        min_row = max(min_row, area_r0)
        min_col = max(min_col, area_c0)
        max_row = min(max_row, area_r1)
        max_col = min(max_col, area_c1)

    # Detect table title from large merged cell
    # Create updated table dict with filtered mask and adjusted bbox
    updated_table = {"bbox": (min_row, min_col, max_row, max_col), "mask": mask}
    table_title, title_cols = detect_table_title(ws, updated_table, merged_lookup, opts, print_area)

    # Find columns that actually have data (non-empty cells in mask), excluding title columns
    # First, collect all columns that are in mask and not in title_cols
    candidate_cols = set()
    for R in range(min_row, max_row+1):
        for C in range(min_col, max_col+1):
            if (R, C) in mask and C not in title_cols:
                candidate_cols.add(C)

    # Then, check which columns actually have non-empty cell values
    # Note: We need to apply the same processing as in extract_table to get the final cell values
    # This includes: merged cell handling, numeric formatting, hyperlink processing, and markdown escaping
    used_cols = set()
    for C in candidate_cols:
        has_data = False
        for R in range(min_row, max_row+1):
            if (R, C) not in mask:
                continue
            cell = ws.cell(row=R, column=C)
            # Check merged cell handling (same as in extract_table)
            tl = merged_lookup.get((R, C))
            if tl:
                if (R, C) == tl:
                    # Top-left cell: check its value
                    cell = ws.cell(row=tl[0], column=tl[1])
                    text = cell_display_value(cell, opts)
                else:
                    # Other cells in merged range: empty if top_left_only
                    if opts.get("merge_policy") == "top_left_only":
                        text = ""
                    else:
                        cell = ws.cell(row=tl[0], column=tl[1])
                        text = cell_display_value(cell, opts)
            else:
                text = cell_display_value(cell, opts)

            # Apply the same processing as in extract_table
            # numeric formatting overrides
            text = normalize_numeric_text(text, opts)

            # hyperlink processing (but we only need to check if there's data, not format it)
            hl = hyperlink_info(cell)
            if hl:
                disp = text if text else (hl.get("display") or "")
                if disp and disp.strip():
                    text = disp  # Use display text for checking

            # Check if text is non-empty after processing
            if text and text.strip():
                has_data = True
                break
        if has_data:
            used_cols.add(C)

    # Sort columns
    used_cols = sorted(used_cols)
    if not used_cols:
        return [], [], False, table_title

    md_rows = []
    note_refs = []
    max_cells = opts["max_cells_per_table"]
    count_cells = 0

    for R in range(min_row, max_row+1):
        if print_area is not None:
            area_r0, area_c0, area_r1, area_c1 = print_area
            if R < area_r0 or R > area_r1:
                continue
        row_is_empty = True
        for C in used_cols:
            if (R, C) not in mask:
                continue
            if print_area is not None:
                area_r0, area_c0, area_r1, area_c1 = print_area
                if R < area_r0 or R > area_r1 or C < area_c0 or C > area_c1:
                    continue  # Cell is outside print area
            cell = ws.cell(row=R, column=C)
            tl = merged_lookup.get((R, C))
            if tl:
                if (R, C) == tl:
                    cell = ws.cell(row=tl[0], column=tl[1])
                    text = cell_display_value(cell, opts)
                    if text and text.strip():
                        row_is_empty = False
                        break
                else:
                    if opts.get("merge_policy") != "top_left_only":
                        cell = ws.cell(row=tl[0], column=tl[1])
                        text = cell_display_value(cell, opts)
                        if text and text.strip():
                            row_is_empty = False
                            break
            else:
                text = cell_display_value(cell, opts)
                if text and text.strip():
                    row_is_empty = False
                    break

        # Skip completely empty rows
        if row_is_empty:
            continue

        row_vals = []
        for C in used_cols:
            count_cells += 1
            if count_cells > max_cells:
                return md_rows, note_refs, True
            if (R,C) not in mask:
                row_vals.append("")
                continue
            if print_area is not None:
                area_r0, area_c0, area_r1, area_c1 = print_area
                if R < area_r0 or R > area_r1 or C < area_c0 or C > area_c1:
                    row_vals.append("")
                    continue
            cell = ws.cell(row=R, column=C)
            # merged handling
            tl = merged_lookup.get((R,C))
            if tl:
                # Check if this is the top-left cell of the merged range
                if (R, C) == tl:
                    # Top-left cell: use its value
                    cell = ws.cell(row=tl[0], column=tl[1])
                    text = cell_display_value(cell, opts)
                else:
                    # Other cells in merged range: empty if top_left_only, otherwise use top-left value
                    if opts.get("merge_policy") == "top_left_only":
                        text = ""
                    else:
                        # expand/repeat: use top-left value
                        cell = ws.cell(row=tl[0], column=tl[1])
                        text = cell_display_value(cell, opts)
            else:
                text = cell_display_value(cell, opts)
            # numeric formatting overrides
            text = normalize_numeric_text(text, opts)

            hl = hyperlink_info(cell)
            if hl:
                disp = text if text else (hl.get("display") or "")
                if hl.get("target"):
                    link = hl["target"]
                    if not is_valid_url(link):
                        warn(f"Invalid URL detected at {a1_from_rc(R,C)}: {link}")
                    if opts["hyperlink_mode"] in ("inline", "both"):
                        text = f"[{disp}]({link})"
                    elif opts["hyperlink_mode"] == "inline_plain":
                        text = f"{disp} ({link})"
                    if opts["hyperlink_mode"] in ("footnote", "both"):
                        n = footnote_index_start + len(note_refs)
                        note_refs.append((n, link))
                        text = f"{text}[^{n}]"
                elif hl.get("location"):
                    loc = hl["location"]
                    if opts["hyperlink_mode"] in ("inline", "both"):
                        text = f"[{disp}]({loc})"
                    elif opts["hyperlink_mode"] == "inline_plain":
                        text = f"{disp} (→{loc})"
                    elif opts["hyperlink_mode"] in ("footnote", "both"):
                        n = footnote_index_start + len(note_refs)
                        note_refs.append((n, loc))
                        text = f"{disp}[^{n}]"

            text = md_escape(text, opts["markdown_escape_level"])
            row_vals.append(text)
        md_rows.append(row_vals)
    return md_rows, note_refs, False, table_title

def dispatch_table_output(ws, tbl, md_rows, opts, merged_lookup, xlsx_path=None):
    """Unified dispatcher: code -> mermaid -> text -> nested -> table

    Args:
        ws: Worksheet object
        tbl: Table dict
        md_rows: Markdown rows
        opts: Options dictionary
        merged_lookup: Merged cell lookup
        xlsx_path: Path to Excel file (not used, kept for compatibility)
    """
    skip_on_fallback = opts.get("dispatch_skip_code_and_mermaid_on_fallback", True)
    code_failed = False

    # 1) Code (最優先)
    try:
        code_block = build_code_block_from_rows(md_rows)
        if code_block:
            return "code", code_block
    except Exception as e:
        warn(f"code detection failed: {e}")
        code_failed = True
        if skip_on_fallback:
            # Skip Mermaid detection and go directly to priority 3
            pass
        else:
            # Continue to Mermaid detection
            pass

    # 2) Mermaid (skip if code failed and skip_on_fallback is True)
    if not (code_failed and skip_on_fallback):
        try:
            if opts.get("mermaid_enabled", False):
                detect_mode = opts.get("mermaid_detect_mode", "shapes")

                # 2a) Table-based modes (column_headers/heuristic)
                # Note: shapes mode is handled at sheet level, not here
                if detect_mode in ("column_headers", "heuristic"):
                    ok, colmap = is_flow_table(md_rows, opts)
                    if ok:
                        mer = build_mermaid(md_rows, opts, colmap)
                        if opts.get("mermaid_keep_source_table", True):
                            hdr = opts.get("header_detection", True)
                            table_md = make_markdown_table(md_rows, header_detection=hdr, align_detect=opts.get("align_detection", True), align_threshold=opts.get("numbers_right_threshold", 0.8))
                            return "mermaid", mer + "\n" + table_md
                        else:
                            return "mermaid", mer
        except Exception as e:
            warn(f"Mermaid generation failed: {e}")
            # fallthrough to (3) - フォールバックは優先度3（単一行）から開始

    # 3-5) delegate to existing logic (フォールバック時は優先度3から)
    ftype, out = format_table_as_text_or_nested(ws, tbl, md_rows, opts, merged_lookup)
    if ftype == "table":
        hdr = opts.get("header_detection", True)
        out = make_markdown_table(md_rows, header_detection=hdr, align_detect=opts.get("align_detection", True), align_threshold=opts.get("numbers_right_threshold", 0.8))
    elif ftype in ("text","nested","code","empty"):
        pass
    else:
        # fallback to normal table
        hdr = opts.get("header_detection", True)
        out = make_markdown_table(md_rows, header_detection=hdr, align_detect=opts.get("align_detection", True), align_threshold=opts.get("numbers_right_threshold", 0.8))
        ftype = "table"
    return ftype, out
