# -*- coding: utf-8 -*-
"""Markdown table formatting utilities.

仕様書参照: §4.3 テーブル形式判定フロー、§6 Markdown生成規約
"""

from typing import List, Tuple, Set

from .cell_utils import cell_display_value, numeric_like, normalize_numeric_text, has_border

def is_source_code(text: str) -> bool:
    """Check if text appears to be source code."""
    if not text or not text.strip():
        return False

    text_lower = text.lower()
    text_stripped = text.strip()

    # Check for code keywords
    code_keywords = [
        'public', 'private', 'protected', 'class', 'interface', 'import', 'package',
        'static', 'final', 'void', 'return', 'if', 'else', 'for', 'while', 'switch',
        'case', 'try', 'catch', 'throw', 'throws', 'extends', 'implements',
        'def', 'function', 'var', 'let', 'const', 'async', 'await',
        'namespace', 'using', 'namespace', 'struct', 'enum'
    ]

    # Check for code symbols
    code_symbols = ['{', '}', ';', '//', '/*', '*/']

    # Check for annotations (Java/C# style) - @ followed by identifier
    # Pattern: @ followed by alphanumeric characters (e.g., @Annotation1, @Override)
    has_annotation = False
    if '@' in text:
        # Check if @ is followed by identifier-like pattern
        import re
        annotation_pattern = r'@[a-zA-Z_][a-zA-Z0-9_]*'
        if re.search(annotation_pattern, text):
            has_annotation = True

    # Check for keywords
    has_keyword = any(kw in text_lower for kw in code_keywords)

    # Check for symbols
    has_symbol = any(sym in text for sym in code_symbols)

    # If it starts with @ and looks like annotation, likely code
    if text_stripped.startswith('@') and len(text_stripped) > 1 and text_stripped[1].isalnum():
        return True

    # If it has annotation marker and looks like code
    if has_annotation and (has_keyword or has_symbol):
        return True

    # If it has both keyword and symbol, likely code
    if has_keyword and has_symbol:
        return True

    # If it has multiple symbols, likely code
    symbol_count = sum(1 for sym in code_symbols if sym in text)
    if symbol_count >= 2:
        return True

    # If it has keyword alone (for single-line code snippets)
    if has_keyword and (has_symbol or has_annotation):
        return True

    return False

def is_code_block(md_rows) -> bool:
    """Check if rows appear to be a code block."""
    if not md_rows:
        return False

    code_lines = []
    is_code = False

    for row in md_rows:
        if not row:
            if is_code:
                code_lines.append("")
            continue

        # Get the first non-empty cell value
        row_text = ""
        for val in row:
            if val and str(val).strip():
                row_text = str(val).strip()
                break

        if row_text and is_source_code(row_text):
            code_lines.append(row_text)
            is_code = True
        elif is_code and row_text:
            # If we've started a code block, continue collecting lines
            code_lines.append(row_text)
        elif is_code and not row_text:
            # Empty line in code block - preserve it
            code_lines.append("")

    # If we detected code, return True
    return is_code and len(code_lines) > 0

def detect_code_language(lines: list) -> str:
    """Detect programming language from code lines."""
    if not lines:
        return ""

    # Combine all lines for analysis
    combined = " ".join(lines).lower()

    # Java indicators
    if any(kw in combined for kw in ['public class', 'private class', 'import java', '@override', '@annotation']):
        return "java"

    # Python indicators
    if any(kw in combined for kw in ['def ', 'import ', 'from ', 'if __name__', 'class ']) and ':' in combined:
        return "python"

    # JavaScript indicators
    if any(kw in combined for kw in ['function ', 'const ', 'let ', 'var ', '=>', 'export ', 'import ']):
        return "javascript"

    # C# indicators
    if any(kw in combined for kw in ['namespace ', 'using ', 'public class', '[attribute']):
        return "csharp"

    # C/C++ indicators
    if any(kw in combined for kw in ['#include', 'int main', 'printf', 'cout']):
        return "c"

    return ""

def build_code_block_from_rows(md_rows):
    """Build code block from rows if they look like source code; otherwise return None."""
    if not md_rows:
        return None

    code_lines = []
    is_code_block = False

    for row in md_rows:
        if not row:
            if is_code_block:
                # Empty line in code block - preserve it
                code_lines.append("")
            continue

        # Get the first non-empty cell value
        row_text = ""
        for val in row:
            if val and val.strip():
                row_text = val.strip()
                break

        if row_text and is_source_code(row_text):
            code_lines.append(row_text)
            is_code_block = True
        elif is_code_block and row_text:
            # If we've started a code block, continue collecting lines
            # even if they don't match code patterns (might be continuation)
            code_lines.append(row_text)
        elif is_code_block and not row_text:
            # Empty line in code block - preserve it
            code_lines.append("")

    if is_code_block and code_lines:
        language = detect_code_language(code_lines)
        code_block = "```" + (language if language else "") + "\n"
        code_block += "\n".join(code_lines)
        code_block += "\n```"
        return code_block

    return None


def format_table_as_text_or_nested(ws, table, md_rows, opts, merged_lookup):
    """Format table as text/nested format.

    Returns:
    - "text": Single line text format (1 cell in first row, no borders)
    - "nested": Nested format (indented text)
    - "code": Source code format (code block)
    - "table": Normal table format
    - "empty": Empty row (just newline)
    """
    if not md_rows:
        return "empty", ""

    min_row, min_col, max_row, max_col = table["bbox"]
    mask = table["mask"]

    # Get used columns by checking which columns have data
    # First, collect all columns that are in mask
    candidate_cols = set()
    for R in range(min_row, max_row + 1):
        for C in range(min_col, max_col + 1):
            if (R, C) in mask:
                candidate_cols.add(C)

    # Then, check which columns actually have non-empty cell values
    used_cols = set()
    for C in candidate_cols:
        has_data = False
        for R in range(min_row, max_row + 1):
            if (R, C) not in mask:
                continue
            cell = ws.cell(row=R, column=C)
            tl = merged_lookup.get((R, C))
            if tl:
                if (R, C) == tl:
                    cell = ws.cell(row=tl[0], column=tl[1])
                    text = cell_display_value(cell, opts)
                else:
                    if opts.get("merge_policy") == "top_left_only":
                        text = ""
                    else:
                        cell = ws.cell(row=tl[0], column=tl[1])
                        text = cell_display_value(cell, opts)
            else:
                text = cell_display_value(cell, opts)
            text = normalize_numeric_text(text, opts)
            if text and text.strip():
                has_data = True
                break
        if has_data:
            used_cols.add(C)

    used_cols = sorted(used_cols)
    if not used_cols:
        return "empty", ""

    # Check first row: single cell text format
    first_row = md_rows[0] if md_rows else []
    non_empty_in_first = [i for i, val in enumerate(first_row) if val and val.strip()]

    if len(non_empty_in_first) == 1:
        # Find the actual column index in used_cols
        first_row_col_idx = non_empty_in_first[0]
        if first_row_col_idx < len(used_cols):
            first_row_actual_col = used_cols[first_row_col_idx]

            # Check if the cell with value has no borders
            cell = ws.cell(row=min_row, column=first_row_actual_col)
            if not has_border(cell):
                # Check if other cells in first row are empty and have no borders
                all_other_empty_no_border = True
                for C in used_cols:
                    if C == first_row_actual_col:
                        continue
                    if (min_row, C) not in mask:
                        continue
                    cell = ws.cell(row=min_row, column=C)
                    text = cell_display_value(cell, opts)
                    text = normalize_numeric_text(text, opts)
                    if text and text.strip():
                        all_other_empty_no_border = False
                        break
                    if has_border(cell):
                        all_other_empty_no_border = False
                        break

                if all_other_empty_no_border:
                    # Single line text format
                    text_value = first_row[first_row_col_idx]
                    return "text", text_value

    # Check for source code format
    code_block = build_code_block_from_rows(md_rows)
    if code_block:
        return "code", code_block

    # Check for nested format
    nested_lines = []
    use_nested = True
    for row_idx, row in enumerate(md_rows):
        if not row:
            nested_lines.append("")
            continue

        non_empty = [i for i, val in enumerate(row) if val and val.strip()]

        if len(non_empty) == 0:
            nested_lines.append("")
        elif len(non_empty) == 1:
            if non_empty[0] == 0:
                # First column only - normal text (no indent)
                nested_lines.append(row[0])
            else:
                # First column is empty, second+ column has value - nested format
                first_non_empty_idx = non_empty[0]
                indent = "  " * first_non_empty_idx
                nested_lines.append(indent + row[first_non_empty_idx])
        else:
            # Multiple columns with values - check if it's still nested format
            # If all non-empty cells are in sequence starting from index > 0, it's nested
            if non_empty[0] > 0 and all(non_empty[i] == non_empty[0] + i for i in range(len(non_empty))):
                # Sequential columns starting from index > 0 - use first non-empty as nested
                first_non_empty_idx = non_empty[0]
                indent = "  " * first_non_empty_idx
                nested_lines.append(indent + row[first_non_empty_idx])
            else:
                # Multiple columns with values in different positions - use table format
                use_nested = False
                break

    # If all rows are nested format, return nested
    if use_nested and nested_lines:
        return "nested", "\n".join(nested_lines)

    return "table", None

def choose_header_row_heuristic(md_rows):
    """Pick header row using simple heuristics:
    - First row with >=50% non-empty AND less numeric_ratio than next row.
    - Fallback: first non-empty row.
    """
    if not md_rows:
        return None
    def numeric_ratio(row):
        vals = [x for x in row if (x or "").strip()]
        if not vals: return 0.0
        nums = sum(1 for v in vals if numeric_like(v))
        return nums/len(vals)
    first_nonempty = None
    for i, row in enumerate(md_rows[:3]):  # peek first up to 3 rows
        nonempty = sum(1 for v in row if (v or "").strip())
        if first_nonempty is None and nonempty>0:
            first_nonempty = i
        if nonempty >= max(1, len(row)//2):
            r_this = numeric_ratio(row)
            r_next = numeric_ratio(md_rows[i+1]) if i+1 < len(md_rows) else 1.0
            if r_this < r_next:  # header tends to be less numeric than data
                return i
    return first_nonempty if first_nonempty is not None else None

def detect_right_align(col_vals, threshold=0.8):
    """Detect if column should be right-aligned based on numeric ratio."""
    if not col_vals:
        return False
    non_empty = [v for v in col_vals if (v or "").strip()]
    if not non_empty:
        return False
    numeric_count = sum(1 for v in non_empty if numeric_like(str(v)))
    return (numeric_count / len(non_empty)) >= threshold

def make_markdown_table(md_rows, header_detection=True, align_detect=True, align_threshold=0.8):
    if not md_rows:
        return ""

    # Remove completely empty columns
    if md_rows:
        cols = len(md_rows[0])
        non_empty_cols = []
        for c in range(cols):
            # Check if column has any non-empty cell
            has_data = any((row[c] if c < len(row) else "").strip() for row in md_rows)
            if has_data:
                non_empty_cols.append(c)

        # Filter rows to only include non-empty columns
        if non_empty_cols:
            md_rows = [[row[c] if c < len(row) else "" for c in non_empty_cols] for row in md_rows]
        else:
            return ""  # All columns are empty

    header = None
    data = md_rows
    if header_detection and any((cell or "").strip() for cell in md_rows[0]):
        header = md_rows[0]
        data = md_rows[1:]
    cols = len(header) if header else len(md_rows[0])
    aligns = []
    if align_detect:
        for c in range(cols):
            col_vals = [row[c] for row in data if c < len(row)]
            right = detect_right_align(col_vals, threshold=align_threshold)
            aligns.append("---:" if right else "---")
    else:
        aligns = ["---"] * cols
    lines = []
    if header:
        lines.append("| " + " | ".join(header) + " |")
        lines.append("| " + " | ".join(aligns) + " |")
    for row in data if header else md_rows:
        if len(row) < cols:
            row = row + [""] * (cols - len(row))
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)

# ===== CSV Markdown Output Functions =====
