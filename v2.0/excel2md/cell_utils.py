# -*- coding: utf-8 -*-
"""Cell and text utilities.

仕様書参照: §5 セル・テーブル処理規則
"""

import re
import unicodedata

MD_RESERVED_SAFE = r"\\|\||\*|_|~|#|>|\[|\]|\(|\)|\{|\}|\+|\-|\.|!|`"
MD_ESCAPE_RE = re.compile(MD_RESERVED_SAFE)

NUMERIC_PATTERN = re.compile(
    r"""
^\s*
(\()?                                   # optional opening parenthesis
(?P<sign>[+-])?                         # optional sign
(?P<currency>[¥$€£₩])?                  # optional currency
(?P<int>\d{1,3}(?:[,，]\d{3})*|\d+)     # integer part
(?:[\.．](?P<frac>\d+))?                # optional fractional part
(?:[eE](?P<exp>[+-]?\d+))?              # optional exponent
(\))?                                   # optional closing parenthesis
\s*(?P<pct>%)?
\s*$
""",
    re.VERBOSE,
)

WHITESPACE_CHARS = {
    "\t", "\u0020", "\u00A0", "\u1680", "\u180E", "\u2000", "\u2001",
    "\u2002", "\u2003", "\u2004", "\u2005", "\u2006", "\u2007",
    "\u2008", "\u2009", "\u200A", "\u202F", "\u205F", "\u3000",
}

CONTROL_REMOVE_RANGES = [
    (0x0000, 0x0008),
    (0x000B, 0x000C),
    (0x000E, 0x001F),
    (0x007F, 0x007F),
    (0x0080, 0x009F),
    (0x00AD, 0x00AD),
    (0x200B, 0x200D),
    (0x2060, 0x2060),
    (0x2066, 0x2069),
]


def remove_control_chars(text: str) -> str:
    out = []
    for ch in text:
        code = ord(ch)
        if any(lo <= code <= hi for lo, hi in CONTROL_REMOVE_RANGES):
            continue
        out.append(ch)
    return "".join(out)


def is_whitespace_only(s: str) -> bool:
    return all((c in WHITESPACE_CHARS) for c in s) and len(s) > 0


def md_escape(s: str, level: str = "safe") -> str:
    if not s:
        return s
    s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
    s = s.replace("|", r"\|")
    if level == "safe":
        s = MD_ESCAPE_RE.sub(lambda m: "\\" + m.group(0), s)
    elif level == "aggressive":
        s = re.sub(r".", lambda m: "\\" + m.group(0), s)
    return s


def no_fill(cell, readonly_fill_policy: str = "assume_no_fill") -> bool:
    """Check if cell has no fill or white fill (white fill is treated as no fill)."""
    fill = getattr(cell, "fill", None)
    if fill is None:
        # Style may be unavailable in read_only. Use policy.
        return True if readonly_fill_policy == "assume_no_fill" else False
    patternType = getattr(fill, "patternType", None)
    if patternType in (None, "none"):
        return True
    if patternType == "solid":
        try:
            fg = fill.fgColor
            bg = fill.bgColor

            # Check foreground color
            fg_rgb = None
            try:
                fg_rgb = getattr(fg, "rgb", None)
                # Check if fg_rgb is actually a valid RGB value (not an error message)
                if fg_rgb and isinstance(fg_rgb, str) and len(fg_rgb) >= 6:
                    fg_rgb_str = fg_rgb[-6:].upper()
                    if fg_rgb_str == "FFFFFF":
                        # White foreground - check if background is also white or no fill
                        bg_rgb = getattr(bg, "rgb", None)
                        if bg_rgb and isinstance(bg_rgb, str) and len(bg_rgb) >= 6:
                            try:
                                bg_rgb_str = bg_rgb[-6:].upper()
                                if bg_rgb_str == "FFFFFF":
                                    return True
                            except Exception:
                                pass
                        # Check indexed background
                        try:
                            bg_type = getattr(bg, 'type', None)
                            if bg_type == 'indexed' and hasattr(bg, 'indexed'):
                                bg_indexed = bg.indexed
                                if bg_indexed == 64:
                                    return True
                        except Exception:
                            pass
                        # If foreground is white but background is not, it's still considered white
                        return True
                    else:
                        # Any other foreground color means the cell has fill (not empty)
                        return False
            except Exception:
                pass

            # Check indexed foreground colors
            fg_indexed = None
            try:
                fg_type = getattr(fg, 'type', None)
                if fg_type == 'indexed' and hasattr(fg, 'indexed'):
                    fg_indexed = fg.indexed
            except Exception:
                pass

            if fg_indexed == 64:
                # White foreground - check background
                try:
                    bg_type = getattr(bg, 'type', None)
                    if bg_type == 'indexed' and hasattr(bg, 'indexed'):
                        bg_indexed = bg.indexed
                        if bg_indexed == 64 or bg_indexed is None:
                            return True
                    else:
                        return True
                except Exception:
                    return True

            # Check background color (if foreground is theme, not set, or white)
            bg_rgb = getattr(bg, "rgb", None)
            if bg_rgb:
                try:
                    bg_rgb_str = str(bg_rgb)[-6:].upper() if len(str(bg_rgb)) >= 6 else str(bg_rgb).upper()
                    if bg_rgb_str == "FFFFFF":
                        return True
                except Exception:
                    pass

            # Check indexed background colors
            try:
                bg_type = getattr(bg, 'type', None)
                if bg_type == 'indexed' and hasattr(bg, 'indexed'):
                    bg_indexed = bg.indexed
                    if bg_indexed == 64:
                        return True
                    elif bg_indexed is not None:
                        return False
            except Exception:
                pass

        except Exception:
            return False
    return False


def excel_is_date(cell) -> bool:
    try:
        return cell.is_date
    except Exception:
        return False


def format_value(cell, opts) -> str:
    """Display value with basic formatting."""
    v = cell.value
    if v is None:
        return ""
    # Dates
    if opts.get("detect_dates", True) and excel_is_date(cell):
        import datetime as _dt
        if isinstance(v, (_dt.datetime, _dt.date)):
            fmt = opts.get("date_format_override") or opts.get("date_default_format", "YYYY-MM-DD")
            # Simple tokens: map to strftime
            fmt_map = {
                "YYYY": "%Y", "MM": "%m", "DD": "%d",
                "HH": "%H", "mm": "%M", "ss": "%S",
            }
            pyfmt = fmt
            for k, f in fmt_map.items():
                pyfmt = pyfmt.replace(k, f)
            return v.strftime(pyfmt)
    # Strings
    if isinstance(v, str):
        return v
    # Numbers & others
    s = str(v)
    return s


def cell_display_value(cell, opts) -> str:
    text = format_value(cell, opts)
    if isinstance(text, str):
        text = unicodedata.normalize("NFC", text)
        text = remove_control_chars(text)
        if opts["strip_whitespace"]:
            text = text.strip()
    return "" if text is None else str(text)


def cell_is_empty(cell, opts) -> bool:
    text = cell_display_value(cell, opts)
    if text == "" or is_whitespace_only(text):
        if no_fill(cell, opts.get("readonly_fill_policy", "assume_no_fill")):
            return True
    return False


def numeric_like(s: str) -> bool:
    m = NUMERIC_PATTERN.match(s)
    if not m:
        return False
    open_paren = m.group(1)
    groups = m.groups()
    close_paren = groups[-2] if len(groups) >= 2 else None
    if open_paren and not close_paren:
        return False
    if close_paren and not open_paren:
        return False
    return True


def normalize_numeric_text(s: str, opts) -> str:
    """Apply percent/currency/thousand settings to a numeric-like string for output control."""
    if not numeric_like(s):
        return s
    raw = s
    # Currency strip (start or surrounded by spaces)
    if opts.get("currency_symbol", "keep") == "strip":
        raw = re.sub(r"^\s*[¥$€£₩]\s*", "", raw)
    # Remove grouping if requested (only digits separators)
    if opts.get("numeric_thousand_sep", "keep") == "remove":
        raw = raw.replace(",", "").replace("，", "")
    # Percent handling
    if opts.get("percent_format", "keep") == "numeric":
        has_pct = raw.strip().endswith("%")
        raw = raw.replace("%", "")
        if has_pct and opts.get("percent_divide_100"):
            try:
                raw_num = float(raw)
                raw = str(raw_num / 100.0)
            except Exception:
                pass
    return raw


def hyperlink_info(cell):
    try:
        hl = cell.hyperlink
    except Exception:
        return None
    if not hl:
        return None
    target = getattr(hl, "target", None)
    location = getattr(hl, "location", None)
    display = getattr(hl, "display", None)
    return {"target": target, "location": location, "display": display}


def is_valid_url(target: str) -> bool:
    return bool(re.match(r"^(https?://|mailto:|file://|/|\.{1,2}/)", target or ""))

def has_border(cell):
    """Check if cell has any border lines.

    For table format detection, we consider a cell to have borders only if it has
    borders on multiple sides (not just one side), as single-side borders are often
    used for visual separation rather than table structure.
    """
    try:
        border = cell.border
        borders = [
            border.left.style is not None and border.left.style != 'none',
            border.right.style is not None and border.right.style != 'none',
            border.top.style is not None and border.top.style != 'none',
            border.bottom.style is not None and border.bottom.style != 'none'
        ]
        # Consider it a border only if 2 or more sides have borders
        return sum(borders) >= 2
    except:
        return False
