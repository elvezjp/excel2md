# -*- coding: utf-8 -*-
"""Mermaid generation utilities.

仕様書参照: §7 Mermaid生成規約
"""

from .output import warn
from .image_extraction import _DRAWINGML_NS

import re
import re as _re
import unicodedata as _unicodedata
import zipfile as _zipfile
import xml.etree.ElementTree as _ET
from typing import Dict, Tuple, Optional, List, Set, Any

def _v14_normalize_header_name(s: str) -> str:
    s = s or ""
    s = _unicodedata.normalize("NFKC", s)
    s = s.casefold()
    s = " ".join(s.split())
    return s

def _v14_sanitize_node_id(name: str) -> str:
    s = _unicodedata.normalize("NFKC", name or "")
    # keep only ascii letters, digits, underscore. Replace others with underscore.
    s = "".join(ch if ("a"<=ch<="z" or "A"<=ch<="Z" or "0"<=ch<="9" or ch=="_") else "_" for ch in s)
    if not s:
        s = "_"
    if s[0].isdigit():
        s = "_" + s
    # collapse multiple underscores
    s = _re.sub(r"_+", "_", s)
    return s

def _v14_resolve_columns_by_name(header_row, mapping):
    norm_map = {k: _v14_normalize_header_name(v) for k, v in mapping.items() if v}
    index = {}
    dups = set()
    for i, name in enumerate(header_row):
        key = _v14_normalize_header_name(str(name or ""))
        for want_k, want in norm_map.items():
            if key == want and want_k not in index:
                index[want_k] = i
            elif key == want and want_k in index:
                dups.add(want_k)
    if dups:
        warn(f"duplicate header names detected for {sorted(list(dups))}: first occurrence will be used")
    if "from" not in index or "to" not in index:
        return None
    return index

def _v14_extract_shapes_to_mermaid(xlsx_path: str, ws, opts) -> Optional[str]:
    """Extract shapes from DrawingML and build Mermaid flowchart.

    Args:
        xlsx_path: Path to Excel file
        ws: Worksheet object
        opts: Options dictionary

    Returns:
        Mermaid code block string or None if no shapes detected
    """
    try:
        # Build cell grid {(row, col): value} (0-based)
        grid: Dict[Tuple[int, int], str] = {}
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            for c_idx, val in enumerate(row, start=1):
                if val is not None and str(val).strip():
                    grid[(r_idx-1, c_idx-1)] = str(val).strip()  # 0-based

        # Open Excel as ZIP and find DrawingML files
        z = _zipfile.ZipFile(xlsx_path, 'r')

        # Find the drawing file corresponding to this sheet
        # 1. Get sheet name from worksheet object
        sheet_name = ws.title

        # 2. Find sheet ID from workbook.xml
        sheet_id = None
        try:
            wb_xml = z.read('xl/workbook.xml').decode('utf-8')
            wb_root = _ET.fromstring(wb_xml)
            ns_wb = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                     'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
            sheets = wb_root.findall('.//main:sheet', ns_wb)
            for sheet in sheets:
                if sheet.get('name') == sheet_name:
                    # Get rId from relationship
                    r_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if r_id:
                        # Find sheet index from workbook.xml.rels
                        wb_rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
                        wb_rels_root = _ET.fromstring(wb_rels_xml)
                        ns_rel = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                        for rel in wb_rels_root.findall('.//r:Relationship', ns_rel):
                            if rel.get('Id') == r_id:
                                target = rel.get('Target')
                                # Extract sheet number from target (e.g., "worksheets/sheet1.xml" -> "1")
                                if 'sheet' in target:
                                    # Find the number after "sheet" and before ".xml"
                                    match = re.search(r'sheet(\d+)\.xml', target)
                                    if match:
                                        sheet_id = match.group(1)
                                break
                    break
        except Exception as e:
            warn(f"Failed to find sheet ID for '{sheet_name}': {e}")

        # 3. Find drawing file from sheet relationship file
        drawing_path = None
        if sheet_id:
            sheet_rel_path = f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels"
            if sheet_rel_path in z.namelist():
                try:
                    rels_xml = z.read(sheet_rel_path)
                    rels_root = _ET.fromstring(rels_xml)
                    ns_rel = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                    for rel in rels_root.findall('.//r:Relationship', ns_rel):
                        if rel.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing':
                            target = rel.get('Target')
                            # Resolve relative path (xl/worksheets/基準)
                            if target.startswith('../'):
                                drawing_path = target.replace('../', 'xl/')
                            else:
                                drawing_path = f"xl/worksheets/{target}"
                            break
                except Exception as e:
                    warn(f"Failed to parse relationship file '{sheet_rel_path}': {e}")

        if not drawing_path or drawing_path not in z.namelist():
            z.close()
            return None  # No drawing file for this sheet

        # 4. Parse the corresponding DrawingML file
        xml_data = z.read(drawing_path)
        root = _ET.fromstring(xml_data)

        nodes: List[Dict[str, Any]] = []
        edges: List[Dict[str, Any]] = []
        tmp_id = 1

        def get_id(cnv):
            if cnv is not None and "id" in cnv.attrib:
                return f"s{cnv.attrib['id']}"
            nonlocal tmp_id
            tmp_id += 1
            return f"local{tmp_id}"

        def get_text_from_txBody(sp):
            """Get text from txBody element.

            Note: In Excel DrawingML, text is stored in xdr:txBody (not a:txBody).
            The structure is: xdr:txBody -> a:p -> a:r -> a:t
            """
            texts = []
            # First try xdr:txBody (correct structure in Excel)
            for txbody in sp.findall(".//xdr:txBody", _DRAWINGML_NS):
                for p in txbody.findall(".//a:p", _DRAWINGML_NS):
                    runs = []
                    for r in p.findall(".//a:r", _DRAWINGML_NS):
                        t = r.find("a:t", _DRAWINGML_NS)
                        if t is not None and t.text:
                            runs.append(t.text)
                    if runs:
                        texts.append("".join(runs))
            # Fallback: try a:txBody (for compatibility)
            if not texts:
                for txbody in sp.findall(".//a:txBody", _DRAWINGML_NS):
                    for p in txbody.findall(".//a:p", _DRAWINGML_NS):
                        runs = []
                        for r in p.findall(".//a:r", _DRAWINGML_NS):
                            t = r.find("a:t", _DRAWINGML_NS)
                            if t is not None and t.text:
                                runs.append(t.text)
                        if runs:
                            texts.append("".join(runs))
            return "\n".join([t.strip() for t in texts if t and t.strip()])

        def cell_bbox(anc):
            def int_or_none(e):
                return int(e.text) if e is not None and e.text and e.text.isdigit() else None
            fr_c = anc.find(".//xdr:from/xdr:col", _DRAWINGML_NS)
            fr_r = anc.find(".//xdr:from/xdr:row", _DRAWINGML_NS)
            to_c = anc.find(".//xdr:to/xdr:col", _DRAWINGML_NS)
            to_r = anc.find(".//xdr:to/xdr:row", _DRAWINGML_NS)
            if all(x is not None for x in (fr_c, fr_r, to_c, to_r)):
                return [int_or_none(fr_r), int_or_none(fr_c), int_or_none(to_r), int_or_none(to_c)]
            return None

        def text_from_bbox(bbox):
            if not bbox or None in bbox:
                return ''
            r1, c1, r2, c2 = bbox
            pad = 1
            texts = []
            for r in range(max(0, r1-pad), r2+pad+1):
                for c in range(max(0, c1-pad), c2+pad+1):
                    v = grid.get((r, c))
                    if v:
                        texts.append(v)
            if texts:
                seen = set()
                uniq = [t for t in texts if t not in seen and not seen.add(t)]
                longest = max(uniq, key=len)
                return longest if len(longest) >= 4 else "\n".join(uniq[:4])
            return ''

        # Scan anchors
        anchors = root.findall(".//xdr:twoCellAnchor", _DRAWINGML_NS) + root.findall(".//xdr:oneCellAnchor", _DRAWINGML_NS)

        for anc in anchors:
            sp = anc.find(".//xdr:sp", _DRAWINGML_NS)
            cxn = anc.find(".//xdr:cxnSp", _DRAWINGML_NS)
            bbox = cell_bbox(anc)

            if sp is not None:
                cNvPr = sp.find(".//xdr:cNvPr", _DRAWINGML_NS)
                _id = get_id(cNvPr)
                name = cNvPr.get("name") if cNvPr is not None else None

                # テキスト取得（優先順位1: txBody、優先順位2: bbox近傍、優先順位3: 名前）
                text = get_text_from_txBody(sp)
                if not text:
                    text = text_from_bbox(bbox)
                if not text:
                    text = name or ""

                # 形状種類の判定
                prstGeom = sp.find(".//a:prstGeom", _DRAWINGML_NS)
                prst = prstGeom.get("prst") if prstGeom is not None else None
                shape_type = "process"  # デフォルト
                if prst:
                    if prst in ("flowChartDecision", "diamond"):
                        shape_type = "decision"
                    elif prst in ("flowChartTerminator", "ellipse", "roundRect"):
                        shape_type = "terminator"
                    elif prst in ("flowChartInputOutput", "trapezoid"):
                        shape_type = "io"
                    elif prst in ("flowChartPreparation", "hexagon"):
                        shape_type = "prep"
                    elif prst == "flowChartManualOperation":
                        shape_type = "manual"
                    elif prst == "flowChartDocument":
                        shape_type = "document"
                    elif prst == "flowChartConnector":
                        shape_type = "connector"

                nodes.append({"id": _id, "name": name, "text": text, "bbox": bbox, "type": shape_type})
            elif cxn is not None:
                st = cxn.find(".//a:stCxn", _DRAWINGML_NS)
                ed = cxn.find(".//a:endCxn", _DRAWINGML_NS)
                st_id = f"s{st.get('id')}" if st is not None and st.get("id") else None
                ed_id = f"s{ed.get('id')}" if ed is not None and ed.get("id") else None
                if st_id and ed_id:
                    edges.append({"from": st_id, "to": ed_id})

        z.close()

        if not nodes:
            return None

        # Infer edges if insufficient
        valid_edges = [e for e in edges if e.get('from') and e.get('to')]
        if len(valid_edges) < max(2, int(0.3*len(nodes))):
            edges = valid_edges + _v14_infer_edges(nodes, edges)

        # Build Mermaid code
        direction = opts.get("mermaid_direction", "TD")
        node_id_policy = opts.get("mermaid_node_id_policy", "auto")

        def escape_mermaid_text(text: str) -> str:
            """Mermaid表示名の特殊文字をHTMLエンティティに変換"""
            if not text:
                return ""
            # 主要な特殊文字をHTMLエンティティに変換
            text = text.replace("[", "&#91;")
            text = text.replace("]", "&#93;")
            text = text.replace("{", "&#123;")
            text = text.replace("}", "&#125;")
            text = text.replace("|", "&#124;")
            text = text.replace('"', "&quot;")
            text = text.replace("'", "&#39;")
            return text

        def format_node(nid: str, display: str, shape_type: str) -> str:
            """シェイプ種類に応じたMermaidノード形式を生成"""
            display_escaped = escape_mermaid_text(display)
            if shape_type == "decision":
                return f'  {nid}{{"{display_escaped}"}}'
            elif shape_type == "terminator":
                return f'  {nid}(["{display_escaped}"])'
            elif shape_type == "io":
                return f'  {nid}[("{display_escaped}")]'
            elif shape_type == "prep":
                return f'  {nid}[{{"{display_escaped}"}}]'
            elif shape_type in ("manual", "document"):
                return f'  {nid}[("{display_escaped}")]'
            else:  # process, connector, その他
                return f'  {nid}["{display_escaped}"]'

        lines = []
        lines.append("```mermaid")
        lines.append(f"flowchart {direction}")

        # Map original IDs to sanitized IDs
        node_map = {}
        for n in nodes:
            display = n["text"] or n["name"] or n["id"]
            shape_type = n.get("type", "process")

            # ノードID生成
            if node_id_policy == "shape_id":
                # Excel描画IDをそのまま使用（s{id}形式）
                nid = n["id"]  # 既に "s{id}" 形式で格納されている
            elif node_id_policy == "auto":
                # 表示名をサニタイズしてID化
                nid = _v14_sanitize_node_id(str(display))
                # Handle duplicates
                base_nid = nid
                counter = 2
                while nid in node_map.values():
                    nid = f"{base_nid}_{counter}"
                    counter += 1
            else:  # explicit（将来拡張）
                nid = n["id"]

            node_map[n["id"]] = nid
            lines.append(format_node(nid, display, shape_type))

        # Add edges
        for e in edges:
            from_id = node_map.get(e.get("from"))
            to_id = node_map.get(e.get("to"))
            if from_id and to_id:
                # 推論されたエッジかどうかで形式を変更（任意）
                inferred = e.get("inferred", False)
                if inferred:
                    lines.append(f'  {from_id} -.->|inferred| {to_id}')
                else:
                    lines.append(f'  {from_id} --> {to_id}')

        lines.append("```")
        return "\n".join(lines)

    except Exception as e:
        warn(f"Shape extraction failed: {e}")
        return None

def _v14_infer_edges(nodes: List[Dict], exist_edges: List[Dict], max_out: int = 2, v_bias: float = 1.0) -> List[Dict]:
    """Infer edges based on vertical proximity."""
    def center_of(bbox):
        if not bbox or None in bbox:
            return None
        r1, c1, r2, c2 = bbox
        return ((r1 + r2) / 2.0, (c1 + c2) / 2.0)

    node_list = [n for n in nodes if n.get("bbox")]
    centers = {n["id"]: center_of(n.get("bbox")) for n in node_list}
    existing_pairs = {(e.get("from"), e.get("to")) for e in exist_edges if e.get("from") and e.get("to")}
    new_edges = []

    for n in node_list:
        nid = n["id"]
        c = centers.get(nid)
        if not c:
            continue
        dists = []
        for m in node_list:
            mid = m["id"]
            if mid == nid:
                continue
            mc = centers.get(mid)
            if not mc:
                continue
            dr = mc[0] - c[0]
            dc = abs(mc[1] - c[1])
            if dr >= 0:  # prefer downward flow
                dist = dr * v_bias + dc
                dists.append((dist, mid))
        dists.sort(key=lambda x: x[0])
        out = 0
        for _, mid in dists:
            if out >= max_out:
                break
            pair = (nid, mid)
            if pair in existing_pairs:
                continue
            new_edges.append({"from": nid, "to": mid, "bbox": None, "inferred": True})
            existing_pairs.add(pair)
            out += 1

    return new_edges

def is_flow_table(md_rows, opts):
    """Return (True, colmap) if table looks like flow table. colmap may be None for heuristic."""
    detect_mode = opts.get("mermaid_detect_mode","shapes")
    if detect_mode == "none":
        return False, None
    if not md_rows or not md_rows[0]:
        return False, None

    # column_headers mode
    if detect_mode == "column_headers":
        header_detection = opts.get("header_detection", True)
        if not header_detection:
            return False, None

        header = md_rows[0]
        def _is_empty_cell(v): return not (v and str(v).strip())
        if all(_is_empty_cell(cell) for cell in header):
            return False, None

        mapping = opts.get("mermaid_columns", {"from":"From","to":"To","label":"Label","group":None,"note":None})
        colmap = _v14_resolve_columns_by_name(header, mapping)
        if colmap:
            return True, colmap
        else:
            return False, None

    # heuristic mode
    if is_code_block(md_rows):
        return False, None

    # thresholds
    min_rows = int(opts.get("mermaid_heuristic_min_rows", 3))
    arrow_ratio = float(opts.get("mermaid_heuristic_arrow_ratio", 0.3))
    ratio_min = float(opts.get("mermaid_heuristic_len_median_ratio_min", 0.4))
    ratio_max = float(opts.get("mermaid_heuristic_len_median_ratio_max", 2.5))
    data = md_rows[1:] if len(md_rows)>1 else []
    if len(md_rows[0]) < 2:
        return False, None
    # count rows with first two columns non-empty (先頭2列: 列インデックス 0 と 1)
    def _val(x): return (str(x).strip() if x is not None else "")
    nonempty12 = [r for r in data if len(r)>=2 and _val(r[0]) and _val(r[1])]
    if len(nonempty12) < min_rows:
        return False, None
    # arrow presence ratio
    arrow_lines = 0
    pat = _re.compile(r"(->|→|⇒)")
    for r in data:
        if any(pat.search(str(c or "")) for c in r):
            arrow_lines += 1
    if len(data) == 0 or (arrow_lines / len(data)) < arrow_ratio:
        return False, None
    # median length ratio (空列削除前の元の列インデックス 0 と 1 を使用)
    import statistics as _stats
    def _len_nfkc(s): return len(_unicodedata.normalize("NFKC", str(s or "")))
    l0 = [_len_nfkc(r[0]) for r in data if len(r) >= 1]
    l1 = [_len_nfkc(r[1]) for r in data if len(r) >= 2]
    if not l0 or not l1:
        return False, None
    med0 = _stats.median(l0)
    med1 = _stats.median(l1) if any(l1) else 1
    ratio = med0 / max(med1, 1)  # ゼロ割防止
    if ratio < ratio_min or ratio > ratio_max:
        return False, None
    # heuristic assumes columns: 0,1,(2 label),(3 group),(4 note)
    colmap = {"from":0, "to":1, "label":2 if len(md_rows[0])>2 else None, "group":3 if len(md_rows[0])>3 else None, "note":4 if len(md_rows[0])>4 else None}
    return True, colmap

def build_mermaid(md_rows, opts, colmap):
    """Build mermaid flowchart codeblock from md_rows and colmap indices (or name indices)."""
    diagram_type = opts.get("mermaid_diagram_type", "flowchart")
    direction = opts.get("mermaid_direction","TD")
    keep_table = bool(opts.get("mermaid_keep_source_table", True))
    node_id_policy = opts.get("mermaid_node_id_policy", "auto")
    group_behavior = opts.get("mermaid_group_column_behavior", "subgraph")

    # Prepare sets
    nodes = {}
    edges = set()
    groups = {}
    header = md_rows[0]
    data_rows = md_rows[1:] if len(md_rows)>1 else []

    def node_id(name):
        if node_id_policy == "explicit":
            # Future extension: use explicit ID column if available
            # For now, fall back to auto
            pass
        # auto policy: sanitize display name
        nid = _v14_sanitize_node_id(str(name))
        i = 2
        base = nid
        while nid in nodes.values():
            nid = f"{base}_{i}"; i += 1
        return nid

    def get_val(row, idx):
        try:
            if idx is None: return ""
            return row[idx] if idx < len(row) else ""
        except Exception:
            return ""

    # Build nodes/edges
    for row in data_rows:
        f = get_val(row, colmap.get("from")) if isinstance(colmap.get("from"), int) else get_val(row, colmap.get("from"))
        t = get_val(row, colmap.get("to"))
        if not str(f).strip() or not str(t).strip():
            continue
        fid = nodes.get(f) or node_id(f); nodes[f]=fid
        tid = nodes.get(t) or node_id(t); nodes[t]=tid
        label = get_val(row, colmap.get("label"))
        key = (fid, tid, str(label or "").strip())
        if opts.get("mermaid_dedupe_edges", True):
            if key in edges:
                continue
        edges.add(key)
        g = get_val(row, colmap.get("group"))
        if g and group_behavior == "subgraph":
            groups.setdefault(g, set()).update([fid, tid])
        # group_behavior="ignore" の場合は groups に追加しない

    # Emit code
    lines = []
    lines.append("```mermaid")

    # diagram_type handling (currently only flowchart is supported)
    if diagram_type == "flowchart":
        lines.append(f"flowchart {direction}")
    elif diagram_type in ("sequence", "state"):
        # Future extension: sequenceDiagram / stateDiagram support
        # For now, fall back to flowchart
        warn(f"mermaid_diagram_type='{diagram_type}' is not yet implemented, using flowchart")
        lines.append(f"flowchart {direction}")
    else:
        lines.append(f"flowchart {direction}")

    # subgraphs (only if group_behavior="subgraph")
    if group_behavior == "subgraph":
        for gname in sorted(groups.keys()):
            lines.append(f"  subgraph {gname}")
            for nid in sorted(groups[gname]):
                # find display name
                disp = None
                for k,v in nodes.items():
                    if v==nid: disp = k; break
                lines.append(f'    {nid}["{str(disp)}"]')
            lines.append("  end")

    # nodes not in any group
    grouped = set().union(*groups.values()) if groups else set()
    for k, nid in sorted(nodes.items(), key=lambda kv: str(kv[0])):
        if nid in grouped:
            continue
        lines.append(f'  {nid}["{str(k)}"]')

    # edges
    for (a,b,lab) in sorted(edges):
        if lab:
            lines.append(f'  {a} -->|{lab}| {b}')
        else:
            lines.append(f'  {a} --> {b}')
    lines.append("```")
    return "\n".join(lines)
