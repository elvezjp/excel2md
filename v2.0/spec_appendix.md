# エクセル→マークダウン変換 仕様書 付録（v2.0）

> 本文書は [spec.md](spec.md) の付録です。

---

# 付録A：実装ガイド（非規範・参考）

> 本付録はアルゴリズムの**具体例**を示す「実装ガイド」です。仕様本体の挙動要件に従い、実装は本付録に厳密に従う必要はありません（同等の結果を満たす代替実装可）。

## A.1 「最大長方形分解」の具体例（§5.2 関連）

### A.1.1 目的
非矩形（L字/凹型）を**複数の長方形**に分割し、重複/包含を整理してテーブル領域を得る。

### A.1.2 推奨アルゴリズム（ヒストグラム法＋彫り抜き）
- 入力：印刷領域内の **非空=1 / 空=0** の2値グリッド（行×列）
- ステップ：
  1. 各行について、**上方向に連続する1の高さ**を積み上げたヒストグラム `H[c]` を作成
  2. 各行の `H` に対し、**スタック**で「ヒストグラム中の最大長方形」を列全体で列挙（O(C)）
  3. 得られた候補長方形を**面積降順**で並べ、グリッドから**確定長方形を彫り抜き**（1→0 に塗り替え）
  4. 1〜3 をグリッドが空になるまで繰り返し（O(R*C) 期待）
- 複雑度：O(R*C)。大域的にほぼ最小個数の長方形で被覆される実績良好。

### A.1.3 擬似コード（行ごと最大長方形の列挙）
```
for r in rows:
  for c in cols:
    H[c] = (H[c] + 1) if grid[r][c] == 1 else 0
  stack = []  # pair(index, height)
  for c in [0..C]:  # 番兵として c=C で高さ0を挿入
    h = H[c] if c < C else 0
    last = c
    while stack and stack[-1].height > h:
      i, hh = stack.pop()
      width = c - i
      emit_rectangle(top=r-hh+1, bottom=r, left=i, right=c-1)
      last = i
    if not stack or stack[-1].height < h:
      stack.push((last, h))
```
- `emit_rectangle` は候補集合に追加（後段で彫り抜き・重複解消）

### A.1.4 代替（貪欲拡張）
- 非空セル集合から BFS で連結成分を抽出し、各成分について
  - 左上から右下へ**最大拡張**可能な長方形を貪欲に作り、確定→除去→再探索
- 実装が容易だが、長方形個数が増える可能性あり（性能/品質は要トレードオフ評価）。

---

## A.2 「幾何学的和集合」の計算（§5.3 関連）

### A.2.1 問題設定
離散の矩形範囲（A1:C10, D3:F12 など）の**重複を除去**して一体として扱う。

### A.2.2 推奨アルゴリズム（ラインスイープ・行方向）
1. すべての矩形を**行イベント**に分解：`(row_start, +rect)` と `(row_end+1, -rect)`
2. 行を昇順に走査し、アクティブ矩形の**列区間**を集約（1D 区間和集合）
3. 連続行で列区間集合が同一なら**垂直結合**し、最大長方形として出力
- 列方向の和集合は**区間マージ**（開始位置でソート→貪欲マージ）で O(N log N)

### A.2.3 代替（グリッド塗りつぶし）
- 直接グリッドに 1 を塗り、**最大長方形分解**（A.1）に委譲。単純明快だが大判ではメモリ/時間増。

---

## A.3 「数値的に解釈可能」の判定（§7.4 関連）

### A.3.1 正規化ルール
- 前後空白除去（§3.1 `strip_whitespace` が true のとき）
- **負数表現**：先頭/末尾の括弧で負値 `"(123)" -> "-123"` を許容
- **符号**：先頭の `+`/`-` を許容（複数不可）
- **桁区切り**：`,` または `，` を**3桁ごと**として許容（不整合があれば非数値）
- **小数点**：`.` または `．` を許容（複数不可）
- **指数**：`e`/`E` + 符号付き整数を許容（例: `1.2e-3`）
- **百分率**：末尾 `%` を許容（値は 0.01 倍に**しない**、判定目的のみ。変換は §5.6 に従う）
- **通貨**：先頭に `¥`/`$`/`€`/`£`/`₩` 等を許容（判定用）。実際の出力は `currency_symbol` に従う。

### A.3.2 推奨正規表現（判定用）
```
^ \s*
(?:\()?
(?P<sign>[+-])?
(?:(?P<currency>[¥$€£₩]))?
(?P<int>\d{1,3}(?:[,，]\d{3})*|\d+)
(?:[\.．](?P<frac>\d+))?
(?:[eE](?P<exp>[+-]?\d+))?
(?:\))?
\s*(?P<pct>%)?
\s*$
```
- マッチしたら「数値的に解釈可能」。右寄せ判定に使用。
- 変換時（`percent_format="numeric"`/`currency_symbol="strip"` 等）は、カンマ/記号を除去し実数へ。

### A.3.3 例
- `¥1,234.56` → 判定OK（数値扱い）。出力は `currency_symbol` 設定に従う。
- `(12,345)%` → 判定OK。`percent_format="numeric"` のとき `12345`。

---

## A.4 NFC 正規化後の不可視/制御文字の扱い（§5.12 関連）

### A.4.1 削除対象（既定：`markdown_escape_level="safe"`）
- **C0 制御**：U+0000–0008, U+000B–000C, U+000E–001F
- **DEL**：U+007F
- **C1 制御**：U+0080–009F
- **ソフトハイフン**：U+00AD
- **ゼロ幅**：U+200B（ZWSP）, U+200C（ZWNJ）, U+200D（ZWJ）, U+2060（WJ）
- **異方向マーク（任意）**：U+2066–U+2069（LRI/RLI/FSI/PDI）
  - *注：LRM/RLM（U+200E/U+200F）は日本語文脈では通常不要。必要に応じて保持可。*

### A.4.2 保持対象
- タブ U+0009、改行 U+000A、復帰 U+000D は**保持**（セル内改行は `<br>` に変換）。

### A.4.3 設定での緩和
- `markdown_escape_level="minimal"` のとき、**C0/C1 制御と DEL のみ**削除（ゼロ幅等は保持）。

---

## A.5 採番/順序の厳密定義（補足）
- **テーブル順序**：キー1=`top` 昇順、キー2=`left` 昇順、キー3=（面積）昇順。
- **脚注番号**：`footnote_scope="book"` のときブック通し、`"sheet"` のとき各シートごとに 1 始まり。

---

## A.6 参考テストケース（追加）
1. L字領域を A.1 のヒストグラム法で 2〜3 長方形へ分解できること
2. 交差する印刷範囲（重複）を A.2 で正規化し、二重出力がないこと
3. `¥(1,234.00)%` の右寄せ判定が **真** になること
4. ZWSP 混入セルが削除規則で除去されること（safe）

---

# 付録B：仕様補遺

以下は、仕様として成立しているが、実装時に解釈の余地があった項目についての公式明確化です。

## B.1 テーブル見出しの座標表記（§8 出力レイアウト関連）

### 問題
出力例 `### Table 1 (A1:C5)` は単一矩形を前提としているが、非矩形領域を「最大長方形分解」で複数の矩形に分割した場合の表記方法が未定義。

### 仕様追加（確定・修正）
- テーブル見出しには**座標範囲を表示しない**。
- 見出しは `### Table 1` の形式のみとする。
- 座標情報が必要な場合は、デバッグモードや別のログ出力で提供する（将来拡張）。

---

## B.2 `max_cells_per_table` のカウント対象（§5.10, §4.8 関連）

### 問題
「逐次カウント」とあるが、空セルを含むか否か明示されていない。

### 仕様追加（確定）
- **対象範囲内の全セル**をカウント対象とする（非空セルのみではない）。
- 理由：安全側の制限とし、巨大な空白領域による過負荷を防ぐため。
- 実装時：テーブル走査時に1セル処理ごとにカウンタを+1し、閾値到達時点で当該テーブルを打切り（WARN ログを記録）。

---

## B.3 正規表現の括弧対応（付録A.3.2 関連）

### 問題
A.3.2 の正規表現は `( ... )` の括弧対応を厳密に保証しない（片側括弧のみでマッチする可能性）。

### 仕様追加（確定）
- マッチ後の**後検証ステップ**で以下を行う：
  - 開き括弧 `'('` が含まれる場合、閉じ括弧 `')'` も含まれていなければ不一致とみなす。
  - 両方ある場合は `()` の位置関係が逆転していないこと（`(` が `)` より前）。
- 例外：空白や通貨記号を除いた実質数値部分が括弧で囲まれている場合のみ有効。
  （例：`(1234)` → OK、`(¥1234%)` → OK、`¥(1234)` → NG）
- この検証は**右寄せ判定および数値変換**の直前で実施する。

---

## 付録B.4（参考） 実装オプションの設定追加提案

| 設定名 | 型 | 既定値 | 概要 |
|--------|----|--------|------|
| `table_range_mode` | string | `"enumerate_all"` | テーブル見出しの範囲表記方法 (`enumerate_all` or `representative`) |

---

以上は仕様として正式確定しています。

---

# 付録C：実装詳細ガイド — ディスパッチ／列名マッチング／フォールバック

## C.1 ディスパッチの導入方法（フォールバック手順を明確化・修正）
> 擬似コードを具体化し、フォールバック時の再評価開始点を明示。

```
def dispatch_table_output(md_rows, ctx):
    # 1) 最優先：コード
    if is_code_block(md_rows):        # 既存のコード判定を関数化
        return emit_code_block(md_rows, ctx)

    # 2) Mermaid
    try:
        if ctx.conf.mermaid_enabled:
            if ctx.conf.mermaid_detect_mode == "shapes":
                # ⑦''：シェイプ（図形）からの検出
                code = extract_shapes_to_mermaid(ctx.xlsx_path, ctx.sheet_name, ctx.grid)
                if code:
                    return "mermaid", code  # シェイプ検出の場合は元テーブルは出力しない
            elif ctx.conf.mermaid_detect_mode in ("column_headers", "heuristic"):
                # ⑦'：フローテーブル判定
                if is_flow_table(md_rows, ctx):
                    code = build_mermaid(md_rows, ctx)  # ⑧'〜⑨'
                    if ctx.conf.mermaid_keep_source_table:
                        table_md = emit_normal_table(
                            md_rows,
                            ctx,
                            header_detection=(ctx.conf.header_detection == "first_row")
                        )
                        return "mermaid", code + "\n" + table_md
                    return "mermaid", code
    except Exception as e:
        log_error("mermaid generation failed", e)
        # 以降はフォールバックで継続

    # フォールバックの開始点（常に優先度3から）
    # dispatch_skip_code_and_mermaid_on_fallback が true/false に関わらず、
    # フォールバックは **単一行（3）から**再評価を開始する。

    # 3) 単一行
    if is_single_line_text(md_rows, ctx):
        return emit_single_line(md_rows, ctx)

    # 4) ネスト
    if is_nested_text(md_rows, ctx):
        return emit_nested(md_rows, ctx)

    # 5) 通常テーブル
    return emit_normal_table(
        md_rows,
        ctx,
        header_detection=(ctx.conf.header_detection == "first_row")
    )
```

## C.2 列名マッチングのアルゴリズム
```
def normalize_header_name(s: str) -> str:
    s_nfkc = unicodedata.normalize("NFKC", s or "")
    s_fold = s_nfkc.casefold()
    s_strip = " ".join(s_fold.split())  # 前後空白除去 + 連続空白の1化
    return s_strip

def resolve_columns_by_name(header_row, mapping):  # mapping: {"from": "From", "to": "To", ...}
    norm_map = {k: normalize_header_name(v) for k, v in mapping.items() if v}
    index = {}
    for i, name in enumerate(header_row):
        key = normalize_header_name(name)
        for want_k, want in norm_map.items():
            if key == want and want_k not in index:
                index[want_k] = i
    # 必須: from/to の存在を確認
    if "from" not in index or "to" not in index:
        warn("required columns not found: from/to")
        return None
    # 同一列名が複数ある場合（最初の一致を採用）
    # - 以降の同名列は無視し、WARN を記録する（実装例：warn("duplicate header name: 'From' -> first match used")）
    # - より厳密な運用が必要な場合は将来設定化（例：first/last/error）
    return index  # 任意: label/group/note は存在すれば含まれる
```

## C.2.5 シェイプ（図形）からのMermaid抽出
> `excel_flow_to_ir.py` を参考にした実装ガイド。`dispatch_table_output` から呼び出される。

**実装上の注意**
- シェイプ検出はテーブル処理とは独立して実行される。
- ExcelファイルをZIPとして開き、`xl/drawings/*.xml` を読み取る必要がある。
- シートのセルグリッド（`{(row, col): value}`）を事前に構築しておく必要がある。
- エッジが不足している場合の推論（`infer_edges`）は、位置情報（bbox）に基づいて実施する。

**擬似コード**（`excel_flow_to_ir.py` を参考）

```
def extract_shapes_to_mermaid(xlsx_path: str, ws, grid: Dict[Tuple[int,int], str]) -> Optional[str]:
    """DrawingMLからシェイプを抽出し、Mermaidコードを生成

    Args:
        xlsx_path: Excelファイルのパス
        ws: Worksheetオブジェクト（openpyxl）
        grid: セルグリッド {(row, col): value} (0-based)
    """
    import zipfile, xml.etree.ElementTree as ET

    NS = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    }

    z = zipfile.ZipFile(xlsx_path)
    # シートに対応するDrawingMLファイルを特定
    # 1. ワークシートオブジェクトからシート名を取得
    sheet_name = ws.title

    # 2. workbook.xmlからシートIDを取得
    sheet_id = None
    wb_xml = z.read('xl/workbook.xml').decode('utf-8')
    wb_root = ET.fromstring(wb_xml)
    ns_wb = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
             'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    sheets = wb_root.findall('.//main:sheet', ns_wb)
    for sheet in sheets:
        if sheet.get('name') == sheet_name:
            r_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if r_id:
                # workbook.xml.relsからシートインデックスを取得
                wb_rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
                wb_rels_root = ET.fromstring(wb_rels_xml)
                ns_rel = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                for rel in wb_rels_root.findall('.//r:Relationship', ns_rel):
                    if rel.get('Id') == r_id:
                        target = rel.get('Target')
                        # ターゲットからシート番号を抽出（例: "worksheets/sheet1.xml" -> "1"）
                        if 'sheet' in target:
                            sheet_id = target.split('sheet')[1].split('.')[0]
                        break
            break

    # 3. シートの関係ファイルからDrawingMLファイルのパスを取得
    drawing_path = None
    if sheet_id:
        sheet_rel_path = f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels"
        if sheet_rel_path in z.namelist():
            rels_xml = z.read(sheet_rel_path)
            rels_root = ET.fromstring(rels_xml)
            ns_rel = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            for rel in rels_root.findall('.//r:Relationship', ns_rel):
                if rel.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing':
                    target = rel.get('Target')
                    # 相対パスを解決（xl/worksheets/基準）
                    if target.startswith('../'):
                        drawing_path = target.replace('../', 'xl/')
                    else:
                        drawing_path = f"xl/worksheets/{target}"
                    break

    if not drawing_path or drawing_path not in z.namelist():
        return None  # 対応するDrawingMLファイルが存在しない

    # 4. 対応するDrawingMLファイルを解析
    xml = z.read(drawing_path)
    root = ET.fromstring(xml)

    nodes = []
    edges = []
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
        for txbody in sp.findall(".//xdr:txBody", NS):
            for p in txbody.findall(".//a:p", NS):
                runs = []
                for r in p.findall(".//a:r", NS):
                    t = r.find("a:t", NS)
                    if t is not None and t.text:
                        runs.append(t.text)
                if runs:
                    texts.append("".join(runs))
        # Fallback: try a:txBody (for compatibility)
        if not texts:
            for txbody in sp.findall(".//a:txBody", NS):
                for p in txbody.findall(".//a:p", NS):
                    runs = []
                    for r in p.findall(".//a:r", NS):
                        t = r.find("a:t", NS)
                        if t is not None and t.text:
                            runs.append(t.text)
                    if runs:
                        texts.append("".join(runs))
        return "\n".join([t.strip() for t in texts if t and t.strip()])

    def cell_bbox(anc):
        def int_or_none(e):
            return int(e.text) if e is not None and e.text and e.text.isdigit() else None
        fr_c = anc.find(".//xdr:from/xdr:col", NS)
        fr_r = anc.find(".//xdr:from/xdr:row", NS)
        to_c = anc.find(".//xdr:to/xdr:col", NS)
        to_r = anc.find(".//xdr:to/xdr:row", NS)
        if all(x is not None for x in (fr_c, fr_r, to_c, to_r)):
            return [int_or_none(fr_r), int_or_none(fr_c), int_or_none(to_r), int_or_none(to_c)]
        return None

    def text_from_bbox(bbox):
        if not bbox: return ''
        r1, c1, r2, c2 = bbox
        pad = 1
        texts = []
        for r in range(max(0, r1-pad), r2+pad+1):
            for c in range(max(0, c1-pad), c2+pad+1):
                v = grid.get((r, c))
                if v:
                    texts.append(v)
        if texts:
            seen = set(); uniq = [t for t in texts if t not in seen and not seen.add(t)]
            longest = max(uniq, key=len)
            return longest if len(longest) >= 4 else "\n".join(uniq[:4])
        return ''

    # アンカー単位で走査
    anchors = root.findall(".//xdr:twoCellAnchor", NS) + root.findall(".//xdr:oneCellAnchor", NS)

    for anc in anchors:
        sp = anc.find(".//xdr:sp", NS)
        cxn = anc.find(".//xdr:cxnSp", NS)
        bbox = cell_bbox(anc)

        if sp is not None:
            cNvPr = sp.find(".//xdr:cNvPr", NS)
            _id = get_id(cNvPr)
            name = cNvPr.get("name") if cNvPr is not None else None

            # テキスト取得（優先順位1: txBody、優先順位2: bbox近傍、優先順位3: 名前）
            text = get_text_from_txBody(sp)
            if not text:
                text = text_from_bbox(bbox)
            if not text:
                text = name or ""

            # 形状種類の判定
            prstGeom = sp.find(".//a:prstGeom", NS)
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
            st = cxn.find(".//xdr:stCxn", NS)
            ed = cxn.find(".//xdr:endCxn", NS)
            st_id = f"s{st.get('id')}" if st is not None and st.get("id") else None
            ed_id = f"s{ed.get('id')}" if ed is not None and ed.get("id") else None
            if st_id and ed_id:
                edges.append({"from": st_id, "to": ed_id})

    # エッジが不足している場合は推論
    if len(edges) < max(2, int(0.3*len(nodes))):
        edges = edges + infer_edges(nodes, edges)  # infer_edgesは実装が必要

    if not nodes:
        return None

    # Mermaidコード生成
    direction = opts.get("mermaid_direction", "TD")
    lines = ["```mermaid", f"flowchart {direction}"]
    node_map = {}

    def escape_mermaid_text(text):
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

    def format_node(nid, display, shape_type):
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

    node_id_policy = opts.get("mermaid_node_id_policy", "auto")
    for n in nodes:
        display = n["text"] or n["name"] or n["id"]
        shape_type = n.get("type", "process")

        # ノードID生成
        if node_id_policy == "shape_id":
            # Excel描画IDをそのまま使用（s{id}形式）
            nid = n["id"]  # 既に "s{id}" 形式で格納されている
        elif node_id_policy == "auto":
            # 表示名をサニタイズしてID化
            nid = sanitize_node_id(str(display))
            # 重複処理（省略）
        else:  # explicit（将来拡張）
            nid = n["id"]

        node_map[n["id"]] = nid
        lines.append(format_node(nid, display, shape_type))

    for e in edges:
        from_id = node_map.get(e["from"])
        to_id = node_map.get(e["to"])
        if from_id and to_id:
            # 推論されたエッジかどうかで形式を変更（任意）
            inferred = e.get("inferred", False)
            if inferred:
                lines.append(f'  {from_id} -.->|inferred| {to_id}')
            else:
                lines.append(f'  {from_id} --> {to_id}')

    lines.append("```")
    return "\n".join(lines)
```

## C.3 heuristic の計算方法
- **矢印比**：`arrow_ratio = (矢印を含む行数) / (データ行数)`（矢印は `->|→|⇒` を正規表現で検出）
- **列長中央値比**：
  - データ行について、`len(NFKC(col0))` と `len(NFKC(col1))` の配列を作成
  - `median0 = median(lengths0)`, `median1 = median(lengths1)` を算出
  - `ratio = median0 / max(median1, 1)`（ゼロ割防止）
  - **対象列**は空列削除前の**元の列インデックス 0 と 1**を用いる（空セルは長さ 0 として扱う）。

## C.4 フォールバックの再評価
- 例外または不成立時、`dispatch_skip_code_and_mermaid_on_fallback=true` の場合：
  - **コード/Mermaidの再評価をスキップ**し、**単一行（3）から**の分岐を実行。
  - ログに `fallback_from="mermaid"` として記録（任意）。

## C.5 `header_detection=none` と元テーブル出力
- `mermaid_keep_source_table=true` の場合でも、元テーブルは `make_markdown_table` を使用し、
  ヘッダ行は生成しない（先頭行をデータ行として出力）。
  - 実装上は `make_markdown_table(md_rows, header_detection=False, ...)` のように、
    **`header_detection=False` を明示**して呼び出す（既存の 1365 行目相当の呼び出しと整合）。
  - `mermaid_keep_source_table=false` の場合はテーブルを出力しないため、この指定は不要。

## C.6 コード判定ロジックの抽出（互換性注意）
- 既存の `format_table_as_text_or_nested` 内の**コード判定ロジック（1090〜1127行目相当）**を
  `is_code_block(md_rows)` として切り出し、**⑦a の先頭で呼び出す**。
  - 判定式・正規表現・トリミング等は**既存実装をそのまま踏襲**し、判定の互換性を保つ。

---

# 付録D：単体テスト仕様

## D.1 概要

本付録は、`v2.0/excel_to_md.py` の単体テスト仕様を定義する。テストはpytestを使用し、openpyxlのモックまたはメモリ内Workbookを活用して、実際のExcelファイルなしでテスト可能とする。

### D.1.1 テストファイル構成

```
v2.0/
├── excel_to_md.py
├── verify_csv_markdown.py
├── spec.md
└── tests/
    ├── __init__.py
    ├── conftest.py              # pytest fixtures
    ├── test_cell_utils.py       # セル処理関連
    ├── test_table_detection.py  # テーブル検出関連
    ├── test_markdown_output.py  # Markdown出力関連
    ├── test_csv_markdown.py     # CSVマークダウン出力関連
    ├── test_mermaid.py          # Mermaid変換関連
    ├── test_hyperlink.py        # ハイパーリンク処理関連
    └── test_integration.py      # 統合テスト
```

### D.1.2 テストの原則

1. **モック優先**: openpyxlのWorkbook/Worksheetオブジェクトをメモリ内で生成
2. **独立性**: 各テストは他のテストに依存しない
3. **再現性**: ファイルシステムや外部リソースに依存しない
4. **網羅性**: 正常系・異常系・境界値を網羅

---

## D.2 テスト対象関数と優先度

### D.2.1 優先度A（コア機能）

| 関数名 | 説明 | テストファイル |
|--------|------|----------------|
| `cell_is_empty(cell, opts)` | セルの空判定 | test_cell_utils.py |
| `cell_display_value(cell, opts)` | セル表示値の取得 | test_cell_utils.py |
| `no_fill(cell, policy)` | セル背景色の判定 | test_cell_utils.py |
| `numeric_like(s)` | 数値判定 | test_cell_utils.py |
| `normalize_numeric_text(s, opts)` | 数値テキスト正規化 | test_cell_utils.py |
| `md_escape(text, level)` | Markdownエスケープ | test_cell_utils.py |
| `build_nonempty_grid(ws, area, ...)` | 非空セルグリッド構築 | test_table_detection.py |
| `grid_to_tables(ws, area, ...)` | テーブル検出・分割 | test_table_detection.py |
| `make_markdown_table(md_rows, ...)` | Markdownテーブル生成 | test_markdown_output.py |
| `extract_table(ws, table, ...)` | テーブル抽出 | test_markdown_output.py |

### D.2.2 優先度B（出力機能）

| 関数名 | 説明 | テストファイル |
|--------|------|----------------|
| `write_csv_markdown(wb, csv_data_dict, ...)` | CSVマークダウン出力 | test_csv_markdown.py |
| `extract_print_area_for_csv(ws, area, ...)` | CSV用印刷領域抽出 | test_csv_markdown.py |
| `hyperlink_info(cell)` | ハイパーリンク情報取得 | test_hyperlink.py |
| `is_valid_url(target)` | URL有効性判定 | test_hyperlink.py |

### D.2.3 優先度C（Mermaid・特殊形式）

| 関数名 | 説明 | テストファイル |
|--------|------|----------------|
| `is_source_code(text)` | ソースコード判定 | test_mermaid.py |
| `detect_code_language(lines)` | プログラミング言語検出 | test_mermaid.py |
| `format_table_as_text_or_nested(...)` | 出力形式判定 | test_mermaid.py |
| `dispatch_table_output(...)` | 出力ディスパッチ | test_mermaid.py |

### D.2.4 優先度D（ヘルパー関数）

| 関数名 | 説明 | テストファイル |
|--------|------|----------------|
| `a1_from_rc(row, col)` | 座標→A1形式変換 | test_cell_utils.py |
| `remove_control_chars(s)` | 制御文字削除 | test_cell_utils.py |
| `is_whitespace_only(s)` | 空白のみ判定 | test_cell_utils.py |
| `get_print_areas(ws, mode)` | 印刷領域取得 | test_table_detection.py |
| `union_rects(rects)` | 矩形和集合 | test_table_detection.py |
| `carve_rectangles(grid)` | 最大長方形分解 | test_table_detection.py |

---

## D.3 テストケース詳細

### D.3.1 セル処理関連（test_cell_utils.py）

#### D.3.1.1 `cell_is_empty` のテストケース

| テストID | 入力 | 期待値 | 説明 |
|----------|------|--------|------|
| CE001 | value=None, no fill | True | 空セル |
| CE002 | value="", no fill | True | 空文字列 |
| CE003 | value="  ", no fill | True | 空白のみ |
| CE004 | value="text", no fill | False | テキストあり |
| CE005 | value=None, yellow fill | False | 空だが背景色あり |
| CE006 | value=None, white fill | True | 白背景は空扱い |
| CE007 | value=0, no fill | False | 数値0は空ではない |
| CE008 | value=0.0, no fill | False | 浮動小数点0は空ではない |

#### D.3.1.2 `numeric_like` のテストケース

| テストID | 入力 | 期待値 | 説明 |
|----------|------|--------|------|
| NL001 | "123" | True | 整数 |
| NL002 | "-123" | True | 負の整数 |
| NL003 | "1,234" | True | 桁区切り |
| NL004 | "1,234.56" | True | 桁区切り＋小数 |
| NL005 | "¥1,234" | True | 通貨記号 |
| NL006 | "(123)" | True | 負数表記（括弧） |
| NL007 | "12.5%" | True | パーセント |
| NL008 | "abc" | False | 文字列 |
| NL009 | "12abc" | False | 混在 |
| NL010 | "(123" | False | 括弧不整合 |
| NL011 | "1.2.3" | False | 複数小数点 |
| NL012 | "" | False | 空文字列 |

#### D.3.1.3 `md_escape` のテストケース

| テストID | 入力 | level | 期待値 | 説明 |
|----------|------|-------|--------|------|
| ME001 | "text" | safe | "text" | エスケープ不要 |
| ME002 | "a\|b" | safe | "a\\|b" | パイプエスケープ |
| ME003 | "line1\nline2" | safe | "line1<br>line2" | 改行変換 |
| ME004 | "*bold*" | safe | "\\*bold\\*" | アスタリスクエスケープ |
| ME005 | "[link]" | minimal | "[link]" | minimalでは保持 |
| ME006 | "`code`" | aggressive | "\\`code\\`" | バッククォートエスケープ |

### D.3.2 テーブル検出関連（test_table_detection.py）

#### D.3.2.1 `grid_to_tables` のテストケース

| テストID | グリッド構成 | 期待テーブル数 | 説明 |
|----------|--------------|----------------|------|
| GT001 | 3x3全て非空 | 1 | 単一連続テーブル |
| GT002 | 空行で2分割 | 2 | 空行によるテーブル分割 |
| GT003 | 空列で2分割 | 2 | 空列によるテーブル分割 |
| GT004 | L字型配置 | 1 | 連結コンポーネント |
| GT005 | 離散した2領域 | 2 | 独立したテーブル |
| GT006 | 全セル空 | 0 | テーブルなし |

#### D.3.2.2 `carve_rectangles` のテストケース

| テストID | グリッド形状 | 期待長方形数 | 説明 |
|----------|--------------|--------------|------|
| CR001 | 2x3矩形 | 1 | 単一長方形 |
| CR002 | L字型 | 2 | L字を2長方形に分解 |
| CR003 | T字型 | 2以上 | T字を複数長方形に分解 |
| CR004 | 凹型 | 2以上 | 凹型を複数長方形に分解 |

### D.3.3 CSVマークダウン出力関連（test_csv_markdown.py）

#### D.3.3.1 `extract_print_area_for_csv` のテストケース

| テストID | 入力 | 期待値 | 説明 |
|----------|------|--------|------|
| EP001 | 3x3データ領域 | 3行×3列のリスト | 基本抽出 |
| EP002 | 結合セル含む | top_left_onlyで空 | 結合セル処理 |
| EP003 | 改行含むセル | スペース置換 | セル内改行処理 |
| EP004 | ハイパーリンク含む | 平文形式 | ハイパーリンク処理 |

#### D.3.3.2 `write_csv_markdown` のテストケース（）

| テストID | オプション | 期待値 | 説明 |
|----------|------------|--------|------|
| WC001 | csv_include_description=True | 概要セクションあり | デフォルト動作 |
| WC002 | csv_include_description=False | 概要セクションなし |  |
| WC003 | csv_include_metadata=True | メタデータセクションあり | デフォルト動作 |
| WC004 | csv_include_metadata=False | メタデータセクションなし | メタデータ省略 |
| WC005 | mermaid_enabled=True, shapes | Mermaidブロックあり |  |
| WC006 | mermaid_enabled=True, column_headers | Mermaidなし（警告） | 非対応モード |

### D.3.4 ハイパーリンク関連（test_hyperlink.py）

#### D.3.4.1 `hyperlink_info` のテストケース

| テストID | リンク種別 | 期待値 | 説明 |
|----------|------------|--------|------|
| HI001 | 外部URL | target="https://..." | 外部リンク |
| HI002 | 内部リンク | location="Sheet2!A1" | シート内リンク |
| HI003 | リンクなし | None | ハイパーリンクなし |
| HI004 | mailto: | target="mailto:..." | メールリンク |

#### D.3.4.2 `is_valid_url` のテストケース

| テストID | 入力 | 期待値 | 説明 |
|----------|------|--------|------|
| VU001 | "https://example.com" | True | HTTPS URL |
| VU002 | "http://example.com" | True | HTTP URL |
| VU003 | "mailto:user@example.com" | True | mailto |
| VU004 | "file:///path" | True | file URL |
| VU005 | "./relative" | True | 相対パス |
| VU006 | "javascript:..." | False | 無効なプロトコル |
| VU007 | "" | False | 空文字列 |
| VU008 | None | False | None |

### D.3.5 Mermaid関連（test_mermaid.py）

#### D.3.5.1 `is_source_code` のテストケース

| テストID | 入力 | 期待値 | 説明 |
|----------|------|--------|------|
| SC001 | "public class Foo {" | True | Javaクラス定義 |
| SC002 | "def foo():" | True | Python関数 |
| SC003 | "@Override" | True | Javaアノテーション |
| SC004 | "function test() {}" | True | JavaScript関数 |
| SC005 | "Hello World" | False | 通常テキスト |
| SC006 | "123.45" | False | 数値 |
| SC007 | "" | False | 空文字列 |

#### D.3.5.2 `detect_code_language` のテストケース

| テストID | 入力 | 期待値 | 説明 |
|----------|------|--------|------|
| DL001 | ["public class", "import java"] | "java" | Java検出 |
| DL002 | ["def foo():", "import os"] | "python" | Python検出 |
| DL003 | ["const x = 1", "function()"] | "javascript" | JS検出 |
| DL004 | ["普通のテキスト"] | "" | 言語不明 |

---

## D.4 フィクスチャ定義（conftest.py）

### D.4.1 Workbook/Worksheet生成ヘルパー

```python
@pytest.fixture
def empty_workbook():
    """空のWorkbookを生成"""
    return openpyxl.Workbook()

@pytest.fixture
def simple_worksheet(empty_workbook):
    """基本的なデータを持つWorksheetを生成"""
    ws = empty_workbook.active
    ws['A1'] = 'Header1'
    ws['B1'] = 'Header2'
    ws['A2'] = 'Data1'
    ws['B2'] = 'Data2'
    return ws

@pytest.fixture
def default_opts():
    """デフォルトのオプション辞書を生成"""
    return {
        "strip_whitespace": True,
        "escape_pipes": True,
        "merge_policy": "top_left_only",
        "hyperlink_mode": "inline_plain",
        "markdown_escape_level": "safe",
        "readonly_fill_policy": "assume_no_fill",
        "csv_include_description": True,
        "csv_include_metadata": True,
        # ... その他のデフォルト値
    }
```

### D.4.2 モックセル生成ヘルパー

```python
def create_mock_cell(value=None, fill_color=None, hyperlink=None, is_date=False):
    """テスト用のモックセルを生成"""
    # MagicMockまたはopenpyxlのCellを使用
    ...
```

---

## D.5 テスト実行方法

### D.5.1 全テスト実行

```bash
cd v2.0
pytest tests/ -v
```

### D.5.2 特定カテゴリのテスト実行

```bash
# セル処理関連のみ
pytest tests/test_cell_utils.py -v

# カバレッジ付き
pytest tests/ --cov=. --cov-report=html
```

### D.5.3 マーカーによるテスト分類

```python
# 遅いテストをスキップ
pytest tests/ -v -m "not slow"

# 統合テストのみ
pytest tests/ -v -m "integration"
```

---

## D.6 カバレッジ目標

| カテゴリ | 目標カバレッジ |
|----------|----------------|
| コア関数（優先度A） | 90%以上 |
| 出力機能（優先度B） | 80%以上 |
| Mermaid関連（優先度C） | 70%以上 |
| ヘルパー関数（優先度D） | 60%以上 |
| 全体 | 80%以上 |

---
