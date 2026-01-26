# excel_to_md.py モジュール化計画（v2.0）

> **✅ 完了**: この計画は PR#12 で対応完了しました。

## 目的
- 3,230行の単一ファイル（`v1.8/excel_to_md.py`）を分割し、保守性を向上させる。
- 状態共有が必要な処理はクラス化し、責務を明確化する。
- CLIの挙動と既存テストの互換性を維持する。
- 段階的に移行できる構成にする。

## 移行元
- **ソースファイル**: `v1.8/excel_to_md.py`（3,230行、55関数）
- **テストスイート**: `v1.8/tests/`（39以上のテストクラス、3,408行）
- **仕様書**: `v1.8/spec.md`

## 目標ディレクトリ構成（案）
- `v2.0/excel2md/__init__.py`
- `v2.0/excel2md/cli.py`
- `v2.0/excel2md/config.py`
- `v2.0/excel2md/output.py` （※ Python標準の `logging` との衝突回避）
- `v2.0/excel2md/workbook_loader.py`
- `v2.0/excel2md/cell_utils.py`
- `v2.0/excel2md/table_detection.py`
- `v2.0/excel2md/table_extraction.py`
- `v2.0/excel2md/table_formatting.py`
- `v2.0/excel2md/mermaid_generator.py`
- `v2.0/excel2md/image_extraction.py`
- `v2.0/excel2md/csv_export.py`
- `v2.0/excel2md/runner.py`
- `v2.0/excel_to_md.py` （shim: 後方互換用エントリポイント）

### 定数・VERSION の配置
- **cell_utils.py**: `MD_ESCAPE_RE`, `NUMERIC_PATTERN`, `WHITESPACE_CHARS`, `CONTROL_REMOVE_RANGES`（エスケープ・数値・制御文字関連）
- **image_extraction.py**: `_DRAWINGML_NS`（DrawingML の XML 名前空間）。`mermaid_generator._v14_extract_shapes_to_mermaid` からは `image_extraction` を import して参照する（または共通の `xl_drawing.py` に分離する場合は、両者から import）。
- **excel2md/__init__.py**: `__version__`（Single Source of Truth として一元管理。他モジュールからは `from . import __version__ as VERSION` で参照）

## クラス構成（案）

以下は責務のグルーピング案。**第一段階では関数のまま配置し、クラス化は「9) 安定化と整理」以降で検討する。** まずは分割とテスト継続を優先する。

- `ExcelToMarkdownRunner`
  - 設定と実行フローを保持し、各サービスを協調させる。
- `WorkbookLoader`
  - 安全なワークブック読み込みと印刷範囲の取得を担当。
- `CellFormatter`
  - 値の整形、日付判定、エスケープ、URL判定、数値処理（`numeric_like`, `normalize_numeric_text`）、罫線判定（`has_border`）を担当。
- `TableDetector`
  - グリッド生成とテーブル矩形検出を担当。
- `TableExtractor`
  - テーブル抽出、タイトル・脚注の抽出を担当。
- `MarkdownTableRenderer`
  - Markdown表やテキスト出力の整形を担当。
- `MermaidGenerator`
  - フロー表や図形からMermaidを生成。
- `ImageExtractor`
  - 画像抽出とセル座標への割り当てを担当。
- `CsvExporter`
  - CSV Markdown 出力を担当。

## 依存関係の方針
- `cli.py` → `runner.py` → 各サービスクラスの順に依存。
- サービス層はCLIに依存しない。
- 共通ユーティリティは `cell_utils.py` / `config.py` に集約。

### 依存関係図
```
cli.py
  └── runner.py
        ├── config.py
        ├── output.py
        ├── workbook_loader.py ─────┬── cell_utils.py
        ├── table_detection.py ─────┤
        ├── table_extraction.py ────┼── table_detection.py
        ├── table_formatting.py ────┤
        ├── mermaid_generator.py ───┘
        ├── image_extraction.py
        └── csv_export.py ── cell_utils.py
```

- **（※）** `table_extraction.dispatch_table_output` が `build_code_block_from_rows` / `make_markdown_table` / `format_table_as_text_or_nested` / `build_mermaid` を呼ぶため、`table_extraction` → `table_formatting`, `table_extraction` → `mermaid_generator` の依存がある。
- `csv_export.extract_print_area_for_csv` の `merged_lookup`, `cell_to_image` は **runner が** `table_detection.build_merged_lookup` および image 抽出結果から用意して渡す。`csv_export` は `table_detection` / `image_extraction` に直接依存しない。

### 循環依存の回避
- `table_extraction` は `table_detection`, `table_formatting`, `mermaid_generator` に依存するが、**table_formatting / mermaid_generator は table_extraction に依存しない**（一方向のみ）。
- `table_detection` は他モジュールに依存せず、`table_formatting` は `table_extraction` に依存しない（`runner` が協調）。
- **`has_border` の配置**: `table_formatting.format_table_as_text_or_nested` と `table_extraction` の両方で使用されるため、循環依存を避けるために `cell_utils.py` に配置。

## 移行ステップ
1) **ベースラインの確保**
   - 現行テストを実行してベースラインを確認。
   - CLI向けの簡易スモークテストが無ければ追加。

2) **パッケージの導入**
   - `excel2md/` パッケージを作成。
   - まずは純粋なヘルパー（エスケープ等）を移動。

3) **画像抽出の分離**
   - `extract_images_from_xlsx_drawing`, `extract_images_from_sheet` を `image_extraction.py` へ。`ImageExtractor` クラスは安定化フェーズで検討。

4) **Mermaid生成の分離**
   - `_v14_*` と `build_mermaid` / `is_flow_table` を移動。
   - `MermaidGenerator` クラスは安定化フェーズで検討。

5) **テーブル検出・抽出の分離**
   - グリッド生成、矩形検出を `table_detection.py` へ。
   - 抽出、タイトル・脚注は `table_extraction.py` へ。

6) **Markdown描画の分離**
   - `make_markdown_table`, `format_table_as_text_or_nested`, `build_code_block_from_rows` と整形系を `table_formatting.py` へ。

7) **設定の集約**
   - `config.py` に dataclass で設定とデフォルトを定義。
   - 引数パースとの対応付けを整理。既存の `opts` dict および `conftest` の `default_opts` と `Config` の対応付けを行う。

8) **実行オーケストレーション**
   - `run` をそのまま `runner.py` に移動し、既存テストを通す。
   - `dispatch_table_output` は `table_extraction` にあり、runner は `table_extraction.dispatch_table_output` を import して呼び出す。`run` 内の各処理の import ・引数（`merged_lookup`, `cell_to_image` 等）は、移動時に該当モジュールから取得する形に切り替える。
   - `cli.py` は `runner.run` の薄いラッパーにする。

9) **安定化と整理**
   - importの循環がないか確認。
   - テストやドキュメントの参照先を更新。
   - クラス化の検討（任意。`CellFormatter`, `TableDetector` 等を関数群からまとめる場合はここで実施）。

## リスク対策
- 移行中は `excel_to_md.py` を shim として残し、
  新パッケージを呼び出す形式にする。
- 小さな単位で移動し、都度テストを実行。
- まずは関数の署名を保ち、挙動変更を避ける。

## テスト方針

### 既存テストの活用
- `v1.8/tests/` から `v2.0/tests/` へテストをコピーして使用。
- 既存テストを段階ごとに実行:
  - `v2.0/tests/test_table_detection.py`
  - `v2.0/tests/test_image_extraction.py`
  - `v2.0/tests/test_mermaid.py`
  - `v2.0/tests/test_markdown_output.py`
  - `v2.0/tests/test_cell_utils.py`
  - `v2.0/tests/test_csv_markdown.py`
  - `v2.0/tests/test_hyperlink.py`

### import更新方針
- **shim方式**: `v2.0/excel_to_md.py` を後方互換shimとして維持。
  - 再エクスポート一覧は `excel2md/__init__.py` の `__all__` に集約し、shim は `from excel2md import ...` で一括して再エクスポートする。これにより shim の行数を抑える。
  - テストファイルの `from excel_to_md import a1_from_rc, format_value, ...` は、shim 経由でそのまま利用可能（`a1_from_rc` は `workbook_loader` に配置するが、shim が再エクスポートする）。
- **移行後の推奨import**: 新規コードでは直接パッケージを参照。第一段階では関数のため、例は関数で記載。クラス化後は `CellFormatter` 等を import。
  ```python
  # 推奨（第一段階: 関数）
  from excel2md.cell_utils import format_value, md_escape
  from excel2md.runner import run

  # 後方互換（shimを経由）
  from excel_to_md import format_value, run
  ```

### テスト追加基準
- 変更で挙動が増える場合のみ追加テストを作成。
- 既存テストが通れば、リファクタリングは成功と見なす。

## 受け入れ条件
- `python v2.0/excel_to_md.py` が従来通り動作する。
- 既存テストがすべて通る。
- メインスクリプト（shim）が 500 行未満になる。
  - **根拠**: 現行3,230行のうち、約85%（2,730行相当）を各モジュールに分離。
  - shim には `build_argparser`, `main` および `excel2md/__init__.py` 経由の再エクスポートのみを残す。再エクスポート一覧を `__init__.py` に集約することで shim を薄く保つ（目標 &lt; 200 行）。
- 責務と依存関係が明確である。

## 関数配置表（参考）

| モジュール | 配置する関数 | 概算行数 |
|-----------|-------------|---------|
| `cell_utils.py` | `no_fill`, `excel_is_date`, `format_value`, `cell_display_value`, `cell_is_empty`, `hyperlink_info`, `is_valid_url`, `numeric_like`, `normalize_numeric_text`, `md_escape`, `remove_control_chars`, `is_whitespace_only`, `has_border`（※） | 320 |
| `table_detection.py` | `build_nonempty_grid`, `enumerate_histogram_rectangles`, `carve_rectangles`, `bfs_components`, `rectangles_for_component`, `union_rects`, `grid_to_tables`, `collect_hidden`, `build_merged_lookup` | 400 |
| `table_extraction.py` | `detect_table_title`, `extract_table`, `dispatch_table_output` | 280 |
| `table_formatting.py` | `choose_header_row_heuristic`, `detect_right_align`, `make_markdown_table`, `format_table_as_text_or_nested`, `is_source_code`, `detect_code_language`, `build_code_block_from_rows`（※2） | 380 |
| `mermaid_generator.py` | `_v14_*` 関数群, `is_flow_table`, `build_mermaid` | 590 |
| `image_extraction.py` | `extract_images_from_xlsx_drawing`, `extract_images_from_sheet`, XML名前空間定数 | 350 |
| `csv_export.py` | `coords_to_excel_range`, `write_csv_markdown`, `extract_print_area_for_csv`, `format_timestamp`, `sanitize_sheet_name` | 200 |
| `workbook_loader.py` | `load_workbook_safe`, `a1_from_rc`, `parse_dimension`, `get_print_areas` | 150 |
| `config.py` | `Config` dataclass, デフォルト値 | 80 |
| `output.py` | `warn`, `info` | 30 |
| `runner.py` | `run`, `ExcelToMarkdownRunner` | 350 |
| `cli.py` | `build_argparser`, `main` | 100 |
| **合計** | | **3,230** |
| `excel_to_md.py` (shim) | 再エクスポート、後方互換ラッパー | **< 200** |

**注釈**:
- （※）循環依存回避のため `cell_utils.py` に配置
- （※2）`is_code_block` を統合し、コード検出からブロック生成までを一括処理

---

## 動作テスト項目

v2.0/spec.md の動作テスト項目について実行し確認する。

### CLI基本動作確認

```bash
uv run python v2.0/excel_to_md.py --help
uv run python v2.0/excel_to_md.py nonexistent.xlsx
```

- [x] `--help` で全オプションの説明が表示される
- [x] 存在しないファイル指定時、エラーメッセージと終了コード2で終了する

### test_standard.xlsx を使用した確認

**基本変換**
```bash
uv run python v2.0/excel_to_md.py test_standard.xlsx
```
- [x] エラーなく `test_standard_csv.md` が生成される
- [x] 5シート分のCSVコードブロックが出力される
- [x] 概要セクションとメタデータセクションが含まれる

**出力形式オプション**
```bash
uv run python v2.0/excel_to_md.py test_standard.xlsx --no-csv-markdown-enabled -o out.md
uv run python v2.0/excel_to_md.py test_standard.xlsx --split-by-sheet
uv run python v2.0/excel_to_md.py test_standard.xlsx --csv-output-dir ./output
```
- [x] `--no-csv-markdown-enabled -o out.md` でGFMテーブル形式のMarkdownが出力される
- [x] `--split-by-sheet` でシートごとに個別ファイル（5ファイル）が生成される
- [x] `--csv-output-dir ./output` で指定ディレクトリに出力される

**Sheet1「基本テーブル」の確認**
- [x] ヘッダー行が正しく認識される
- [x] 数値列が右寄せで出力される

**Sheet2「結合セル」の確認**
- [x] 結合セルの値が左上セルに出力される
- [x] 結合範囲の他セルは空で出力される

**Sheet3「複数テーブル」の確認**
- [x] 空行を境界として複数テーブルが検出される
- [x] 印刷領域内のみが変換対象となる

**Sheet4「ハイパーリンク」の確認**
```bash
uv run python v2.0/excel_to_md.py test_standard.xlsx --hyperlink-mode inline
uv run python v2.0/excel_to_md.py test_standard.xlsx --hyperlink-mode inline_plain
uv run python v2.0/excel_to_md.py test_standard.xlsx --hyperlink-mode footnote
```
- [x] `inline` モードで `[テキスト](URL)` 形式で出力される
- [x] `inline_plain` モードで `テキスト (URL)` 形式で出力される
- [x] `footnote` モードで脚注形式 `テキスト[^1]` で出力される
- [x] 内部リンク（シート参照）が `→シート名!セル` 形式で出力される
- [x] mailtoリンクが正しく処理される

**Sheet5「画像」の確認**
- [x] `test_standard_images/` ディレクトリが作成される
- [x] PNG画像が `画像_img_1.png` として抽出される
- [x] JPEG画像が `画像_img_2.jpg` として抽出される
- [x] CSVに `![alt](test_standard_images/...)` 形式でリンクが出力される

```bash
uv run python v2.0/excel_to_md.py test_standard.xlsx --no-image-extraction
```
- [x] `--no-image-extraction` で画像が抽出されない

### test_mermaid.xlsx を使用した確認

**Sheet1「図形フロー」の確認**
```bash
uv run python v2.0/excel_to_md.py test_mermaid.xlsx --mermaid-enabled --mermaid-detect-mode shapes
uv run python v2.0/excel_to_md.py test_mermaid.xlsx --mermaid-enabled --mermaid-detect-mode shapes --mermaid-direction LR
```
- [x] フローチャート図形からMermaidコードが生成される
- [x] 矩形が `["テキスト"]`、ひし形が `{"テキスト"}` で出力される
- [x] コネクタが `-->` で接続される
- [x] `--mermaid-direction LR` で `flowchart LR` が出力される

**Sheet2「テーブルフロー」の確認**
```bash
uv run python v2.0/excel_to_md.py test_mermaid.xlsx --mermaid-enabled --mermaid-detect-mode column_headers
uv run python v2.0/excel_to_md.py test_mermaid.xlsx --mermaid-enabled --mermaid-detect-mode column_headers --no-mermaid-keep-source-table
```
- [x] From/To/Label列からMermaidコードが生成される
- [x] `--no-mermaid-keep-source-table` で元テーブルが出力されない

---

## 既知の問題（すべて修正済み）

> 以下の3つの問題は動作テストで発見され、修正済みです。

### 問題1: 脚注番号の重複 ✅ 修正済み
- **現象**: `--hyperlink-mode footnote` で全リンクの脚注番号が `[^1]` になる
- **期待**: 各リンクに連番 `[^1]`, `[^2]`, `[^3]`... が付与される
- **原因箇所**: `table_extraction.py` 248行目、258行目
- **原因**: 脚注番号の計算で `len(footnotes)` を使用しているが、`footnotes` は関数引数で関数内では変化しない
- **修正内容**: `len(footnotes)` → `len(note_refs)` に変更

### 問題2: CSV内の画像リンク ✅ 修正済み
- **現象**: CSVコードブロック内に画像リンク `![alt](path)` が含まれない
- **期待**: 画像が配置されたセルに `![alt](test_standard_images/...)` 形式でリンクが出力される
- **原因箇所**: `runner.py` 233-236行目
- **原因**: `used_range` はテキストがある範囲のみで、画像が配置されている列を含まない
- **修正内容**: `cell_to_image` のキーを参照し、画像がある位置を含むように `union_area` を拡張

### 問題3: Shapesモードのコネクタ表現 ✅ 修正済み
- **現象**: shapes検出モードでコネクタが `-.->|inferred|` （破線＋推論ラベル）で出力される
- **期待**: 実線矢印 `-->` で接続される
- **原因箇所**: `mermaid_generator.py` 251-252行目
- **原因**: コネクタの接続情報は `a:` 名前空間にあるが、コードは `xdr:` 名前空間で検索していた
- **修正内容**: `xdr:stCxn` → `a:stCxn`、`xdr:endCxn` → `a:endCxn` に変更
