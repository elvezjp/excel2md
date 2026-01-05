# excel2md

このプログラムは、**仕様書 v1.7**に完全準拠した Excel → Markdown 変換ツールです。
Excelブック（.xlsx/.xlsm）を読み取り、各シートの印刷領域を解析してMarkdown形式で自動生成します。
**v1.7では、CSVマークダウンモードの拡張機能（Mermaid出力対応、説明文除外オプション）を追加しました。**

---

## 📘 基本情報

- **最新バージョン:** v1.7
  - **ファイル名:** `v1.7/excel_to_md.py`
  - **仕様書:** `v1.7/spec.md`
- **旧バージョン:**
  - v1.6: `v1.6/excel_to_md.py` / `v1.6/spec.md`
  - v1.5: `v1.5/excel_to_md_v1_5.py` / `v1.5/spec_v1_5.md`
  - v1.4: `v1.4/excel_to_md_v1_4.py` / `v1.4/spec_v1_4.md`
  - v1.3: `v1.3/excel_to_md_v1_3.py` / `v1.3/spec_v1_3_1.md`
- **対応フォーマット:** `.xlsx` / `.xlsm`
- **出力:** UTF-8 通常マークダウン (`.md`) + CSVマークダウン (`_csv.md`)
  - **通常マークダウン:** シート内容を矩形分割し、Markdownテーブル形式で出力
  - **CSVマークダウン:** シート全体をCSVブロックで出力。説明文・検証用メタデータを併記
- **主要依存ライブラリ:** [openpyxl](https://pypi.org/project/openpyxl/)

---

## 🚀 使い方

### v1.7（最新版）

```bash
python v1.7/excel_to_md.py input.xlsx -o output.md
```

### 入出力形式

| 形式 | 必要なオプション | 説明 |
|-----|----------------|------|
| CSVマークダウン | `--csv-markdown-enabled` | シート全体をCSVブロックで出力。説明文・検証用メタデータを併記 |
| 通常マークダウンのみ | `--no-csv-markdown-enabled` | シート内容を矩形分割してMarkdownテーブルで出力 |
| 通常マークダウン + Mermaid | `--no-csv-markdown-enabled` `--mermaid-enabled` | フローチャートをMermaid形式で追加出力 |


**デフォルト動作（v1.7）:**
- CSVマークダウンを出力（`--no-csv-markdown-enabled` で通常マークダウンのみ）
- Mermaidフローチャート変換は無効（`--mermaid-enabled` で有効化、CSVマークダウンでも有効）
- ハイパーリンクは脚注形式で出力（`--hyperlink-mode` で変更可能）

### 出力ファイル分割

| オプション | 出力ファイル | 説明 |
|-----------|-------------|------|
| デフォルト | `output.md` | 全シートを1ファイルにまとめて出力 |
| `--split-by-sheet` | `output_{シート名}.md` | シートごとに個別ファイルを生成 |

### オプション一覧

| カテゴリ | オプション名 | デフォルト値 | 説明 |
|---------|-------------|-------------|------|
| **出力形式** | `-o`, `--output` | - | 出力ファイルパス |
| | `--split-by-sheet` | `false` | シートごとに個別ファイルを生成 |
| **CSVマークダウン** | `--csv-markdown-enabled` / `--no-csv-markdown-enabled` | `true` | CSVマークダウン出力の有効化/無効化 |
| | `--csv-output-dir` | （入力と同じ） | CSVマークダウンの出力先ディレクトリ |
| | `--csv-include-metadata` / `--no-csv-include-metadata` | `true` | 検証用メタデータを含めるか |
| | `--csv-include-description` / `--no-csv-include-description` | `true` | 概要セクション（説明文）を含めるか【v1.7】 |
| | `--csv-apply-merge-policy` / `--no-csv-apply-merge-policy` | `true` | CSV抽出時にmerge_policyを適用するか |
| | `--csv-normalize-values` / `--no-csv-normalize-values` | `true` | CSV値に数値正規化を適用するか |
| **Mermaid** | `--mermaid-enabled` | `false` | Mermaidフローチャート変換を有効化（CSVマークダウンでも有効）【v1.7】 |
| | `--mermaid-detect-mode` | `shapes` | 検出モード（`shapes` / `column_headers` / `heuristic`） |
| | `--mermaid-direction` | `TD` | フローチャートの方向（`TD` / `LR` / `BT` / `RL`） |
| | `--mermaid-keep-source-table` | `true` | 元のテーブルも出力するか |
| **ハイパーリンク** | `--hyperlink-mode` | `footnote` | 出力方法（`inline` / `inline_plain` / `footnote` / `both` / `text_only`） |
| | `--footnote-scope` | `book` | 脚注番号の採番スコープ（`book` / `sheet`） |
| **テーブル処理** | `--header-detection` | `first_row` | テーブル先頭行をヘッダとして扱う |
| | `--align-detection` | `numbers_right` | 数値列の右寄せ判定（80%ルール） |
| | `--no-print-area-mode` | `used_range` | 印刷領域未設定時の動作（`used_range` / `entire_sheet_range` / `skip_sheet`） |
| | `--max-cells-per-table` | `200000` | テーブル1つあたりの最大セル数 |
| | `--markdown-escape-level` | `safe` | Markdown記号のエスケープレベル（`safe` / `minimal` / `aggressive`） |

### 使用例

#### 基本的な使い方（CSVマークダウン自動出力）
```bash
python v1.7/excel_to_md.py input.xlsx -o output.md
# → output.md と input_csv.md が生成される
```

#### リンクを平文形式で出力
```bash
python v1.7/excel_to_md.py input.xlsx -o output.md --hyperlink-mode inline_plain
# → セル内のリンクを「表示テキスト (URL)」形式で出力
```

#### CSVマークダウンで表示テキストのみ出力
```bash
python v1.7/excel_to_md.py input.xlsx -o output.md --hyperlink-mode text_only
# → CSVマークダウンではリンク情報を含まず、表示テキストのみを出力
```

#### シートごとに個別ファイルを生成
```bash
python v1.7/excel_to_md.py input.xlsx -o output.md --split-by-sheet
# → output_Sheet1.md, output_Sheet2.md のように、各シートを個別のファイルに出力
```

#### Mermaid有効化 + CSVマークダウン出力
```bash
python v1.7/excel_to_md.py input.xlsx -o output.md --mermaid-enabled --mermaid-detect-mode shapes
# → フローチャートをMermaid形式で出力 + CSVマークダウンでもMermaid出力
```

#### CSVマークダウンの説明文を除外（トークン数削減）
```bash
python v1.7/excel_to_md.py input.xlsx -o output.md --no-csv-include-description
# → CSVマークダウンから概要セクションを除外し、トークン数を削減
```

#### ヘルプ表示
```bash
python v1.7/excel_to_md.py --help
```

---

## 🧪 出力例

### 通常マークダウン出力例

```markdown
# 変換結果: sample.xlsx

- 仕様バージョン: 1.7
- シート数: 2
- シート一覧: Sheet1, 集計

---

## Sheet1

### Table 1 (A1:C4)
| 品目 | 数量 | 備考 |
| --- | ---: | --- |
| りんご | 10 | [発注先](https://example.com)[^1] |
| みかん | 5 |  |

[^1]: https://example.com

## 集計

### Table 1 (A1:C3)
| 月 | 売上 | 前年比 |
| --- | ---: | --- |
| 1月 | 100 | 105% |
| 2月 | 120 | 98% |
```

### CSVマークダウン出力例

```markdown
# CSV出力: sample.xlsx

## 概要

### ファイル情報

- 元のExcelファイル名: sample.xlsx
- シート数: 2
- 生成日時: 2025-11-11 12:34:56

### このファイルについて

このCSVマークダウンファイルは、AIがExcelの内容を理解できるよう、各シートの印刷領域をCSV形式で出力したファイルです。各シートはマークダウン見出しで区切られ、CSVコードブロックで内容が記載されています。ファイル末尾に検証用メタデータセクションがあり、Excel原本との整合性を確認できます。

### CSV生成方法

- **出力対象範囲**: 各シートの印刷領域のみを出力（印刷領域外のセルは含まない）
- **印刷領域の統合**: 複数の印刷領域がある場合は外接矩形として統合
- **テーブル分割なし**: 通常マークダウン出力と異なり、空行・空列でテーブル分割せず印刷領域全体を1つのCSVとして出力
- **結合セルの処理**: `top_left_only` 設定に従って処理
- **数式の扱い**: `display` 設定に従って表示値・数式・両方のいずれかを出力
- **値の正規化**: `True` 設定に従って正規化処理を適用

### CSV形式の仕様

- **区切り文字**: `,`
- **引用符の使用**: `minimal` 設定に従い、RFC 4180準拠で処理
- **セル内改行**: 半角スペースに変換される（1レコード=1行を保証）
- **セル内特殊文字**: 区切り文字や引用符は引用符でエスケープされる
- **エンコーディング**: `utf-8`
- **空セルの表現**: 空文字列として出力（空行・空列も含めて出力される）
- **ハイパーリンク**: `--hyperlink-mode` 設定に従って出力（既定: `inline_plain` で `表示テキスト (URL)` 形式）
  - `inline`: Markdown形式（例: `[表示テキスト](URL)`）
  - `inline_plain`: 平文形式（例: `表示テキスト (URL)`）
  - `text_only`: 表示テキストのみ（リンク情報なし）
  - `footnote`/`both`: `inline_plain` にフォールバック（脚注非対応）

### 検証用メタデータについて

- このファイルの末尾に「検証用メタデータ」セクションが付記されています
- メタデータには各シートのExcel原本情報（範囲、行数、列数）とCSV出力結果の比較が含まれます
- 検証ステータス（OK/FAILED）により、CSVとExcel原本の整合性を確認できます
- 詳細情報はファイル末尾の「検証用メタデータ」セクションを参照してください

---

## Sheet1

```csv
品目,数量,備考
りんご,10,発注先
みかん,5,
```

## 集計

```csv
月,売上,前年比
1月,100,105%
2月,120,98%
```

---

## 検証用メタデータ

- **生成日時**: 2025-11-11 12:34:56
- **元Excelファイル**: sample.xlsx
- **CSVマークダウンファイル**: sample_csv.md
- **CSVモード**: csv_markdown
- **検証フェーズ**: 1 (Basic - Row/Column Count)

### シート: Sheet1

- **Excel範囲**: A1:C3
- **Excel行数**: 3
- **Excel列数**: 3
- **CSV行数**: 3
- **CSV列数**: 3
- **セル総数**: 9
- **ステータス**: OK

### シート: 集計

- **Excel範囲**: A1:C3
- **Excel行数**: 3
- **Excel列数**: 3
- **CSV行数**: 3
- **CSV列数**: 3
- **セル総数**: 9
- **ステータス**: OK

### 全体の検証結果

- **総セル数**: 18
- **検証ステータス**: OK
```

---

## ⚙️ 必要環境

- Python 3.9 以上
- インストール:
  ```bash
  pip install openpyxl
  ```
  または
  ```bash
  uv pip install openpyxl
  ```

---

## 🧭 ライセンスとバージョン管理

- **最新仕様バージョン:** v1.7
  - v1.7: CSVマークダウンモードの拡張機能（Mermaid出力対応、説明文除外オプション）
  - v1.6: ハイパーリンク平文出力モードとシート分割出力機能追加
  - v1.5: CSVマークダウン出力機能追加
  - v1.4: Mermaidフローチャート変換追加
  - v1.3: 基本実装（補遺 v1.3.1 を含む）
- **ファイルバージョン:**
  - v1.7: `v1.7/excel_to_md.py`
  - v1.6: `v1.6/excel_to_md.py`
  - v1.5: `v1.5/excel_to_md_v1_5.py`
  - v1.4: `v1.4/excel_to_md_v1_4.py`
  - v1.3: `v1.3/excel_to_md_v1_3.py`
- **ライセンス:** MIT（予定）

---

## 🧑‍💻 開発メモ

- 各バージョンでコード内に `VERSION = "1.X"` を明記し、仕様との同期を保証。
- v1.7では、v1.6の全機能 + CSVマークダウンモードの拡張機能（Mermaid出力対応、説明文除外オプション）を統合。
- バージョン間の互換性を維持し、旧バージョンも引き続き使用可能。

---

## 🧩 主な仕様対応ポイント

### v1.7 新機能

#### CSVマークダウンモードの拡張機能

**Mermaid出力対応（CSVマークダウン）**
- **`--mermaid-enabled`オプションがCSVマークダウンでも有効に**
- `mermaid_detect_mode="shapes"`の場合のみ対応（Excelの図形からフローチャート抽出）
- `column_headers`/`heuristic`モードはCSVマークダウンでは非対応（WARNログを出力してスキップ）
- 各シートのCSVブロック直後にMermaidコードブロックを出力

**概要セクション除外オプション（--csv-include-description）**
- **`--no-csv-include-description`オプションを追加**
- CSVマークダウンの概要セクション（説明文）を除外可能
- 複数ファイルを変換・結合する際のトークン数削減に対応
- デフォルトは`true`（従来通り概要セクションを出力）

### v1.6 新機能

#### ハイパーリンク平文出力モード（inline_plain）
- **`--hyperlink-mode inline_plain`オプションを追加**
- セル内のハイパーリンクを平文形式で出力: `表示テキスト (URL)`
- Markdown記法を使わずにリンク情報を明示的に表示
- 従来の`inline`（Markdown形式: `[表示テキスト](URL)`）、`footnote`（脚注形式）、`both`（両方）に加えて選択可能

#### シート分割出力機能（--split-by-sheet）
- **`--split-by-sheet`オプションを追加**
- 各シートを個別のMarkdownファイルとして出力
- ファイル名形式: `{出力ファイル名}_{シート名}.md`
- 各シートファイルには、シート名、仕様バージョン、元ファイル名を記載
- シートごとに独立した脚注番号を使用（`footnote_scope`が`book`の場合でも、各シートで独立）

### v1.5 新機能

#### CSVマークダウン出力機能（§3.2.2）
- **ファイル名形式**: `{元のExcelファイル名}_csv.md`
- **各シートの印刷領域をCSV形式のコードブロックとして記載**
- **ファイル冒頭に概要セクションを自動生成**:
  - ファイル情報（元のExcelファイル名、シート数、生成日時）
  - このファイルについて（CSVマークダウンの説明）
  - CSV生成方法（出力対象範囲、印刷領域の統合、結合セル処理など）
  - CSV形式の仕様（区切り文字、引用符、セル内改行、エンコーディングなど）
  - 検証用メタデータについて
- **ファイル末尾に検証用メタデータセクションを自動生成**（マークダウン形式）:
  - 各シートのExcel原本情報（範囲、行数、列数）とCSV出力結果の比較
  - 検証ステータス（OK/FAILED）
- **CSV形式の特徴**:
  - RFC 4180準拠
  - セル内改行は半角スペースに変換（1レコード=1行を保証）
  - ハイパーリンクは表示テキストのみ出力（リンク先URLは含まない）

### v1.4 新機能

#### Mermaidフローチャート変換（§⑦〜§⑨）
- **列名ベース検出**: `From`/`To`/`Label`列を検出してフローチャート化
- **ヒューリスティック検出**: テーブル構造から自動判定
- **シェイプ検出**: ExcelのDrawingML図形からフローチャートを抽出
- ノードID自動生成、重複エッジ除去、サブグラフ対応

### v1.3 基本機能

#### 最大長方形分解アルゴリズム（§5.2, 付録A.1）
- **ヒストグラム法＋彫り抜き法**を実装。
- 非矩形領域（L字・凹型）を複数矩形に分解し、上から順に処理。

#### 印刷領域と空セル判定（§4〜§6）
- 印刷領域未設定時は `used_range` を使用。
- 空セル定義：
  - 値が空または空白のみ（全角/半角スペース含む）
  - 塗りつぶしが **No Fill** の場合のみ空とみなす。

#### 結合セルと空判定（§5.4）
- 結合セル内のいずれかが非空なら全体を非空。
- `merge_policy=top_left_only`（既定）で左上セル値のみを出力。

#### Markdown 出力（§7）
- テーブルは `|` 区切り形式。
- ヘッダ自動検出: 先頭行をヘッダ扱い (`header_detection=first_row`)。
- 改行は `<br>` に変換。連続改行も保持。
- 数値列は右寄せ (`---:`) 自動判定（80%以上が数値なら）。

#### ハイパーリンク処理（§5.5）
- 外部URL, メールアドレス, ファイル参照, シート内リンクを検出。
- `footnote` モードでは脚注として出力。採番方式は `footnote-scope` に従う。

#### パフォーマンスと制限（§5.10）
- `openpyxl.load_workbook(read_only=True, data_only=True)` を使用。
- `max_cells_per_table` 超過で当該テーブルを安全に打切り。
- 巨大ファイル対策として逐次セル走査。

---

## 🔧 リリースノート

### 2025-12-25 v1.7 リリース（CSVマークダウンモード拡張機能）

#### 🎯 主な新機能

**CSVマークダウンでのMermaid出力対応**
- **`--mermaid-enabled`オプションがCSVマークダウンでも有効に**
- `mermaid_detect_mode="shapes"`の場合のみ対応（Excelの図形からフローチャート抽出）
- `column_headers`/`heuristic`モードはCSVマークダウンでは非対応（テーブル分割がないため）
- 従来は`--mermaid-enabled`を指定してもCSVマークダウンには反映されなかったが、v1.7で統一的な動作に改善

**概要セクション除外オプション**
- **`--no-csv-include-description`オプションを追加**
- CSVマークダウンの概要セクション（説明文）を除外可能
- 複数ファイルを変換・結合する際のトークン数削減に対応

#### 🔄 破壊的変更
なし（v1.6との後方互換性を維持）

#### 📝 使用例
```bash
# CSVマークダウンでもMermaid出力を有効化
python v1.7/excel_to_md.py input.xlsx -o output.md --mermaid-enabled

# CSVマークダウンの説明文を除外（トークン数削減）
python v1.7/excel_to_md.py input.xlsx -o output.md --no-csv-include-description

# 両方を組み合わせ
python v1.7/excel_to_md.py input.xlsx -o output.md --mermaid-enabled --no-csv-include-description
```

---

### 2025-11-18 v1.6 リリース（ハイパーリンク平文出力とシート分割機能追加）

#### 🎯 主な新機能

**ハイパーリンク平文出力モード（inline_plain）**
- **`--hyperlink-mode inline_plain`オプションを追加**
- セル内のハイパーリンクを平文形式で出力: `表示テキスト (URL)`
- 内部リンクの場合: `表示テキスト (→場所)`
- Markdown記法を使わずにリンク情報を明示的に表示

**シート分割出力機能（--split-by-sheet）**
- **`--split-by-sheet`オプションを追加**
- 各シートを個別のMarkdownファイルとして出力
- ファイル名形式: `{出力ファイル名}_{シート名}.md`
- 各シートファイルには独立した脚注番号を使用

#### 🔄 破壊的変更
なし（v1.5との後方互換性を維持）

#### 📝 使用例
```bash
# ハイパーリンク平文出力
python v1.6/excel_to_md.py input.xlsx -o output.md --hyperlink-mode inline_plain

# シート分割出力
python v1.6/excel_to_md.py input.xlsx -o output.md --split-by-sheet
```

---

### 2025-11-11 v1.5 リリース（CSVマークダウン出力機能追加）

#### 🎯 主な新機能

**CSVマークダウン出力機能（デフォルト有効）**
- **ファイル名形式**: `{basename}_csv.md`
- **各シートの印刷領域をCSVコードブロックとして記載**
  - 印刷領域を対象に出力
  - 結合セルの処理（`merge_policy`に従う）
  - 値の正規化（NFC正規化、空白処理、日付・数値フォーマット）
  - セル内改行を半角スペースに変換（1レコード=1行を保証）
  - ハイパーリンクは表示テキストのみ出力
- **概要セクションと検証用メタデータセクションを自動生成**
  - 概要: ファイル情報、CSV生成方法、CSV形式の仕様など
  - メタデータ: 各シートのExcel原本情報とCSV出力結果の比較、検証ステータス（OK/FAILED）

**バッチ処理対応**
- `batch_test.py` をv1.5対応に更新
- CSVマークダウン出力統計の表示機能を追加

#### 🔄 破壊的変更
なし（v1.4との後方互換性を維持）

#### 📝 使用例
```bash
# デフォルト: Markdown + CSVマークダウン
python v1.5/excel_to_md_v1_5.py input.xlsx -o output.md

# Mermaid有効化 + CSVマークダウン
python v1.5/excel_to_md_v1_5.py input.xlsx -o output.md --mermaid-enabled --mermaid-detect-mode shapes
```

---

### 2025-11-08 v1.4 リリース（Mermaidフローチャート変換追加）

- Mermaidフローチャート変換機能を追加
- 列名ベース検出、ヒューリスティック検出、シェイプ検出の3つの検出モードに対応
- ノードID自動生成、重複エッジ除去、サブグラフ対応

---

### 2025-11-08 v1.3 初回リリース

- 基本実装（印刷領域処理、最大長方形分解、Markdown出力など）

---

(c) 2025 Elvez

詳細な仕様については、各バージョンの仕様書を参照してください。
- v1.7: `v1.7/spec.md`
- v1.6: `v1.6/spec.md`
- v1.5: `v1.5/spec_v1_5.md`
