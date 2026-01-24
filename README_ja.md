# excel2md

[English](./README.md) | [日本語](./README_ja.md)

[![Elvez](https://img.shields.io/badge/Elvez-Product-3F61A7?style=flat-square)](https://elvez.co.jp/)
[![IXV Ecosystem](https://img.shields.io/badge/IXV-Ecosystem-3F61A7?style=flat-square)](https://elvez.co.jp/ixv/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.9+-blue?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Stars](https://img.shields.io/github/stars/elvezjp/excel2md?style=social)](https://github.com/elvezjp/excel2md/stargazers)

Excel → Markdown 変換ツール。Excelブック（.xlsx/.xlsm）を読み取り、Markdown形式で自動生成します。

## 特徴

- **スマートテーブル検出**: Excel印刷領域を自動検出してMarkdownテーブルに変換
- **CSVマークダウン出力**: シート全体をCSV形式で出力（検証用メタデータ付き）
- **画像抽出** (v1.8): Excelファイル内の画像を外部ファイルとして抽出し、Markdownリンク形式で出力
- **Mermaidフローチャート**: Excel図形やテーブルからMermaid図を生成
- **ハイパーリンク対応**: 複数の出力モード（インライン、脚注、平文）
- **シート分割出力**: シートごとに個別ファイルを生成可能
- **カスタマイズ可能**: 書式、配置、データ処理の詳細設定が可能

## ユースケース

- **ドキュメント生成**: Excel仕様書をMarkdownに変換
- **AI/LLM処理**: トークン効率に最適化されたCSVマークダウン形式
- **フローチャート抽出**: Excel図形から図を抽出
- **データ移行**: ExcelデータをポータブルなMarkdown形式にエクスポート
- **バージョン管理**: Excelの変更をテキストベース形式で追跡

## ドキュメント

- [CHANGELOG.md](CHANGELOG.md) - バージョン履歴
- [CONTRIBUTING.md](CONTRIBUTING.md) - コントリビューション方法
- [SECURITY.md](SECURITY.md) - セキュリティポリシーとベストプラクティス
- [v1.8/spec.md](v1.8/spec.md) - 技術仕様書（v1.8）

## セットアップ

### 必要環境

- Python 3.9 以上
- [uv](https://docs.astral.sh/uv/) パッケージマネージャー

### 依存関係のインストール

```bash
# uv をインストール（未インストールの場合）
# 詳細: https://docs.astral.sh/uv/getting-started/installation/
curl -LsSf https://astral.sh/uv/install.sh | sh

uv sync
```

## 使い方

```bash
uv run python v1.8/excel_to_md.py input.xlsx
```
これにより以下が生成されます:
- `input_csv.md`: CSVマークダウン形式（デフォルト）
- `input_images/`: 画像ディレクトリ（画像がある場合）

**注意**
- 出力ファイル名とディレクトリ名は入力ファイル名をベースに決定されます（例: `input.xlsx` → `input_csv.md`, `input_images/`）
- 入力ファイルと同じディレクトリに出力されます（`--csv-output-dir` で変更可能）

### よく使う例

**Mermaidフローチャート対応で変換:**
```bash
uv run python v1.8/excel_to_md.py input.xlsx --mermaid-enabled
```

**シートごとに個別ファイルを生成:**
```bash
uv run python v1.8/excel_to_md.py input.xlsx --split-by-sheet
```

**CSVマークダウンの出力先を指定:**
```bash
uv run python v1.8/excel_to_md.py input.xlsx --csv-output-dir ./output
# CSVマークダウン: ./output/input_csv.md
# 画像: ./output/input_images/
```

**標準Markdownのみ出力（CSV出力なし）:**
```bash
uv run python v1.8/excel_to_md.py input.xlsx -o output.md --no-csv-markdown-enabled
```

**平文ハイパーリンク（Markdown記法なし）:**
```bash
uv run python v1.8/excel_to_md.py input.xlsx --hyperlink-mode inline_plain
```

**トークン数削減（CSV概要セクション除外）:**
```bash
uv run python v1.8/excel_to_md.py input.xlsx --no-csv-include-description
```

## 主要オプション

### 出力制御

| オプション | デフォルト | 説明 |
|--------|---------|-------------|
| `--split-by-sheet` | false | シートごとに個別ファイルを生成 |
| `--csv-markdown-enabled` | true | CSVマークダウン出力を有効化 |
| `--csv-output-dir` | 入力ファイルと同じ | CSVマークダウンの出力先ディレクトリ |
| `--csv-include-description` | true | CSV出力に概要セクションを含める |
| `--csv-include-metadata` | true | CSV出力に検証メタデータを含める |
| `-o`, `--output` | - | 標準Markdownの出力ファイルパス |

### ハイパーリンク形式

| モード | 説明 | 出力例 |
|------|-------------|----------------|
| `inline` | Markdown形式 | `[テキスト](URL)` |
| `inline_plain` | 平文形式 | `テキスト (URL)` |
| `footnote` | 脚注形式 | `[テキスト][^1]` + `[^1]: URL` |
| `text_only` | 表示テキストのみ | `テキスト` |
| `both` | インライン+脚注 | 両方の形式 |

### Mermaidフローチャート

| オプション | デフォルト | 説明 |
|--------|---------|-------------|
| `--mermaid-enabled` | false | Mermaid変換を有効化 |
| `--mermaid-detect-mode` | shapes | 検出モード: `shapes`, `column_headers`, `heuristic` |
| `--mermaid-direction` | TD | フローチャート方向: `TD`, `LR`, `BT`, `RL` |
| `--mermaid-keep-source-table` | true | 元のテーブルもMermaidと一緒に出力 |

### テーブル処理

| オプション | デフォルト | 説明 |
|--------|---------|-------------|
| `--header-detection` | first_row | 先頭行をヘッダとして扱う |
| `--align-detection` | numbers_right | 数値列を右寄せ |
| `--max-cells-per-table` | 200000 | テーブルあたりの最大セル数 |
| `--no-print-area-mode` | used_range | 印刷領域未設定時の動作 |

## 出力例

### 標準Markdown出力

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
```

### CSVマークダウン出力

````markdown
# CSV出力: sample.xlsx

## 概要

### ファイル情報
- 元のExcelファイル名: sample.xlsx
- シート数: 2
- 生成日時: 2025-01-05 10:00:00

### このファイルについて
このCSVマークダウンファイルは、AIがExcelの内容を理解できるよう...

---

## Sheet1

```csv
品目,数量,備考
りんご,10,発注先
みかん,5,
```

---

## 検証用メタデータ

- **生成日時**: 2025-01-05 10:00:00
- **元Excelファイル**: sample.xlsx
- **検証ステータス**: OK
````

### 画像抽出

Excelファイル内の画像は自動的に処理されます:

1. **自動抽出**: 各シートの画像が外部ファイルとして保存されます
   - ファイル名形式: `{シート名}_img_{連番}.{拡張子}`
   - 例: `Sheet1_img_1.png`, `Sheet1_img_2.jpg`

2. **保存場所**: CSVマークダウンと同じディレクトリに出力
   - ディレクトリ名: `{入力ファイル名}_images/`
   - 例: `input.xlsx` → `input_images/` ディレクトリ
   - `--csv-output-dir` オプションで出力先を変更可能

3. **Markdownリンク**: 画像が配置されているセルにMarkdown画像リンクを生成
   - 形式: `![代替テキスト](相対パス)`
   - セル値がある場合は代替テキストとして使用
   - セル値がない場合は `Image at A1` のように自動生成

4. **対応形式**: PNG, JPEG, GIF

**例:**

Excelのセル位置 (B2) に会社ロゴ画像がある場合:
- 画像ファイル: `input_images/Sheet1_img_1.png` として保存
- CSV出力: `![Company Logo](input_images/Sheet1_img_1.png)`
- セルに "Company Logo" というテキストがあれば代替テキストとして使用

## 高度なオプション

全オプションの一覧:

```bash
uv run python v1.8/excel_to_md.py --help
```

主な高度なオプション:
- セル結合ポリシー
- 日付/数値フォーマット制御
- 空白処理
- Markdownエスケープレベル
- 非表示行/列ポリシー
- ロケール固有のフォーマット


## ディレクトリ構成

```
excel2md/
├── v1.8/                       # 最新バージョン
│   ├── excel_to_md.py          # メイン変換プログラム
│   ├── spec.md                 # 仕様書
│   └── tests/                  # テストスイート
├── v1.7/                       # 旧バージョン
│   ├── excel_to_md.py          # メイン変換プログラム
│   ├── spec.md                 # 仕様書
│   └── tests/                  # テストスイート
├── pyproject.toml          # プロジェクトメタデータ
├── LICENSE                 # MITライセンス
├── README.md               # このファイル
├── CONTRIBUTING.md         # コントリビューションガイド
├── SECURITY.md             # セキュリティポリシー
└── CHANGELOG.md            # バージョン履歴
```

## セキュリティ

セキュリティに関する懸念は [SECURITY.md](SECURITY.md) をご確認ください。

**主要なセキュリティ注意事項:**
- 信頼できるソースからのExcelファイルのみを処理してください
- `read_only=True` モードを使用してファイル変更を防止
- Excelマクロは実行しません
- Markdown出力をサニタイズしてインジェクションを防止

## コントリビューション

コントリビューションを歓迎します！詳細は [CONTRIBUTING.md](CONTRIBUTING.md) をご覧ください。

- バグ報告は [GitHub Issues](https://github.com/elvez/excel2md/issues) へ
- 改善のためのプルリクエストを提出
- 既存のコードスタイルに従ってください
- 新機能にはテストを追加してください

## 変更履歴

詳細は [CHANGELOG.md](CHANGELOG.md) を参照してください。

## 開発の背景

本ツールは、日本語の開発文書・仕様書を対象とした開発支援AI **IXV（イクシブ）** の開発過程で生まれた小さな実用品です。

IXVでは、システム開発における日本語の文書について、理解・構造化・活用という課題に取り組んでおり、本リポジトリでは、その一部を切り出して公開しています。

## ライセンス

MIT License - 詳細は [LICENSE](LICENSE) を参照してください。

## 問い合わせ先

- **メールアドレス**: info@elvez.co.jp
- **宛先**: 株式会社エルブズ
