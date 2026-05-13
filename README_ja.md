# excel2md

[English](https://github.com/elvezjp/excel2md/blob/main/README.md) | [日本語](https://github.com/elvezjp/excel2md/blob/main/README_ja.md)

[![Elvez](https://img.shields.io/badge/Elvez-Product-3F61A7?style=flat-square)](https://elvez.co.jp/)
[![IXV Ecosystem](https://img.shields.io/badge/IXV-Ecosystem-3F61A7?style=flat-square)](https://elvez.co.jp/ixv/)
[![PyPI version](https://img.shields.io/pypi/v/excel2md?style=flat-square)](https://pypi.org/project/excel2md/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.10+-blue?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Stars](https://img.shields.io/github/stars/elvezjp/excel2md?style=social)](https://github.com/elvezjp/excel2md/stargazers)

![excel2md 変換例: Excel シートから Markdown / CSV マークダウン / Mermaid フローチャートへ](https://raw.githubusercontent.com/elvezjp/excel2md/main/docs/assets/example.png)

Excel → Markdown 変換ツール。Excelブック（.xlsx/.xlsm）を読み取り、Markdown形式で自動生成します。

## 特徴

- **スマートテーブル検出**: Excel印刷領域を自動検出してMarkdownテーブルに変換
- **CSVマークダウン出力**: シート全体をCSV形式で出力（検証用メタデータ付き）
- **画像抽出**: Excelファイル内の画像を外部ファイルとして抽出し、Markdownリンク形式で出力
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

- [CHANGELOG_ja.md](https://github.com/elvezjp/excel2md/blob/main/CHANGELOG_ja.md) - バージョン履歴
- [CONTRIBUTING_ja.md](https://github.com/elvezjp/excel2md/blob/main/CONTRIBUTING_ja.md) - コントリビューション方法
- [SECURITY_ja.md](https://github.com/elvezjp/excel2md/blob/main/SECURITY_ja.md) - セキュリティポリシーとベストプラクティス
- [v2.2.0/spec.md](https://github.com/elvezjp/excel2md/blob/main/v2.2.0/spec.md) - 技術仕様書（v2.2.0, 最新）
- [v2.1.0/spec.md](https://github.com/elvezjp/excel2md/blob/main/v2.1.0/spec.md) - 技術仕様書（v2.1.0, 凍結スナップショット）
- [v1.8/spec.md](https://github.com/elvezjp/excel2md/blob/main/v1.8/spec.md) - 技術仕様書（v1.8）

## インストール

Python 3.10 以上が必要です。

```bash
pip install excel2md
# または uv の場合
uv add excel2md
```

## 使い方

```bash
excel2md input.xlsx
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
excel2md input.xlsx --mermaid-enabled
```

**シートごとに個別ファイルを生成:**
```bash
excel2md input.xlsx --split-by-sheet
```

**CSVマークダウンの出力先を指定:**
```bash
excel2md input.xlsx --csv-output-dir ./output
# CSVマークダウン: ./output/input_csv.md
# 画像: ./output/input_images/
```

**標準Markdownのみ出力（CSV出力なし）:**
```bash
excel2md input.xlsx -o output.md --no-csv-markdown-enabled
```

**平文ハイパーリンク（Markdown記法なし）:**
```bash
excel2md input.xlsx --hyperlink-mode inline_plain
```

**トークン数削減（CSV概要セクション除外）:**
```bash
excel2md input.xlsx --no-csv-include-description
```

## ライブラリとしての利用

`excel2md` は Python ライブラリとしても利用できます。

```python
from excel2md import convert_to_markdown

# パス、または xlsx の bytes を直接渡せる（Pyodide / Web アップロード向け）
result = convert_to_markdown("input.xlsx", csv_markdown_enabled=False)

print(result["markdown"])      # 生成された Markdown 文字列
print(result["output_path"])   # .md ファイルが書き出されたパス
```

CLI オプションはキーワード引数として 1:1 で受け取れます（例: `mermaid_enabled=True`, `split_by_sheet=True`）。同じ設定を使い回す場合は `ConversionConfig` + `ExcelConverter` を直接利用するのが効率的です。


### ソースから利用

```bash
git clone https://github.com/elvezjp/excel2md.git
cd excel2md
uv sync
```

詳細は [CONTRIBUTING_ja.md](https://github.com/elvezjp/excel2md/blob/main/CONTRIBUTING_ja.md) を参照してください。


## 主要オプション

### 出力制御

| オプション | デフォルト | 説明 |
|--------|---------|-------------|
| `--split-by-sheet` | false | シートごとに個別ファイルを生成 |
| `--csv-markdown-enabled` | true | CSVマークダウン出力を有効化 |
| `--csv-output-dir` | 入力ファイルと同じ | CSVマークダウンの出力先ディレクトリ |
| `--csv-include-description` | true | CSV出力に概要セクションを含める |
| `--csv-include-metadata` | true | CSV出力に検証メタデータを含める |
| `--image-extraction` | true | 画像抽出を有効化 |
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


### 高度なオプション

全オプションの一覧:

```bash
excel2md --help
```

主な高度なオプション:
- セル結合ポリシー
- 日付/数値フォーマット制御
- 空白処理
- Markdownエスケープレベル
- 非表示行/列ポリシー
- ロケール固有のフォーマット


## 出力例

実際の入出力サンプル（画像含む）は [docs/examples/](https://github.com/elvezjp/excel2md/tree/main/docs/examples) 配下にあります。各バージョンディレクトリには以下を含みます:

- 入力 `.xlsx` ファイル
- `output-default/` — デフォルト設定（CSV markdown + 画像抽出）
- `output-markdown/` — 標準 Markdown モード (`--no-csv-markdown-enabled`)
- `output-mermaid/` — Mermaid フローチャート有効 (`--mermaid-enabled`)

各パターンの再生成コマンドは [docs/examples/README.md](https://github.com/elvezjp/excel2md/blob/main/docs/examples/README.md) を参照してください。


## ディレクトリ構成

```
excel2md/
├── v2.2.0/                     # 最新バージョン
│   ├── excel_to_md.py          # エントリーポイント
│   ├── excel2md/               # メインパッケージ
│   ├── tests/                  # テストスイート
│   ├── spec.md                 # 仕様書
│   └── spec_appendix.md        # 仕様書付録
├── v2.1.1/                     # 旧バージョン（凍結スナップショット、PyPI 未公開）
├── v2.1.0/                     # 旧バージョン（凍結スナップショット）
├── v2.0.1/                     # 旧バージョン
├── v2.0/                       # 旧バージョン
├── v1.8/                       # 旧バージョン
│   ├── excel_to_md.py          # メイン変換プログラム
│   ├── spec.md                 # 仕様書
│   └── tests/                  # テストスイート
├── v1.7/                       # 旧バージョン
│   ├── excel_to_md.py          # メイン変換プログラム
│   ├── spec.md                 # 仕様書
│   └── tests/                  # テストスイート
├── docs/                   # ドキュメント
├── pyproject.toml          # プロジェクトメタデータ
├── LICENSE                 # MITライセンス
├── README.md / _ja.md     # README（英語 / 日本語）
├── CONTRIBUTING.md / _ja.md # コントリビューションガイド（英語 / 日本語）
├── SECURITY.md / _ja.md   # セキュリティポリシー（英語 / 日本語）
└── CHANGELOG.md / _ja.md  # バージョン履歴（英語 / 日本語）
```

## セキュリティ

セキュリティに関する懸念は [SECURITY_ja.md](https://github.com/elvezjp/excel2md/blob/main/SECURITY_ja.md) をご確認ください。

**主要なセキュリティ注意事項:**
- 信頼できるソースからのExcelファイルのみを処理してください
- `read_only=True` モードを使用してファイル変更を防止
- Excelマクロは実行しません
- Markdown出力をサニタイズしてインジェクションを防止

## コントリビューション

コントリビューションを歓迎します！詳細は [CONTRIBUTING_ja.md](https://github.com/elvezjp/excel2md/blob/main/CONTRIBUTING_ja.md) をご覧ください。

- バグ報告は [GitHub Issues](https://github.com/elvezjp/excel2md/issues) へ
- 改善のためのプルリクエストを提出
- 既存のコードスタイルに従ってください
- 新機能にはテストを追加してください

## 変更履歴

詳細は [CHANGELOG_ja.md](https://github.com/elvezjp/excel2md/blob/main/CHANGELOG_ja.md) を参照してください。

## 開発の背景

本ツールは、日本の開発現場でAIを活かすためのAI開発エコシステム **IXV（イクシブ）** の開発過程で生まれた小さな実用品です。

IXVでは、開発方法論とOSSを提供することで、AI活用を現場に根付かせる取り組みを進めており、本リポジトリでは、その一部を切り出して公開しています。

## ライセンス

MIT License - 詳細は [LICENSE](https://github.com/elvezjp/excel2md/blob/main/LICENSE) を参照してください。

## 問い合わせ先

- **メールアドレス**: info@elvez.co.jp
- **宛先**: 株式会社エルブズ
