# excel2md

Excel → Markdown 変換ツール。Excelブック（.xlsx/.xlsm）を読み取り、Markdown形式で自動生成します。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)

## 特徴

- **スマートテーブル検出**: Excel印刷領域を自動検出してMarkdownテーブルに変換
- **CSVマークダウン出力**: シート全体をCSV形式で出力（検証用メタデータ付き）
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
- [v1.7/spec.md](v1.7/spec.md) - 技術仕様書

## プロジェクト構成

```
excel2md/
├── v1.7/
│   ├── excel_to_md.py      # メイン変換プログラム（最新版）
│   ├── spec.md             # 仕様書
│   └── tests/              # テストスイート
├── pyproject.toml          # プロジェクトメタデータ
├── LICENSE                 # MITライセンス
├── README.md               # このファイル
├── CONTRIBUTING.md         # コントリビューションガイド
├── SECURITY.md             # セキュリティポリシー
└── CHANGELOG.md            # バージョン履歴
```

## クイックスタート

### 必要環境

- Python 3.9 以上
- openpyxl 3.1.5 以上

### 依存関係のインストール

```bash
pip install openpyxl
```

または uv を使用:

```bash
uv pip install openpyxl
```

### 基本的な使い方

```bash
python3 v1.7/excel_to_md.py input.xlsx -o output.md
```

これにより以下が生成されます:
- `output.md` - 標準Markdownテーブル形式
- `input_csv.md` - CSVマークダウン形式（デフォルトで有効）

### よく使う例

**Mermaidフローチャート対応で変換:**
```bash
python3 v1.7/excel_to_md.py input.xlsx -o output.md --mermaid-enabled
```

**シートごとに個別ファイルを生成:**
```bash
python3 v1.7/excel_to_md.py input.xlsx -o output.md --split-by-sheet
```

**標準Markdownのみ出力（CSV出力なし）:**
```bash
python3 v1.7/excel_to_md.py input.xlsx -o output.md --no-csv-markdown-enabled
```

**平文ハイパーリンク（Markdown記法なし）:**
```bash
python3 v1.7/excel_to_md.py input.xlsx -o output.md --hyperlink-mode inline_plain
```

**トークン数削減（CSV概要セクション除外）:**
```bash
python3 v1.7/excel_to_md.py input.xlsx -o output.md --no-csv-include-description
```

## 主要オプション

### 出力制御

| オプション | デフォルト | 説明 |
|--------|---------|-------------|
| `-o`, `--output` | - | 出力ファイルパス |
| `--split-by-sheet` | false | シートごとに個別ファイルを生成 |
| `--csv-markdown-enabled` | true | CSVマークダウン出力を有効化 |
| `--csv-include-description` | true | CSV出力に概要セクションを含める |
| `--csv-include-metadata` | true | CSV出力に検証メタデータを含める |

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

## 高度なオプション

全オプションの一覧:

```bash
python3 v1.7/excel_to_md.py --help
```

主な高度なオプション:
- セル結合ポリシー
- 日付/数値フォーマット制御
- 空白処理
- Markdownエスケープレベル
- 非表示行/列ポリシー
- ロケール固有のフォーマット

## コントリビューション

コントリビューションを歓迎します！詳細は [CONTRIBUTING.md](CONTRIBUTING.md) をご覧ください。

- バグ報告は [GitHub Issues](https://github.com/elvez/excel2md/issues) へ
- 改善のためのプルリクエストを提出
- 既存のコードスタイルに従ってください
- 新機能にはテストを追加してください

## 問い合わせ先

バグ報告や質問については、以下までご連絡ください。

- **メールアドレス**: info@elvez.co.jp
- **宛先**: 株式会社エルブズ

GitHub Issuesでの報告も可能です: [GitHub Issues](https://github.com/elvez/excel2md/issues)

## セキュリティ

セキュリティに関する懸念は [SECURITY.md](SECURITY.md) をご確認ください。

**主要なセキュリティ注意事項:**
- 信頼できるソースからのExcelファイルのみを処理してください
- `read_only=True` モードを使用してファイル変更を防止
- Excelマクロは実行しません
- Markdown出力をサニタイズしてインジェクションを防止

## ライセンス

このプロジェクトはMITライセンスの下でライセンスされています - 詳細は [LICENSE](LICENSE) ファイルをご覧ください。

## リンク

- [リポジトリ](https://github.com/elvez/excel2md)
- [Issues](https://github.com/elvez/excel2md/issues)
