# Examples

[English](https://github.com/elvezjp/excel2md/blob/main/docs/examples/README.md) | 日本語

各バージョンの `excel2md` 実行結果のサンプルです。入力ファイル (xlsx) と出力結果 (md / png / jpg) を含みます。

## ディレクトリ構成

```
docs/examples/
└── バージョン番号/
    ├── test_standard.xlsx              # 入力: 標準テーブル / 結合セル / ハイパーリンク / 画像
    ├── test_mermaid.xlsx               # 入力: Mermaid 検出用のフロー記述
    ├── output-default/                 # デフォルト設定 (CSV markdown + 画像抽出)
    │   ├── test_standard_csv.md
    │   └── test_standard_images/
    ├── output-markdown/                # 標準 Markdown モード (--no-csv-markdown-enabled)
    │   └── test_standard.md
    └── output-mermaid/                 # Mermaid 有効 (--mermaid-enabled)
        └── test_mermaid_csv.md
```

## 入力ファイル

| ファイル | 内容 |
|---|---|
| `test_standard.xlsx` | 標準的なテーブル、結合セル、複数テーブル、ハイパーリンク、画像を含む 5 シート |
| `test_mermaid.xlsx` | `From` / `To` / `Label` 列を持つフローテーブル（Mermaid フローチャート検出のサンプル） |

## 生成方法

対象バージョンのディレクトリで作業する。先頭で `VERSION` を設定しておくと以下のコマンドをそのまま使い回せる。

```bash
VERSION=vX.Y.Z

# デフォルト設定 (CSV markdown + 画像抽出)
excel2md docs/examples/${VERSION}/test_standard.xlsx \
  --csv-output-dir docs/examples/${VERSION}/output-default

# 標準 Markdown モード (CSV markdown 無効)
excel2md docs/examples/${VERSION}/test_standard.xlsx \
  -o docs/examples/${VERSION}/output-markdown/test_standard.md \
  --no-csv-markdown-enabled

# Mermaid フローチャート有効
excel2md docs/examples/${VERSION}/test_mermaid.xlsx \
  --mermaid-enabled \
  --csv-output-dir docs/examples/${VERSION}/output-mermaid
```

ソースから動かす場合は各コマンドの先頭に `uv run` を付ける。
