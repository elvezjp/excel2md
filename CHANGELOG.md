# 変更履歴

このプロジェクトに対するすべての重要な変更はこのファイルに記録されます。

フォーマットは [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) に基づいており、
このプロジェクトは [セマンティックバージョニング](https://semver.org/spec/v2.0.0.html) に準拠しています。

## [1.7.0] - 2025-12-25

### 追加
- **CSVマークダウンでのMermaid出力対応**
  - `--mermaid-enabled` オプションがCSVマークダウンでも有効に
  - `mermaid_detect_mode="shapes"` の場合のみ対応（Excelの図形からフローチャート抽出）
  - `column_headers` / `heuristic` モードはCSVマークダウンでは非対応（WARNログを出力してスキップ）
  - 各シートのCSVブロック直後にMermaidコードブロックを出力

- **概要セクション除外オプション**
  - `--csv-include-description` / `--no-csv-include-description` オプションを追加
  - CSVマークダウンの概要セクション（説明文）を除外可能
  - 複数ファイルを変換・結合する際のトークン数削減に対応
  - デフォルトは `true`（従来通り概要セクションを出力）

### 変更
- v1.6との後方互換性を維持

## [1.6.0] - 2025-11-18

### 追加
- **ハイパーリンク平文出力モード（inline_plain）**
  - `--hyperlink-mode inline_plain` オプションを追加
  - セル内のハイパーリンクを平文形式で出力: `表示テキスト (URL)`
  - 内部リンクの場合: `表示テキスト (→場所)`
  - Markdown記法を使わずにリンク情報を明示的に表示

- **シート分割出力機能**
  - `--split-by-sheet` オプションを追加
  - 各シートを個別のMarkdownファイルとして出力
  - ファイル名形式: `{出力ファイル名}_{シート名}.md`
  - 各シートファイルには、シート名、仕様バージョン、元ファイル名を記載
  - シートごとに独立した脚注番号を使用

### 変更
- v1.5との後方互換性を維持

## [1.5.0] - 2025-11-11

### 追加
- **CSVマークダウン出力機能（デフォルト有効）**
  - ファイル名形式: `{basename}_csv.md`
  - 各シートの印刷領域をCSVコードブロックとして記載
  - 概要セクションと検証用メタデータセクションを自動生成
  - セル内改行を半角スペースに変換（1レコード=1行を保証）
  - ハイパーリンクは表示テキストのみ出力

- **バッチ処理対応**
  - `batch_test.py` をv1.5対応に更新
  - CSVマークダウン出力統計の表示機能を追加

- **新しいオプション**
  - `--csv-markdown-enabled` / `--no-csv-markdown-enabled`: CSVマークダウン出力の有効化/無効化
  - `--csv-output-dir`: CSVマークダウンの出力先ディレクトリ
  - `--csv-include-metadata` / `--no-csv-include-metadata`: 検証用メタデータを含めるか
  - `--csv-apply-merge-policy` / `--no-csv-apply-merge-policy`: CSV抽出時にmerge_policyを適用するか
  - `--csv-normalize-values` / `--no-csv-normalize-values`: CSV値に数値正規化を適用するか

### 変更
- v1.4との後方互換性を維持

## [1.4.0] - 2025-11-08

### 追加
- **Mermaidフローチャート変換機能**
  - 列名ベース検出: `From` / `To` / `Label` 列を検出してフローチャート化
  - ヒューリスティック検出: テーブル構造から自動判定
  - シェイプ検出: ExcelのDrawingML図形からフローチャートを抽出
  - ノードID自動生成、重複エッジ除去、サブグラフ対応

- **新しいオプション**
  - `--mermaid-enabled`: Mermaidフローチャート変換を有効化
  - `--mermaid-detect-mode`: 検出モード（`shapes` / `column_headers` / `heuristic`）
  - `--mermaid-direction`: フローチャートの方向（`TD` / `LR` / `BT` / `RL`）
  - `--mermaid-keep-source-table`: 元のテーブルも出力するか

## [1.3.0] - 2025-11-08

### 追加
- **基本機能の実装**
  - 最大長方形分解アルゴリズム（ヒストグラム法＋彫り抜き法）
  - 印刷領域と空セル判定
  - 結合セルと空判定
  - Markdown出力（テーブル形式）
  - ハイパーリンク処理（脚注形式）
  - パフォーマンス最適化と制限

- **基本オプション**
  - `-o`, `--output`: 出力ファイルパス
  - `--header-detection`: テーブル先頭行をヘッダとして扱う
  - `--align-detection`: 数値列の右寄せ判定（80%ルール）
  - `--no-print-area-mode`: 印刷領域未設定時の動作
  - `--max-cells-per-table`: テーブル1つあたりの最大セル数
  - `--markdown-escape-level`: Markdown記号のエスケープレベル
  - `--hyperlink-mode`: ハイパーリンクの出力方法
  - `--footnote-scope`: 脚注番号の採番スコープ

### 技術詳細
- Python 3.9以上をサポート
- openpyxl 3.1.5以上を依存ライブラリとして使用
- `read_only=True, data_only=True` モードで安全なファイル読み込み

## リンク

- [リポジトリ](https://github.com/elvez/excel2md)
- [Issue](https://github.com/elvez/excel2md/issues)

---

## バージョン比較

| バージョン | 主な機能 |
|------------|----------|
| 1.7.0      | CSVマークダウンモード拡張（Mermaid出力、説明文除外） |
| 1.6.0      | ハイパーリンク平文出力、シート分割出力 |
| 1.5.0      | CSVマークダウン出力 |
| 1.4.0      | Mermaidフローチャート変換 |
| 1.3.0      | 基本実装 |
