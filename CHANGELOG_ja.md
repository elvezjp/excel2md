# 変更履歴

[English](./CHANGELOG.md) | [日本語](./CHANGELOG_ja.md)

このプロジェクトに対するすべての重要な変更はこのファイルに記録されます。

フォーマットは [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) に基づいており、
このプロジェクトは [セマンティックバージョニング](https://semver.org/spec/v2.0.0.html) に準拠しています。

## [2.3.0] - 2026-07-02

### パフォーマンス
- **CSV Markdown出力有効時（デフォルト）に通常Markdown処理をスキップ** ([#26](https://github.com/elvezjp/excel2md/issues/26), [#10](https://github.com/elvezjp/excel2md/issues/10))
  - 従来の `runner.run()` は CSV Markdown 出力時にもシートごとにテーブル検出・抽出・形式判定・整形・脚注管理を実行し、結果を破棄していた
  - 変換出力は不変。ベンチマーク: 50,000セル 1.05秒→0.45秒、200,000セル 4.87秒→1.88秒（約6割短縮）
- **shapesモードのMermaid抽出をシートごとに1回化** ([#26](https://github.com/elvezjp/excel2md/issues/26))
  - 従来は同一引数で2回実行され（破棄される通常Markdown用とCSV Markdown用）、その都度 xlsx ZIP のオープンと全セルグリッド構築が発生していた
- **`grid_to_tables()` に構築済み `merged_lookup` を渡せるように変更** ([#26](https://github.com/elvezjp/excel2md/issues/26))
  - 同一印刷領域に対する結合セルマップの再構築を解消。未指定時は従来どおり内部構築するため既存の呼び出しに影響なし

### 追加
- **`excel2md.drawing_index.WorkbookDrawingIndex`** — xlsx ZIP（workbook.xml / workbook.xml.rels / sheet rels）をワークブックにつき1回だけ解析し、シート名→drawingパス・画像リレーションを解決する索引 ([#26](https://github.com/elvezjp/excel2md/issues/26))
  - `image_extraction.py` と `mermaid_generator.py` に重複実装されシートごとに ZIP を開き直していた約60行×2箇所の解決ロジックを置き換え
  - 図形の無いシートでは Mermaid 用セルグリッド構築を省略するように改善
- 生成 fixture（5万/20万セル、結合セル、複数シートMermaid）で変換時間を計測するベンチマークスクリプト `scripts/benchmark_issue26.py`

### 変更
- リポジトリ構成: バージョン番号入りのソースディレクトリ（`v2.2.1/`）と `versions/` スナップショットアーカイブを廃止し、リポジトリ直下のフラットな構成に変更。過去のリリースはgit履歴とこのCHANGELOGに引き続き保持されている。`docs/examples/` も同様にフラット化。`versions/` 廃止により重複しなくなったため、SECURITY.md の「旧バージョンのDependabotアラートはDismissする」という運用方針は削除（[#11](https://github.com/elvezjp/excel2md/issues/11)）

### テスト
- `TestIssue26CsvOnlySkipsNormalMarkdown`（CSV専用経路でテーブル検出・抽出が実行されないこと、shapes Mermaid がシートごとに1回だけ実行されること）と `tests/test_drawing_index.py`（6件）を追加 — 計336件

### ドキュメント
- 実施計画書・計測記録を `docs/20260702issue26_performance_plan.md` に追加
- `spec.md` を v2.3 に更新（モジュール一覧・依存関係・全体処理フロー・ルート基準のコマンド表記）
- `docs/examples/` 配下の実行例を v2.3.0 で再生成。生成日時とバージョン表記を除き従来の出力と一致し、性能改善が変換出力へ影響しないことを確認。`-o` 指定の再生成コマンドに必要な `mkdir -p` の前提を `docs/examples/README.md` に追記

## [2.2.1] - 2026-05-14

### 追加
- **ライブラリ層の例外型: `ExcelConversionError` (基底) と `WorkbookOpenError` を新設** ([#36](https://github.com/elvezjp/excel2md/issues/36))
  - `excel2md` および `excel_to_md` から公開。ライブラリ利用者は文字列マッチに頼らず例外型で変換エラーを捕捉できる

### 変更
- **`load_workbook_safe()` がオープン失敗時に `sys.exit(2)` を呼ばず、`WorkbookOpenError` を raise するように変更** ([#36](https://github.com/elvezjp/excel2md/issues/36))
  - 従来は不正なパスや壊れた bytes を `ExcelConverter` / `convert_to_markdown` に渡すと呼び出し元プロセス全体が終了していた。PyPI 公開済みの再利用可能 API としては受け入れがたい挙動 (Pyodide / MCP サーバー / Notebook / Web サービス向け)
  - 元の openpyxl 例外は `raise ... from e` で保持 (`__cause__` から参照可能)
  - CLI 挙動は不変: `cli.main()` で `ExcelConversionError` を catch して stderr に `[ERROR] ...` を出力し、exit code 2 で終了する従来挙動を維持
- `excel2md.__version__` を `"2.2.1"` に更新 (v2.2.0 → v2.2.1 ディレクトリコピー時に `"2.2.0"` のままになっていた)

### 修正
- **印刷領域外の画像がある場合に CSV Markdown が印刷領域外のコンテンツを取り込んでしまう問題を修正** ([#14](https://github.com/elvezjp/excel2md/issues/14))
  - `runner.run()` が `extract_images_from_sheet` から返された画像位置をすべて含むように `union_area` を拡張していたため、印刷領域外に置かれた画像が CSV 出力レンジを引きずって広げ、関係のないセル値まで巻き込んでいた
  - この拡張は v2.1.0 仕様（§「印刷領域内のみが変換対象となる」）および v1.8 までの挙動と矛盾していた
  - 拡張ブロックを削除し、`extract_print_area_for_csv()` が印刷領域内のみを反復するようにした。印刷領域外の `cell_to_image` エントリは自然に無視される
  - 補足: `extract_images_from_sheet` 自体は呼び出されたままなので、印刷領域外の画像ファイルがディスクに保存される副作用は残る。これは別タスクで追跡しており、本修正のスコープ外

### 変更
- 新しい開発用ディレクトリ `v2.2.1/` を追加。`v2.2.0/` は v2.2.0 リリース時点の凍結スナップショットとして保持（リポのバージョン管理ポリシーに準拠 — 既存 v*/ は凍結し、新バージョンは新規ディレクトリで作業）
- **リポジトリ構成: 過去バージョンディレクトリ（`v1.7` / `v1.8` / `v2.0` / `v2.0.1` / `v2.1.0` / `v2.1.1` / `v2.2.0`）を `versions/` 配下に集約** ([#11](https://github.com/elvezjp/excel2md/issues/11))
  - 現在開発中のディレクトリ（`v2.2.1/`）はリポジトリルートに残置
  - `pyproject.toml` の sdist `exclude` を `versions/**` に統合。PyPI に配布される wheel/sdist の中身は不変
  - 移動はすべて git rename として記録されており、履歴は引き続き辿れる
- `versions/README.md` を追加し、本ディレクトリの目的と PyPI 配布対象外である旨を明記

### テスト
- `tests/test_runner_regression.py` に `TestIssue14CsvPrintAreaRespect` を追加。印刷領域 `A1:B2` + 領域外画像 `(5, 5)` + 領域外セル値 (`C5`, `D6`) の workbook で、出力レンジが `A1:B2` のままで、領域外画像リンクおよび領域外セル値が CSV に混入しないことを assert する
- **`.xlsm`（マクロ有効ブック）を正式サポート対象として試験を追加** ([#43](https://github.com/elvezjp/excel2md/issues/43))
  - `v2.2.1/tests/fixtures/test_macro.xlsm` を追加（基本テーブル + VBA マクロ入りシート）
  - `v2.2.1/tests/test_xlsm_support.py` を新設。openpyxl 既定 (`keep_vba=False`) で VBA バイナリが破棄されること、`workbook_loader.load_workbook_safe` 経由で開けること、パス／バイト列入力どちらでも変換が成功すること、同等内容の `.xlsx` とテーブル内容がパリティすること、変換でマクロ起因の副作用が出ないことを回帰確認する

### ドキュメント
- `v2.2.1/spec.md` §3.1.1 / §11.3 に `.xlsm` のサポート範囲（読み込みのみ・VBA 破棄・マクロ非実行）とフィクスチャ説明を追記 ([#43](https://github.com/elvezjp/excel2md/issues/43))
- `SECURITY.md` / `SECURITY_ja.md` に `openpyxl.load_workbook(..., keep_vba=False)` で VBA を破棄し、`Auto_Open` / `Workbook_Open` などの自動実行マクロも発火しない旨を明記 ([#43](https://github.com/elvezjp/excel2md/issues/43))
- `docs/examples/v2.2.1/` に v2.2.1 用の実行サンプル（`test_standard.xlsx` / `test_mermaid.xlsx` / `test_macro.xlsm` の変換例）を追加し、`docs/examples/README.md` に `.xlsm` 経路の生成コマンドを追記

## [2.2.0] - 2026-05-13

### 追加
- **ライブラリ公開 API**: `ConversionConfig` / `ExcelConverter` / `convert_to_markdown` を `excel2md` および `excel_to_md` から公開 ([#16](https://github.com/elvezjp/excel2md/issues/16))
  - `convert_to_markdown(data: bytes | str | Path, **opts) -> dict` — CLI 経由でない呼び出し元（Pyodide / MCP サーバー / ノートブック / Web サービス）向けのワンショット便宜関数
  - `ConversionConfig` — CLI オプションを反映した型ヒント付き dataclass。`from_args()` と `to_opts_dict()` を提供
  - `ExcelConverter` — config を保持して bytes 入力 / dict 出力で変換する再利用可能クラス
- pure Python 実装のため、Pyodide でも `micropip.install('excel2md')` でそのまま動作（ネイティブ依存なし）
- PyPI / TestPyPI 公開用 GitHub Actions ワークフロー（Trusted Publisher 方式、API トークン不要）を追加 ([#19](https://github.com/elvezjp/excel2md/issues/19))
  - `.github/workflows/publish.yml` — tag 起点の本番リリース
  - `.github/workflows/publish-testpypi.yml` — 手動トリガーの TestPyPI リハーサル
  - 管理者セットアップ手順は [docs/20260513_pypi_trusted_publisher_setup.md](docs/20260513_pypi_trusted_publisher_setup.md) を参照
- **PyPI に `excel2md` として初公開**
- 検証用スタンドアロン CLI を `excel2md-verify` として `project.scripts` に登録（[#38](https://github.com/elvezjp/excel2md/issues/38)）

### 修正
- **`verify_csv_markdown` モジュールが配布物に含まれず、CSV Markdown 検証メタデータの追記が silent に失敗していたバグを修正** ([#38](https://github.com/elvezjp/excel2md/issues/38))
  - `csv_export.py` が `sys.path` 細工経由で兄弟ファイル `v2.2.0/verify_csv_markdown.py` を import していたが、wheel/sdist には同モジュールが含まれていなかった
  - 開発時は cwd が自動的に `sys.path` に入る偶然動作で解決されていただけで、`pip install` 後のユーザーでは `[WARN] Failed to append verification metadata: No module named 'verify_csv_markdown'` が発生し、検証メタデータが欠落していた
  - `verify_csv_markdown.py` を `v2.2.0/excel2md/` 配下に移動し、相対 import (`from .verify_csv_markdown import ...`) に変更

### 変更
- `runner.run()` は内部で `argparse.Namespace`（CLI 経路）と `ConversionConfig`（ライブラリ経路）の両方を受け取り、`ConversionConfig` に正規化する。CLI 挙動は不変。従来 inline で組み立てていた options 辞書は `ConversionConfig.to_opts_dict()` への委譲に置換され、roundtrip テストで旧 inline 辞書との完全一致を保証
- `pyproject.toml` の `authors` に `email = "info@elvez.co.jp"` を追加（PyPI メタデータの連絡先明示）
- v2.2.0 開発用ディレクトリ `v2.2.0/` を新規追加。`v2.1.1/` は v2.1.1 リリース時点（commit 034fa57）の凍結スナップショットとしてそのまま残す（リポのバージョン管理ポリシーに準拠 — 既存 v*/ は凍結し、新バージョンは新規ディレクトリで作業）

### 備考
- `v2.1.1/` は内部バージョン番号として割り振っただけで PyPI には公開されていない。`v2.2.0` が PyPI 初公開バージョン
- `v2.1.1/` ディレクトリは凍結スナップショットとして残してあるが、本リリースで導入したライブラリ API（`ConversionConfig` / `ExcelConverter` / `convert_to_markdown`）は含まない（v2.2.0 で初導入）

## [2.1.1] - 2026-05-11

### 修正
- **`is_code_block` / `build_code_block_from_rows` の v1.x 互換 re-export を復活** ([#15](https://github.com/elvezjp/excel2md/issues/15))
  - v2.0 で両関数が `excel2md.table_formatting` に切り出され、トップレベルの `excel2md` / `excel_to_md` から import できなくなっていた
  - `excel2md/__init__.py` および `excel_to_md.py` から再エクスポートし、v1.8 時点の公開 API 互換性を回復
- **`extract_table()` の打ち切り戻り値の要素数が一致しないバグを修正** ([#24](https://github.com/elvezjp/excel2md/issues/24))
  - `max_cells_per_table` 超過時に 3 要素を返していたが、`runner.run()` 側は 4 要素 unpack を要求しており `ValueError` が発生していた
  - 打ち切り経路でも `(md_rows, note_refs, True, table_title)` の 4 要素戻り値に統一
- **複数テーブルで脚注番号が重複するバグを修正** ([#25](https://github.com/elvezjp/excel2md/issues/25))
  - `runner.run()` が各テーブルに同じ `global_footnote_start` を渡していたため、`[^1]` が再採番され参照先が曖昧になっていた
  - テーブル処理後に `len(note_refs)` 分だけ開始番号を前進させ、`footnote_scope=book` ではブック内連番、`footnote_scope=sheet` ではシート単位リセットを正しく動作させる
- **`--split-by-sheet` 未指定時にシートスコープの脚注定義が出力されないバグを修正**
  - `footnote_scope=sheet` を `--split-by-sheet` なしで指定した場合、シート末尾の脚注定義ブロックが出力されなかった
  - 各シートのセクション末尾に脚注定義を出力するよう修正

### テスト
- `tests/test_public_api.py` を追加し、v1.x 公開 API の再エクスポートを回帰検証
- `tests/test_runner_regression.py` を追加し、打ち切り戻り値・脚注番号の回帰を検証

### 備考
- `v2.1.0/` は v2.1.0 リリース時点の凍結スナップショットとしてそのまま残してある。本リリースの修正は `v2.1.1/` 配下に集約されている。

## [2.1.0] - 2026-04-17

### 変更
- **サポートする最低 Python バージョンを 3.10 に引き上げ**
  - Python 3.9 は 2025-10 に公式 EOL を迎え、サポート対象外となりました
  - `requires-python` を `>=3.10` に更新
  - CI マトリクスを最低サポート（3.10）と現行最新（3.14）の 2 バージョンに変更

### セキュリティ
- **pytest を 9.0.3 に更新** ([CVE-2025-71176](https://github.com/advisories/GHSA-6w46-j5rx-g56g))
  - pytest の tmpdir 処理における脆弱性を修正
- **Pygments を 2.20.0 に更新** ([CVE-2026-4539](https://github.com/advisories/GHSA-5239-wwwm-4pmq))
  - GUID マッチング用の非効率な正規表現に起因する ReDoS を修正

### ドキュメント
- spec.md / spec_appendix.md のヘッダを v2.1 に更新
- README.md / README_ja.md の Python バッジおよびパス参照を更新

## [2.0.1] - 2026-04-16

### 修正
- **mermaid_generator.py の `is_code_block` import 漏れを修正** ([#13](https://github.com/elvezjp/excel2md/issues/13))
  - heuristic 検出モードで `NameError` が発生するバグを修正
  - `from .table_formatting import is_code_block` を追加

- **`import re` の重複を解消** ([#13](https://github.com/elvezjp/excel2md/issues/13))
  - `import re` と `import re as _re` の重複を削除（v1.8 からの移植残り）
  - `_re` に統一

### ドキュメント
- 仕様書（spec.md）のモジュール依存関係図を修正
- 仕様書の heuristic 検出モード判定条件にコードブロック除外を明記

## [2.0.0] - 2026-01-26

### 変更
- **コードベースのモジュール化**
  - 単一実装ファイルを機能別モジュールに分割
  - `excel2md/` パッケージとして再構成

### ドキュメント
- 仕様書の構成を整理
- 詳細・補足は付録として分離

### テスト
- モジュール別テストスイートを追加

### 互換性
- v1.8との機能互換性を維持

## [1.8.0] - 2026-01-24

### 追加
- **画像抽出機能**
  - Excelファイル内の画像を外部ファイルとして自動抽出
  - 画像ファイル形式: `{シート名}_img_{連番}.{拡張子}`
  - 保存先: Markdownファイル名をベースにしたサブディレクトリ
  - 対応フォーマット: PNG, JPEG, GIF
  - 画像フォーマット自動判定（format属性またはマジックバイト検出）
  - セル位置の自動特定（TwoCellAnchor, OneCellAnchor対応）
  - Markdownリンクの自動生成: `![代替テキスト](相対パス)`
  - CSV Markdownモードでも画像リンクが有効
  - セル値を代替テキストとして使用（空の場合はセル参照を生成）
  - エラー時のグレースフルな処理（画像スキップして続行）

### テスト
- 画像抽出機能の包括的なユニットテストを追加（18テストケース）
  - 画像フォーマット検出テスト（PNG, JPEG, GIF）
  - アンカー位置抽出テスト（TwoCellAnchor, OneCellAnchor）
  - CSV抽出との統合テスト
  - エラーハンドリングとエッジケーステスト
  - 実際のopenpyxlワークシートを使用した統合テスト

### ドキュメント
- README.mdに画像抽出機能の説明を追加
  - 特徴リストに画像抽出を追加
  - 使用例と出力例を追加
  - 画像抽出の詳細な動作説明を追加
- spec.md（v1.7）に技術仕様を追加
  - §7.8 画像抽出とMarkdownリンク生成セクションを追加
  - 画像処理フロー、フォーマット判定、エラーハンドリングを文書化

### コード品質
- PEP 8スタイルガイドラインに準拠
- 包括的なdocstringsを追加（PEP 257準拠）
- 複雑なロジックに詳細なインラインコメントを追加
- より説明的な変数名を使用（ext → file_extension等）

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

- [リポジトリ](https://github.com/elvezjp/excel2md)
- [Issue](https://github.com/elvezjp/excel2md/issues)

---

## バージョン比較

| バージョン | 主な機能 |
|------------|----------|
| 2.3.0      | 性能改善: CSV出力時の通常Markdown処理スキップ、ワークブック単位DrawingML索引 (#26, #10)、ソースをリポジトリルートへ移動 |
| 2.2.1      | ライブラリ例外 (`ExcelConversionError` / `WorkbookOpenError`) (#36)、CSV印刷領域の修正 (#14)、`.xlsm` テスト整備 (#43)、旧バージョンを `versions/` へ移動 |
| 2.2.0      | ライブラリAPI (`ConversionConfig` / `ExcelConverter` / `convert_to_markdown`) (#16)、PyPI 初回公開 (#19) |
| 2.1.1      | バグ修正: v1.x re-export (#15)、打ち切り戻り値 (#24)、脚注番号 (#25) |
| 2.1.0      | 最低 Python バージョンを 3.10 に引き上げ、セキュリティ更新（pytest・Pygments） |
| 2.0.1      | mermaid_generator.py のバグ修正（import 漏れ・重複解消） |
| 2.0.0      | コードベースのモジュール化 |
| 1.8.0      | 画像抽出機能（Excelファイル内の画像を外部ファイルとして抽出） |
| 1.7.0      | CSVマークダウンモード拡張（Mermaid出力、説明文除外） |
| 1.6.0      | ハイパーリンク平文出力、シート分割出力 |
| 1.5.0      | CSVマークダウン出力 |
| 1.4.0      | Mermaidフローチャート変換 |
| 1.3.0      | 基本実装 |
