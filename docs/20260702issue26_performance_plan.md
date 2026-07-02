# Issue #26 性能改善 実施計画（v2.3.0）

- 作成日: 2026-07-02
- 対象 Issue: [#26 性能改善の優先実施案: CSV専用経路とキャッシュ導入](https://github.com/elvezjp/excel2md/issues/26)
- 関連 Issue: [#10 パフォーマンス最適化](https://github.com/elvezjp/excel2md/issues/10)（本対応をもって解決扱いとする）
- 方針コメント: https://github.com/elvezjp/excel2md/issues/26#issuecomment-4863335354

## 背景

v2.2.1 の実装確認により、Issue #26 の指摘が現在も有効であることを確認した。
効果の大きさと実装後のコードの綺麗さの観点から、以下の3件に絞って対応する。

## 対応スコープ

### ① CSV Markdown 出力時の通常 Markdown 処理スキップ + shapes Mermaid 抽出の1回化

- `runner.run()` はデフォルト（`csv_markdown_enabled=true`）でも通常 Markdown 用の全処理
  （テーブル検出 → 抽出 → 形式判定 → 整形 → 脚注管理）を実行したうえで結果を破棄している。
  CSV のみ出力するケースではこれらをスキップする。
- `mermaid_enabled` + shapes モード時に `_v14_extract_shapes_to_mermaid()` が
  通常 Markdown 用と CSV 用で同一引数のまま2回実行される（各回 ZIP オープン + XML パース +
  全セルグリッド構築）。シートごとに1回だけ計算し結果を使い回す。
- 出力内容は変更しない。既存テストが回帰検証となる。

### ② `build_merged_lookup()` の再利用

- 同一 `union_area` に対して runner 側（通常 Markdown 用・CSV 用）と `grid_to_tables()` 内部の
  計3箇所で同じ lookup が構築されている。
- `grid_to_tables(..., merged_lookup=None)` の形で外から渡せるようにし、
  runner でエリアごとに1回だけ構築して再利用する。

### ③ DrawingML 解析の統合（`WorkbookDrawingIndex` 導入）

- 「シート名 → sheet_id → drawing path」の解決ロジックが image_extraction.py と
  mermaid_generator.py にほぼ同一コードとして重複実装されており、シートごとに ZIP を開き直している。
- ワークブック単位で1回だけ workbook.xml / workbook.xml.rels / sheet rels を解析する
  索引構造 `WorkbookDrawingIndex` を導入し、画像抽出と Mermaid 図形抽出の両方から参照する。
- 性能改善に加え、約60行×2箇所のコード重複を解消する。

## 見送る項目

- セル値・空判定・スタイル判定キャッシュ: ①によりセル評価の実行回数が大幅に減るため、
  効果を計測してから導入要否を判断する。
- multiprocessing / async I/O: Issue #26 の非目標どおり導入しない。

## 別 Issue に切り出す事項

- 印刷領域が複数矩形に分かれるシートで `csv_markdown_data[sname]` がエリアごとに上書きされ、
  最後の矩形しか CSV に残らない挙動（バグの疑い）。本対応では出力を変えない範囲に留める。

## 前提（Issue #11 との関係）

- Issue #11（バージョン管理の改善）の対応により、実装コードは今後リポジトリルートに配置される。
- そのため本対応では最初に v2.2.1 の内容をリポジトリルートへコピーしてコミットし、
  以降はルート側のコードのみを編集する。`v2.2.1/` と `versions/` は凍結スナップショットとして触らない。
- 新バージョンは v2.3.0 とし、ドキュメント類（CHANGELOG / spec.md / README ほか）を更新する。

## 実施ステップ

各ステップ完了時に本ドキュメントの進捗欄を更新し、コミットする。

| # | ステップ | 内容 | 状態 |
|---|---------|------|------|
| 1 | 計画書作成 | 本ドキュメントの作成 | 完了 |
| 2 | ルートへコピー | `v2.2.1/` の内容（excel2md/, tests/, excel_to_md.py, spec.md, spec_appendix.md）をリポジトリルートへコピーし、pyproject.toml のパス参照をルートに変更 | 完了 |
| 3 | ベンチマーク追加 | 生成 fixture + 区間タイマーの簡易ベンチマークスクリプトを追加し、改善前ベースラインを計測 | 完了 |
| 4 | 実装① | CSV専用経路スキップ + shapes Mermaid 1回化 | 完了 |
| 5 | 実装② | `build_merged_lookup()` の再利用 | 未着手 |
| 6 | 実装③ | `WorkbookDrawingIndex` 導入 | 未着手 |
| 7 | 効果計測 | 改善後ベンチマークを計測し本ドキュメントに記録 | 未着手 |
| 8 | ドキュメント更新 | v2.3.0 として CHANGELOG / CHANGELOG_ja / spec.md / README / `__version__` / pyproject.toml を更新 | 未着手 |

## 受け入れ条件（Issue #26 より）

- [ ] CSV Markdown のみ出力する既定ケースで、通常 Markdown 用整形処理が実行されない
- [ ] 既存の CSV Markdown 出力内容が変わらない
- [ ] `uv run pytest` が通る
- [ ] ベンチマークで処理時間の改善を確認できる

## 計測記録

### 計測環境・ケース

- スクリプト: `scripts/benchmark_issue26.py`（`uv run python scripts/benchmark_issue26.py`）
- 計測値は3回実行の best / median（秒）。変換全体（`runner.run()`、出力書き込み含む）を計測。
- ケース:
  - `50k_cells`: 500行 × 100列（デフォルト設定 = CSV Markdown 出力）
  - `200k_cells`: 2000行 × 100列（同上）
  - `merged_cells`: 500行 × 100列 + 結合セル多数（同上）
  - `multi_sheet_mermaid`: 10シート × 100行 × 100列、mermaid shapes モード有効
  - `mermaid_fixture`: tests/fixtures/test_mermaid.xlsx × 30回、mermaid 有効

### 改善前ベースライン

| ケース | best | median |
|--------|------|--------|
| 50k_cells | 1.045s | 1.049s |
| 200k_cells | 4.865s | 4.891s |
| merged_cells | 1.042s | 1.058s |
| multi_sheet_mermaid | 2.009s | 2.016s |
| mermaid_fixture | 0.169s | 0.169s |

### 改善後

（ステップ7で記載）

## 進捗記録

- 2026-07-02: 計画書作成（ステップ1完了）
- 2026-07-02: v2.2.1 の内容をリポジトリルートへコピーし、pyproject.toml のパス参照（wheel packages / sdist include・exclude / pytest testpaths / coverage）をルートに変更。`uv run --extra test pytest` で327件全件通過を確認（ステップ2完了）
- 2026-07-02: `scripts/benchmark_issue26.py` を追加し、改善前ベースラインを計測・記録（ステップ3完了）
- 2026-07-02: 実装①完了（ステップ4完了）。`runner.run()` に `emit_normal_md` フラグを導入し、CSV Markdown 出力時（デフォルト）は通常 Markdown 用の組み立て（テーブル検出・抽出・形式判定・整形・脚注管理）をスキップ。shapes モードの Mermaid 抽出はシートごとに1回だけ実行し通常/CSV 両経路で共用。回帰テスト3件追加（計330件通過）。凍結スナップショット v2.2.1 と全 fixture の変換出力が一致（生成日時行を除く）することを確認。この時点の計測: 50k_cells 0.446s / 200k_cells 1.844s / merged_cells 0.448s / multi_sheet_mermaid 0.907s / mermaid_fixture 0.134s
