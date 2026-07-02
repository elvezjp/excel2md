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
| 2 | ルートへコピー | `v2.2.1/` の内容（excel2md/, tests/, excel_to_md.py, spec.md, spec_appendix.md）をリポジトリルートへコピーし、pyproject.toml のパス参照をルートに変更 | 完了（後日 PR #50 の git mv による移動で代替） |
| 3 | ベンチマーク追加 | 生成 fixture + 区間タイマーの簡易ベンチマークスクリプトを追加し、改善前ベースラインを計測 | 完了 |
| 4 | 実装① | CSV専用経路スキップ + shapes Mermaid 1回化 | 完了 |
| 5 | 実装② | `build_merged_lookup()` の再利用 | 完了 |
| 6 | 実装③ | `WorkbookDrawingIndex` 導入 | 完了 |
| 7 | 効果計測 | 改善後ベンチマークを計測し本ドキュメントに記録 | 完了 |
| 8 | ドキュメント更新 | v2.3.0 として CHANGELOG / CHANGELOG_ja / spec.md / README / `__version__` / pyproject.toml を更新 | 完了 |

## 受け入れ条件（Issue #26 より）

- [x] CSV Markdown のみ出力する既定ケースで、通常 Markdown 用整形処理が実行されない
      （回帰テスト `TestIssue26CsvOnlySkipsNormalMarkdown` で担保）
- [x] 既存の CSV Markdown 出力内容が変わらない
      （凍結スナップショット v2.2.1 と全 fixture の変換出力一致を確認。生成日時行を除く）
- [x] `uv run pytest` が通る（336件通過）
- [x] ベンチマークで処理時間の改善を確認できる（下記計測記録参照。50k〜200kセルで約6割短縮）

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

### 改善後（実装①②③適用後）

| ケース | best | median | ベースライン比 (best) |
|--------|------|--------|----------------------|
| 50k_cells | 0.445s | 0.446s | -57% |
| 200k_cells | 1.876s | 1.927s | -61% |
| merged_cells | 0.457s | 0.458s | -56% |
| multi_sheet_mermaid | 0.877s | 0.883s | -56% |
| mermaid_fixture | 0.120s | 0.121s | -29% |

- 効果の大部分は実装①（CSV専用経路スキップ + Mermaid 1回化）による。
- 実装③は DrawingML 系ケース（mermaid_fixture 0.134s→0.120s、multi_sheet_mermaid 0.907s→0.877s）にさらに効いた。
- 実装②の単独効果は計測誤差の範囲内（結合セルの多寡によらず数%程度）だが、重複構築の解消としてコード意図が明確になった。

## 進捗記録

- 2026-07-02: 計画書作成（ステップ1完了）
- 2026-07-02: v2.2.1 の内容をリポジトリルートへコピーし、pyproject.toml のパス参照（wheel packages / sdist include・exclude / pytest testpaths / coverage）をルートに変更。`uv run --extra test pytest` で327件全件通過を確認（ステップ2完了）
- 2026-07-02: `scripts/benchmark_issue26.py` を追加し、改善前ベースラインを計測・記録（ステップ3完了）
- 2026-07-02: 実装①完了（ステップ4完了）。`runner.run()` に `emit_normal_md` フラグを導入し、CSV Markdown 出力時（デフォルト）は通常 Markdown 用の組み立て（テーブル検出・抽出・形式判定・整形・脚注管理）をスキップ。shapes モードの Mermaid 抽出はシートごとに1回だけ実行し通常/CSV 両経路で共用。回帰テスト3件追加（計330件通過）。凍結スナップショット v2.2.1 と全 fixture の変換出力が一致（生成日時行を除く）することを確認。この時点の計測: 50k_cells 0.446s / 200k_cells 1.844s / merged_cells 0.448s / multi_sheet_mermaid 0.907s / mermaid_fixture 0.134s
- 2026-07-02: 実装②完了（ステップ5完了）。`grid_to_tables()` に `merged_lookup` 引数を追加し、runner で構築済みの lookup を再利用（未指定時は従来どおり内部構築、既存呼び出し互換）。実装①で通常MD経路とCSV経路が排他になったため、残っていた重複は「runner構築 → grid_to_tables 内部で再構築」の2重構築であり、これを解消。テスト330件通過
- 2026-07-02: 実装③完了（ステップ6完了）。`excel2md/drawing_index.py` に `WorkbookDrawingIndex` を新設し、workbook.xml / workbook.xml.rels / sheet rels の解析をワークブック単位で1回に集約。image_extraction / mermaid_generator の重複実装（各約60行）を索引参照に置換し、runner が索引のライフサイクル（生成・close）を管理。図形の無いシートではセルグリッド構築を省くよう判定順も改善。索引のユニットテスト6件を追加（計336件通過）。凍結スナップショット v2.2.1 との出力一致を再確認。改善後ベンチマークを計測・記録（ステップ7完了）
- 2026-07-02: ドキュメント更新完了（ステップ8完了）。`excel2md.__version__` / pyproject.toml を 2.3.0 に更新。CHANGELOG / CHANGELOG_ja に 2.3.0 エントリと末尾バージョン一覧表（2.2.0/2.2.1 の欠落分含む）を追加。spec.md を v2.3 に更新（モジュール一覧・依存関係・§4.1 全体処理フロー・ルート基準のコマンド表記）。README / README_ja のディレクトリ構成図をルート配置に更新。`uv build` で sdist/wheel がルートの excel2md/ を含み v2.2.1/・versions/ を除外することを確認。全336テスト通過
- 2026-07-02: Issue #11 対応の [PR #50](https://github.com/elvezjp/excel2md/pull/50)（v2.2.1/ と versions/ を廃止しルート直下のフラット構成へ移行、バージョン 2.3.0 化を含む）を先行してマージ。本 Issue のブランチを新しい main の上に載せ替えた（ソースコードの性能改善コミットはそのまま適用、ステップ2のコピーコミットは PR #50 の移動で代替、CHANGELOG は #11 分と #26 分を単一の 2.3.0 エントリに統合、README は PR #50 の構成図を採用し scripts/ 行を追記、spec.md は本 Issue 側の v2.3 更新を採用）
- 2026-07-02: フラット化された docs/examples/ の実行例を v2.3.0 で再生成。差分が「生成日時」行と「仕様バージョン」行のみで画像は無変更であることを確認し、性能改善が出力へ影響しないことを実サンプルで再確認。テスト336件通過
