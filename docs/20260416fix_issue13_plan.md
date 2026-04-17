# Issue #13 修正計画書（v2.0.1）

> **ステータス**: 検証完了
> **ブランチ**: `tominaga/20260416-fix-mermaid-import-bug`
> **対象Issue**: [#13 2.0における潜在的なバグの可能性について](https://github.com/elvezjp/excel2md/issues/13)

---

## 1. 概要

v2.0 の `mermaid_generator.py` に以下2件のバグが存在する。

| # | 種別 | 内容 | 影響 |
|---|------|------|------|
| A | バグ | `is_code_block()` が import されていない | heuristic モードで `NameError` 発生 |
| B | 軽微 | `import re` と `import re as _re` が重複 | 動作影響なし。移植残りによるコード品質低下 |

## 2. 修正方針

- `v2.0/` をコピーして `v2.0.1/` を作成し、そこで修正を行う。
- 仕様書（`spec.md`）を先に更新し、修正内容と仕様の整合性を確保する。
- 既存テストの互換性を維持する。

---

## 3. タスク一覧

### タスク1: 仕様書（spec.md）の更新

**目的**: 修正内容を仕様に反映し、仕様と実装の整合性を確保する。

**対象ファイル**: `v2.0.1/spec.md`

**更新内容**:
- モジュール依存関係図に `mermaid_generator.py → table_formatting.py` の依存を追記
- heuristic 検出モードの判定条件に「コードブロック除外（`is_code_block`）」の前処理を明記

**実装記録**:
- [x] 完了
- 実施日: 2026-04-16
- 備考: 依存関係図に `mermaid_generator.py → table_formatting.py` を追記、heuristic 判定条件に条件0（コードブロック除外）を追記

---

### タスク2: v2.0 → v2.0.1 のコピー作成

**目的**: 修正用ディレクトリを作成する。

**手順**:
1. `v2.0/` を `v2.0.1/` としてコピー
2. `v2.0.1/excel2md/__init__.py` のバージョンを `"2.0.1"` に更新

**実装記録**:
- [x] 完了
- 実施日: 2026-04-16
- 備考: `cp -r v2.0 v2.0.1`、`__version__` を `"2.0.1"` に更新

---

### タスク3: `is_code_block` の import 追加（バグA）

**目的**: heuristic モードで `NameError` が発生するバグを修正する。

**対象ファイル**: `v2.0.1/excel2md/mermaid_generator.py`

**修正内容**:
- `from .table_formatting import is_code_block` を import に追加

**実装記録**:
- [x] 完了
- 実施日: 2026-04-16
- 備考: 9行目に `from .table_formatting import is_code_block` を追加

---

### タスク4: `import re` 重複の解消（バグB）

**目的**: 移植残りによる `import re` / `import re as _re` の重複を解消する。

**対象ファイル**: `v2.0.1/excel2md/mermaid_generator.py`

**修正内容**:
- `import re` を削除し、`import re as _re` に統一
- ファイル内の `re.` 参照を `_re.` に統一（既存の使用箇所を確認のうえ）

**実装記録**:
- [x] 完了
- 実施日: 2026-04-16
- 備考: `import re` を削除、103行目の `re.search` を `_re.search` に変更（`_re.` に統一）

---

### タスク5: pyproject.toml のバージョン・パス更新

**目的**: プロジェクトメタデータのバージョンとテスト・カバレッジパスを v2.0.1 に合わせる。

**対象ファイル**: `pyproject.toml`

**更新内容**:
- `version` を `"2.0.1"` に更新
- `testpaths` を `v2.0.1/tests` に更新
- `tool.coverage.run.source` を `v2.0.1` に更新

**実装記録**:
- [x] 完了
- 実施日: 2026-04-16
- 備考: version, testpaths, coverage source/omit の4箇所を更新

---

### タスク6: README.md / README_ja.md のパス更新

**目的**: ドキュメント内の v2.0 パス参照を v2.0.1 に更新する。

**対象ファイル**: `README.md`, `README_ja.md`

**更新内容**:
- `v2.0/` へのパス参照を `v2.0.1/` に更新
- ディレクトリ構成の記載を更新

**実装記録**:
- [x] 完了
- 実施日: 2026-04-16
- 備考: spec.md リンク、excel_to_md.py パス、ディレクトリ構成を両ファイルで更新

---

### タスク7: CHANGELOG.md に v2.0.1 エントリ追加

**目的**: バージョン履歴に v2.0.1 の変更内容を記録する。

**対象ファイル**: `CHANGELOG.md`

**更新内容**:
- v2.0.1 のエントリを追加（バグ修正: is_code_block import 漏れ、import re 重複解消）

**実装記録**:
- [x] 完了
- 実施日: 2026-04-16
- 備考: 修正・ドキュメントセクション追加、バージョン比較テーブルにも追記

---

---

## 4. 検証・受入工程

タスク1〜4の実装完了後に実施する。

### 4.1 バグA: `is_code_block` import 漏れの修正確認

- [x] `mermaid_generator.py` に `from .table_formatting import is_code_block` が存在すること
- [x] heuristic モードで `is_code_block` が正常に呼び出されること（`NameError` が発生しないこと）

### 4.2 バグB: `import re` 重複解消の確認

- [x] `mermaid_generator.py` 内に `import re` と `import re as _re` の重複が存在しないこと
- [x] `re.` / `_re.` の参照が統一されていること
- [x] 正規表現を使用する処理が正しく動作すること

### 4.3 リグレッション確認

- [x] 既存テストスイートが全件パスすること（253 passed）
- [ ] v2.0 の出力結果と v2.0.1 の出力結果が同一であること（バグ修正箇所を除く）※テスト用Excelファイルによる出力比較は未実施

### 4.4 仕様整合性確認

- [x] `spec.md` のモジュール依存関係図が実装と一致していること
- [x] `spec.md` の heuristic 検出モード記述が実装と一致していること

### 検証実施記録

- 実施日: 2026-04-16
- 実施者: tominaga
- テスト結果サマリ: 253 tests passed (0.22s)、バグ修正確認OK、仕様整合性OK
- 備考: v2.0 vs v2.0.1 の出力比較は手動確認が必要（テスト用Excelファイルでの実行比較）
