# Python 最低バージョン引き上げ修正計画書（v2.1.0）

> **ステータス**: 完了
> **ブランチ**: `tominaga/20260417-upgrade-python-min-version`（予定）
> **関連 Dependabot alerts**:
> - [#2 Pygments ReDoS (CVE-2026-4539)](https://github.com/elvezjp/excel2md/security/dependabot/2)
> - [#3 pytest tmpdir handling (CVE-2025-71176)](https://github.com/elvezjp/excel2md/security/dependabot/3)

---

## 1. 概要（仕様駆動開発 工程② 企画・要件定義）

### 1.1 背景と問題点

GitHub の Dependabot alerts により、開発用依存ライブラリに以下 2 件の脆弱性が検出された。

| # | パッケージ | Severity | CVE | 脆弱なバージョン | 修正バージョン |
|---|---|---|---|---|---|
| 1 | pytest | Medium | CVE-2025-71176 | `< 9.0.3` | `9.0.3` |
| 2 | Pygments | Low | CVE-2026-4539 | `< 2.20.0` | `2.20.0` |

Dependabot は Pygments の修正 PR を自動作成したが、pytest については **`Dependabot cannot update pytest to a non-vulnerable version`** として PR 作成を断念した。

**原因**: `pytest 9.0.3` は Python 3.10+ のみサポートしており、現在の `requires-python = ">=3.9"` の制約下では全 Python バージョンで動作するバージョンへ解決できない。

### 1.2 なぜ（Why）変更するか

- **セキュリティ**: 2 件の Dependabot alerts を解消する
- **Python 3.9 は 2025-10 に公式 EOL 済み**
- pytest 9.x への追従は今後のメンテナンスにおいても必要
- これらはすべて開発用依存（test extras）だが、CI 環境の脆弱性を残すべきではない

### 1.3 何を（What）変更するか

- `requires-python` を `>=3.9` → `>=3.10` に引き上げ
- classifiers から Python 3.9 を削除
- `pytest`, `Pygments` を含む dev 依存を最新パッチへ更新
- **バージョン方針**: **v2.0.1 → v2.1.0**（マイナーバージョン更新）
  - 理由: サポート Python の下限引き上げは利用者への互換性影響があり、SemVer 上はパッチではなくマイナー更新が適切

### 1.4 影響範囲

| 項目 | 影響 |
|---|---|
| エンドユーザーの実行環境 | Python 3.9 利用者は v2.1.0 以降を利用不可（v2.0.x 系は継続利用可能） |
| 本番依存（`openpyxl`） | 変更なし |
| 開発用依存（pytest, pytest-cov, Pygments 等） | 最新パッチに更新 |
| CI マトリクス | テスト対象を「最低サポート(3.10) + 現行最新(3.14)」に変更（従来: 3.9 / 3.12） |
| 既存ソースコード（`v2.0.1/excel2md/`） | 変更なし（v2.1.0 へコピーするのみ） |

---

## 2. 修正方針（仕様駆動開発 工程③ 設計計画）

- `v2.0.1/` をコピーして `v2.1.0/` を作成し、そこで作業を行う（既存リリース構造を踏襲）
- 仕様書（`spec.md`）を先に更新し、仕様と実装の整合性を確保する
- `pyproject.toml` の `requires-python` 更新と `uv.lock` 再生成により、脆弱性パッケージを修正版に置き換える
- CI ワークフローのテスト対象を「最低サポート(3.10) + 現行最新(3.14)」に変更する
- CHANGELOG / README などのドキュメント整合を取る

---

## 3. タスク一覧（仕様駆動開発 工程④ タスク分割）

### タスク1: v2.0.1 → v2.1.0 のコピー作成

**目的**: v2.1.0 作業用ディレクトリを作成する。

**手順**:
1. `v2.0.1/` を `v2.1.0/` としてコピー
2. `v2.1.0/excel2md/__init__.py` の `__version__` を `"2.1.0"` に更新

**実装記録**:
- [x] 完了
- 実施日: 2026-04-17
- 備考: `cp -r v2.0.1 v2.1.0` 実施済（コミット `c352694`）。`__version__` を `"2.1.0"` に、docstring も `"excel2md package (v2.1.0)."` に更新。

---

### タスク2: 仕様書（spec.md）の更新

**目的**: 対応 Python バージョンの記載を v2.1.0 基準に更新する。

**対象ファイル**: `v2.1.0/spec.md`（必要に応じて `spec_appendix.md` も）

**更新内容**:
- ヘッダを「excel2md v2.1 仕様書」に更新
- 前提条件・動作環境の記載がある場合、Python 3.10+ に更新
- 既存仕様に実装面の変更はないため、それ以外の変更は最小限に留める

**実装記録**:
- [x] 完了
- 実施日: 2026-04-17
- 備考: `v2.1.0/spec.md` および `v2.1.0/spec_appendix.md` のヘッダを「excel2md v2.1 仕様書」「excel2md v2.1 仕様書 付録」に更新。Python バージョンの前提条件・動作環境に関する記載は本仕様書内になかったため該当変更なし。

---

### タスク3: pyproject.toml の更新

**目的**: Python 最低バージョンとパッケージメタデータを更新する。

**対象ファイル**: `pyproject.toml`

**更新内容**:
- `version` を `"2.1.0"` に更新
- `requires-python = ">=3.10"` に更新
- classifiers から `"Programming Language :: Python :: 3.9"` を削除
- `testpaths` を `v2.1.0/tests` に更新
- `tool.coverage.run.source` / `omit` を `v2.1.0` に更新

**実装記録**:
- [x] 完了
- 実施日: 2026-04-17
- 備考: version・requires-python・classifiers（3.9 削除、3.13/3.14 追加）・testpaths・coverage source/omit を更新。

---

### タスク4: uv.lock の再生成

**目的**: 脆弱性のある pytest / Pygments を修正バージョンに置き換える。

**手順**:
1. 既存 `uv.lock` を削除
2. `uv sync` を実行し lockfile を再生成
3. 以下のバージョン更新を確認する
   - `pytest`: `9.0.2` → `9.0.3`（CVE-2025-71176 修正）
   - `Pygments`: `2.19.2` → `2.20.0`（CVE-2026-4539 修正）
   - その他の dev 依存（coverage / packaging / pytest-cov / tomli）の更新も許容

**実装記録**:
- [x] 完了
- 実施日: 2026-04-17
- 備考: `rm uv.lock && uv sync --all-extras` で再生成。`pytest==9.0.3` / `pygments==2.20.0` を確認。Python 3.10+ 制約となったため Python バージョン別の分岐解消、lockfile が単一解決に。`uv run pytest` で 253 tests passed を確認。

---

### タスク5: CI ワークフローの更新

**目的**: CI マトリクスを「最低サポートバージョン + 現行最新バージョン」のテスト方針に変更する。

**対象ファイル**: `.github/workflows/ci.yml`

**更新内容**:
- Python バージョンマトリクスを `["3.9", "3.12"]` → `["3.10", "3.14"]` に変更
  - `3.10`: 本リリースでの最低サポートバージョン
  - `3.14`: 現行の最新 Python バージョン
- OS マトリクス（ubuntu / windows / macos）は従来通り維持

**実装記録**:
- [x] 完了
- 実施日: 2026-04-17
- 備考: `python-version: ["3.10", "3.14"]` に変更。OS マトリクス（ubuntu / windows / macos）は維持。

---

### タスク6: README.md / README_ja.md の更新

**目的**: ドキュメントのパス参照・サポート Python バージョン記載を更新する。

**対象ファイル**: `README.md`, `README_ja.md`

**更新内容**:
- `v2.0.1/` へのパス参照を `v2.1.0/` に更新
- ディレクトリ構成の記載を更新
- サポート Python バージョンの記載があれば `3.10+` に更新

**実装記録**:
- [x] 完了
- 実施日: 2026-04-17
- 備考: Python バッジ（3.9+ → 3.10+）、Requirements/必要環境の記載（3.9 以上 → 3.10 以上）、仕様書リンク（v2.0.1 → v2.1.0）、コマンド例のパス（v2.0.1 → v2.1.0）、ディレクトリ構成（v2.1.0 を最新、v2.0.1 を旧バージョン欄へ）を両ファイルで更新。

---

### タスク7: CHANGELOG.md / CHANGELOG_ja.md に v2.1.0 エントリ追加

**目的**: バージョン履歴に v2.1.0 の変更内容を記録する。

**対象ファイル**: `CHANGELOG.md`, `CHANGELOG_ja.md`

**更新内容**:
- v2.1.0 のエントリを追加
  - Changed: サポート Python バージョンを 3.10+ に引き上げ（3.9 EOL 対応）
  - Security: pytest を 9.0.3 に更新（CVE-2025-71176）
  - Security: Pygments を 2.20.0 に更新（CVE-2026-4539）
- バージョン比較テーブルがあれば v2.1.0 行を追記

**実装記録**:
- [x] 完了
- 実施日: 2026-04-17
- 備考: v2.1.0 エントリ（Changed / Security / Documentation セクション）を英日両ファイルに追加し、バージョン比較テーブルにも v2.1.0 行を追記。

---

## 4. 検証・受入工程（仕様駆動開発 工程⑥）

タスク 1〜7 の実装完了後に実施する。本リリースはソースコードに変更を加えないメンテナンスリリースであり、**機能的な回帰がないこと**と**脆弱性が解消されたこと**の確認を主眼とする。

### 4.1 単体テストの実施

- [x] ローカル環境で `uv run pytest` が全件パスすること（253 件を目安）
- [x] CI が対象マトリクス（Python 3.10 / 3.14 × ubuntu / windows / macos）で全ジョブ成功すること（PR #22 にて全6ジョブ SUCCESS を確認）

### 4.2 前バージョン（v2.0.1）との出力比較

本リリースはソースコードの変更を伴わないため、同一の入力に対し v2.0.1 と v2.1.0 の変換結果は完全に一致する必要がある。

- [x] テスト用 Excel ファイルを v2.0.1 と v2.1.0 でそれぞれ変換し、出力 Markdown が同一であること
- [x] 主要オプション（`--mermaid-enabled`, `--no-csv-markdown-enabled` 等）の組み合わせで出力差分がないこと
- [x] 画像抽出・CSV Markdown 出力についても差分がないこと

### 4.3 Dependabot アラート解消確認

- [x] PR マージ後、Dependabot alert #2（Pygments）が自動クローズされること（`state: fixed`、2026-04-17 05:51:36 UTC）
- [x] PR マージ後、Dependabot alert #3（pytest）が自動クローズされること（`state: fixed`、2026-04-17 05:51:36 UTC）

### 検証実施記録

- 実施日: 2026-04-17
- 実施者: tominaga
- 単体テスト結果サマリ: `uv run pytest` で **253 tests passed** を確認（ローカル macOS / Python 3.14）。CI マトリクス（Python 3.10 / 3.14 × 3 OS）は PR 作成後に確認。
- 出力比較結果サマリ: 下記 4 ケースで差分が意図通りのもの（生成日時 / 仕様バージョン表記）のみであることを確認。
  - `test_standard.xlsx` + CSV Markdown 出力: 差分は「生成日時」のみ（実行時刻差）
  - `test_standard.xlsx` + 標準 Markdown 出力（`--no-csv-markdown-enabled`）: 差分は「仕様バージョン: 2.0.1 → 2.1.0」のみ（`__version__` の反映）
  - `test_standard.xlsx` の抽出画像（`test_standard_images/*`）: 差分なし（完全一致）
  - `test_mermaid.xlsx` + `--mermaid-enabled` + CSV Markdown 出力: 差分は「生成日時」のみ（実行時刻差）
- 備考: ソースコード無変更のメンテナンスリリースであり、機能的回帰がないことを確認済み。Dependabot アラートの自動クローズは PR マージ後に確認する。

---

## 5. 移行・運用工程（仕様駆動開発 工程⑦）

- [x] PR 作成 → レビュー → main マージ（PR #22、マージコミット `2f971cf`）
- [x] 既存の Dependabot 自動生成 PR #21（Pygments 単体更新）は本 PR と重複するためクローズ
