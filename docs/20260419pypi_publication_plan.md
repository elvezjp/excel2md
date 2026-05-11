# PyPI 公開対応 計画書

- 作成日: 2026-04-19
- 対象 Issue: [#19 PyPI パッケージ公開対応](https://github.com/elvezjp/excel2md/issues/19)
- 対象バージョン: **v2.1.1（最新バージョン）**
- 担当: claude (on behalf of repository maintainer)

## 1. 背景と目的

`excel2md` を他の Python プロジェクトから `pip install excel2md` で再利用できるようにしたい。
現在のリポジトリ構成では `pyproject.toml` にビルドバックエンドやパッケージ検出設定が欠落しており、
`python -m build` / `uv build` でのパッケージングに失敗する。

本計画では、最新バージョン v2.1.1 を PyPI へ公開可能な構成へ整備する。
（旧バージョン v1.7 / v1.8 / v2.0 / v2.0.1 / v2.1.0 はパッケージ対象外とする）

## 2. 公開に必要な仕様

### 2.1 リポジトリ構成の前提

- ソースコード：`v2.1.1/excel2md/`（Python パッケージ本体）
- CLI エントリポイント：`v2.1.1/excel2md/cli.py` の `main()`
- 互換レイヤー：`v2.1.1/excel_to_md.py`（テストから参照される再エクスポートモジュール。PyPI に含めるかは任意）
- 旧版：`v1.7/` `v1.8/` `v2.0/` `v2.0.1/` `v2.1.0/`（sdist/wheel には含めない）
- ドキュメント：`docs/`（sdist からは除外する）
- ライセンス：`LICENSE`（MIT。sdist/wheel に含める）

### 2.2 パッケージ仕様（PyPI 配布物）

| 項目 | 値 |
| ---- | ---- |
| パッケージ名 | `excel2md` |
| バージョン | `2.1.1`（現状のまま） |
| Python 要件 | `>=3.10` |
| ランタイム依存 | `openpyxl>=3.1.5` |
| ライセンス | MIT |
| 配布形式 | sdist + wheel |
| ビルドバックエンド | `hatchling` |
| wheel 内の Top-level パッケージ | `excel2md` |
| CLI | `excel2md` コマンド（`excel2md.cli:main`） |

### 2.3 pyproject.toml に追加する設定（Issue #19 準拠）

1. `[build-system]` セクションの追加
   ```toml
   [build-system]
   requires = ["hatchling"]
   build-backend = "hatchling.build"
   ```
2. wheel 用パッケージ検出：`[tool.hatch.build.targets.wheel]` で `v2.1.1/excel2md` を `excel2md` としてパッケージング
3. sdist 用 include/exclude：`v2.1.1/excel2md` 配下のみをパッケージング対象とし、旧バージョン・`docs/` を除外
4. `[project.scripts]` に `excel2md = "excel2md.cli:main"` を追加
5. `[project.urls]` は v2.1.1 で既に `elvezjp/excel2md` に修正済 → 変更不要

## 3. 公開までの全体計画

| # | フェーズ | 内容 | 本PRでの対応 |
| - | ------- | ---- | ------------- |
| 1 | パッケージング構成整備 | `pyproject.toml` を修正し、wheel/sdist を正しく作れるようにする | ✅ 対応 |
| 2 | ビルド検証 | `uv build` / `python -m build` で sdist/wheel が生成できることを確認 | ✅ 対応 |
| 3 | ロック更新 | `uv.lock` を最新化 | ✅ 対応 |
| 4 | 単体テスト通過確認 | `uv run pytest` が全件成功することを確認 | ✅ 対応 |
| 5 | TestPyPI 公開 | TestPyPI にアップロードし、別環境で `pip install` / CLI 動作確認 | ⏳ 本PRマージ後に管理者が実施 |
| 6 | PyPI 公開 | PyPI 本番へアップロード | ⏳ 本PRマージ後に管理者が実施 |
| 7 | GitHub Actions 自動公開 | Trusted Publisher 方式での publish ワークフロー追加 | ⏳ 本PR対象外（別Issue/PR 推奨） |
| 8 | パッケージ名の空き確認 | `excel2md` が PyPI で未取得であることを確認 | ⏳ 本PR対象外（管理者確認） |

## 4. 本PRのタスク詳細（実施スコープ）

### 4.1 pyproject.toml の修正
- `[build-system]` に `hatchling` を追加
- `[tool.hatch.build.targets.wheel]` で `v2.1.1/excel2md` をパッケージングソースとして指定
- `[tool.hatch.build.targets.sdist]` で v2.1.1 配下と必要な同梱ファイル（README/LICENSE/CHANGELOG 等）のみを include
- `[project.scripts]` に CLI エントリポイント `excel2md = "excel2md.cli:main"` を追加

### 4.2 uv.lock の再生成
- 既存の `uv.lock` を削除し、`uv lock` を再実行して最新化

### 4.3 ビルド検証
- `uv build` で sdist / wheel を生成し、成果物の中身（top-level に `excel2md/` があること、旧版や `docs/` が混入しないこと）を確認

### 4.4 テスト通過確認
- `uv run pytest` を実行し、261 件（v2.1.1 時点）が全て成功することを確認
- 旧バージョン（v1.7/v1.8/v2.0/v2.0.1/v2.1.0）のテスト修正は対象外

### 4.5 ドキュメント更新（最小限）
- 本計画書に実装記録と検証結果を追記

## 5. 本PRのスコープ外

- 旧バージョンディレクトリ（v1.7/v1.8/v2.0/v2.0.1/v2.1.0）の削除や修正
- バージョン番号の更新（現状の 2.1.1 のまま公開する想定）
- GitHub Actions による自動公開ワークフローの追加
- TestPyPI / PyPI への実アップロード
- README（ライブラリ利用方法）の大幅改訂

## 6. 検証・受け入れ項目（管理者確認用）

以下をチェックリストとして管理者がレビュー・受入れしてください。

### 6.1 pyproject.toml 構成
- [ ] `[build-system]` が hatchling で設定されている
- [ ] `[tool.hatch.build.targets.wheel]` に `v2.1.1/excel2md` が指定され、wheel 内で `excel2md` パッケージとして展開される
- [ ] `[tool.hatch.build.targets.sdist]` で旧バージョン・docs が除外されている
- [ ] `[project.scripts]` に `excel2md = "excel2md.cli:main"` が登録されている
- [ ] `[project.urls]` が `elvezjp/excel2md` を指している

### 6.2 ビルド / インストール検証
- [ ] `uv build` が成功し、`dist/excel2md-2.1.1-*.whl` と `dist/excel2md-2.1.1.tar.gz` が生成される
- [ ] 生成された wheel の中身に `excel2md/` ディレクトリがトップレベルで含まれる
- [ ] 生成された sdist に旧バージョン（v1.7/v1.8/v2.0/v2.0.1/v2.1.0）と `docs/` が含まれない
- [ ] クリーンな仮想環境で `pip install dist/excel2md-2.1.1-*.whl` が成功する
- [ ] `python -c "import excel2md; print(excel2md.__version__)"` が `2.1.1` を出力する
- [ ] インストール後に `excel2md --help` が動作する

### 6.3 テスト
- [ ] `uv run pytest` が全件パス（v2.1.1 時点 261 件）
- [ ] `uv.lock` が再生成され、コミット済み

### 6.4 次アクション（管理者実施想定）
- [ ] PyPI でのパッケージ名 `excel2md` の空き確認
- [ ] TestPyPI にアップロードし、別環境で CLI / import 確認
- [ ] PyPI 本番へアップロード
- [ ] GitHub Actions での Trusted Publisher 自動公開化（別Issue）

## 7. 実装記録

（本セクションは実装中に追記）

### 7.1 変更ファイル

- `pyproject.toml` — PyPI 公開用設定を追加
- `uv.lock` — 再生成
- `docs/20260419pypi_publication_plan.md` — 本計画書

### 7.2 実装の流れ

1. `pyproject.toml` に `[build-system]` / `[tool.hatch.build.targets.wheel|sdist]` / `[project.scripts]` を追加
2. `uv.lock` を削除し、`uv lock` で再生成
3. `uv sync --all-extras` で環境再構築
4. `uv run pytest` で全テストがパスすることを確認
5. `uv build` でパッケージングが成功することを確認

### 7.3 検証結果

- `uv run pytest`: 261 passed（v2.1.1 リベース後）
- `uv build`: sdist / wheel の生成成功
- wheel 内容: `excel2md/` ディレクトリがトップレベルにあり、`excel2md-2.1.1.dist-info/` を含む
- sdist 内容: `excel2md-2.1.1/` 配下に `excel2md/` のみ（旧バージョン・docs なし）
