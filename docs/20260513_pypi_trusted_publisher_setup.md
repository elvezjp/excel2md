# PyPI Trusted Publisher セットアップ手順

- 作成日: 2026-05-13
- 対象 Issue: [#19 PyPI パッケージ公開対応](https://github.com/elvezjp/excel2md/issues/19)
- 関連 PR: 本ブランチ (`feat/issue-19-trusted-publisher-workflow`)
- 関連計画書: [docs/20260419pypi_publication_plan.md](20260419pypi_publication_plan.md) §3 #7

## 1. 概要

GitHub Actions から API トークンを使わずに PyPI / TestPyPI へ自動公開するための仕組みが [Trusted Publisher (OIDC)](https://docs.pypi.org/trusted-publishers/) です。
本ドキュメントは、`.github/workflows/publish.yml` および `.github/workflows/publish-testpypi.yml` を機能させるために、**コードの外で必要な管理者作業** をまとめます。

## 2. アーキテクチャ

```
[開発者]  git tag v2.1.1 && git push origin v2.1.1
                ↓
[GitHub]  tag push を検知して publish.yml を起動
                ↓
[publish.yml]  uv build → sdist + wheel
                ↓
[GitHub OIDC]  「これは elvezjp/excel2md の publish.yml ジョブ」と署名
                ↓
[PyPI]  Trusted Publisher 設定と一致するか検証
                ↓
        一致したら excel2md 2.1.1 を公開
```

API トークンは一切リポジトリに置きません。

## 3. 管理者の作業（公開前に 1 回だけ）

### 3.1 PyPI アカウント・プロジェクトの確保

1. PyPI (https://pypi.org/) でアカウント作成（既存なら省略）
2. **2FA を必ず有効化** — Trusted Publisher の前提として PyPI が要求するため
3. パッケージ名 `excel2md` が未取得であることを確認（2026-05-13 時点では未取得）
4. 同様に TestPyPI (https://test.pypi.org/) でもアカウント作成・2FA 有効化

### 3.2 PyPI 側で Trusted Publisher を登録

**PyPI 本番** (https://pypi.org/manage/account/publishing/) で「Pending publisher」を新規追加:

| 項目 | 値 |
|---|---|
| PyPI Project Name | `excel2md` |
| Owner | `elvezjp` |
| Repository name | `excel2md` |
| Workflow filename | `publish.yml` |
| Environment name | `pypi` |

「Add」を押すと Pending 状態になります。初回 publish 時にプロジェクトが自動作成され、Pending → Active に昇格します。

**TestPyPI** (https://test.pypi.org/manage/account/publishing/) でも同様に追加:

| 項目 | 値 |
|---|---|
| PyPI Project Name | `excel2md` |
| Owner | `elvezjp` |
| Repository name | `excel2md` |
| Workflow filename | `publish-testpypi.yml` |
| Environment name | `testpypi` |

### 3.3 GitHub 側で Environment を作成

リポジトリ Settings → Environments で 2 つ作成:

#### `pypi` Environment（本番用、保護必須）

- 「Required reviewers」に管理者を 1〜2 名指定 → 手動承認後に publish される
- 「Deployment branches」を `main` のみに制限（タグは main から打つ前提）
- Secret 不要（OIDC 経由のため）

#### `testpypi` Environment（リハーサル用、保護は緩め）

- Required reviewers は任意（無くても OK）
- Deployment branches 制限も任意
- Secret 不要

## 4. リリース実行手順

### 4.1 リハーサル（TestPyPI）

```bash
# pyproject.toml の version を rc / dev 系に上げてから
git checkout main
# version を 2.1.2.dev0 などに変更してコミット
git push origin main
```

GitHub Actions の workflow_dispatch で `publish-testpypi.yml` を手動起動。完了後:

```bash
# 別環境 (clean venv) で TestPyPI からインストール確認
uv venv /tmp/verify && source /tmp/verify/bin/activate
uv pip install --index-url https://test.pypi.org/simple/ \
               --extra-index-url https://pypi.org/simple/ \
               excel2md==2.1.2.dev0
python -c "import excel2md; print(excel2md.__version__)"
excel2md --help
```

### 4.2 本番リリース（PyPI）

```bash
# pyproject.toml の version を確定 (例 2.1.2)
git checkout main && git pull --ff-only
# (version 変更を commit & push 済みの想定)
git tag v2.1.2
git push origin v2.1.2
```

GitHub Actions が `publish.yml` を起動 → `pypi` Environment の reviewer が承認 → PyPI に公開。

完了後の確認:

```bash
pip install excel2md
python -c "import excel2md; print(excel2md.__version__)"
excel2md --help
```

PyPI ページ: https://pypi.org/project/excel2md/

### 4.3 GitHub Release の作成

tag push と並行して [GitHub Releases](https://github.com/elvezjp/excel2md/releases/new) で Release を作成し、CHANGELOG の該当バージョンの抜粋を貼る。これにより:

- GitHub 上でバージョン履歴が一覧化される
- `pip install excel2md==X.Y.Z` で固定インストールするユーザーが、対応する変更内容を辿りやすくなる

## 5. トラブルシュート

### Trusted Publisher 認証エラー

ログに `Token request failed` / `invalid-publisher` のような文言が出る場合、次を確認:

1. PyPI 側の Pending publisher 設定が現在の `owner` / `repo` / `workflow filename` / `environment` と完全一致しているか
2. GitHub Environment 名が `pypi` / `testpypi` と一致しているか（大文字小文字も含めて）
3. `permissions: id-token: write` がジョブに設定されているか

### TestPyPI に同バージョンを再アップロードしたい

TestPyPI も PyPI と同じく、同一バージョンの再アップロードは拒否されます。リハーサルを繰り返す場合は `version` を `2.1.2.dev0` → `2.1.2.dev1` → ... と dev / rc サフィックスで進める運用が無難です。

### 公開後にバグが発覚したら

PyPI は **すでに公開された wheel の差し替えを禁止** しています。代わりに:

1. `yank`: バグありの版を「インストール非推奨」にマーク（残り続けるがデフォルトでは新規インストールされない）
2. 修正版を `2.1.2.post1` などのポストリリース、または `2.1.3` で公開

`yank` は PyPI Web UI のプロジェクト管理画面から実行可能。

## 6. 受け入れ条件

- [ ] PyPI 本番・TestPyPI に `excel2md` の Trusted Publisher が登録されている
- [ ] GitHub Environments `pypi` / `testpypi` が作成されている
- [ ] `pypi` Environment に Required reviewers が設定されている
- [ ] テストとして TestPyPI へ 1 度 publish できた
- [ ] `pip install --index-url https://test.pypi.org/simple/ excel2md` が成功する
- [ ] CLI (`excel2md --help`) と import (`from excel2md import convert_to_markdown`) が動作する
- [ ] 本番 PyPI に `excel2md` を公開できた
- [ ] GitHub Release が同時に作成された

## 7. このドキュメント自体について

このドキュメントは管理者作業のチェックリスト兼ガイドです。実装コード（`.github/workflows/*.yml`）は本 PR でマージできますが、上記 §3 の管理者作業が完了するまで、ワークフローは認証エラーで動きません。マージのタイミングと管理者作業の完了は **歩調を合わせる** 必要があります。
