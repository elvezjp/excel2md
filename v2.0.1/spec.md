# excel2md v2.0 仕様書

## 1. 目的・概要

### 1.1 目的

本ツールは、Microsoft Excelファイル（.xlsx / .xlsm）をMarkdown形式に変換するCLIツールである。Excelのテーブル、フローチャート図形、埋め込み画像を解析し、構造化されたMarkdownドキュメントとして出力する。

### 1.2 主要機能

| 機能カテゴリ | 機能概要 |
|-------------|---------|
| テーブル変換 | Excel表をMarkdownテーブルに変換 |
| Mermaid生成 | フローチャート図形・テーブルからMermaidコードを自動生成 |
| 画像抽出 | 埋め込み画像をファイル出力しMarkdownリンクとして参照 |
| CSV Markdown | 検証可能なCSV形式での出力（メタデータ付） |
| ハイパーリンク | 多様な形式（インライン、脚注、テキストのみ等）での出力 |

### 1.3 設計思想

- **堅牢性**: エラー発生時も処理を継続し、部分的な出力を確保
- **柔軟性**: 35以上のオプションによる細かな動作制御
- **モジュール性**: 機能ごとに分離されたモジュール構成

---

## 2. モジュール構成

### 2.1 モジュール一覧

```
excel2md/
├── __init__.py           # パッケージ初期化・バージョン定義
├── cli.py                # CLIエントリーポイント・引数解析
├── runner.py             # メイン処理オーケストレーション
├── output.py             # ログ出力ユーティリティ
├── cell_utils.py         # セル値取得・テキスト正規化
├── workbook_loader.py    # Excelワークブック読み込み・座標変換
├── table_detection.py    # テーブル検出・矩形分解
├── table_extraction.py   # テーブル抽出・セル値取得
├── table_formatting.py   # テーブル形式判定・Markdown生成
├── mermaid_generator.py  # Mermaidフローチャート生成
├── image_extraction.py   # 画像抽出・ファイル出力
└── csv_export.py         # CSV Markdown出力
```

### 2.2 モジュール依存関係

```
cli.py
  └─→ runner.py
        ├─→ workbook_loader.py
        ├─→ table_detection.py
        │     └─→ cell_utils.py
        ├─→ table_extraction.py
        │     └─→ cell_utils.py
        ├─→ table_formatting.py
        ├─→ mermaid_generator.py
        ├─→ image_extraction.py
        └─→ csv_export.py
              └─→ output.py
```

### 2.3 各モジュールの責務

| モジュール | 責務 |
|-----------|------|
| cli.py | コマンドライン引数の解析、オプション辞書の構築 |
| runner.py | 処理フロー全体の制御、各モジュールの呼び出し順序管理 |
| cell_utils.py | セル値の取得、日付・数値・通貨のフォーマット、制御文字除去 |
| workbook_loader.py | openpyxlによるワークブック読み込み、座標変換（A1形式⇔数値） |
| table_detection.py | 非空セルの検出、連結成分解析、テーブル境界の特定 |
| table_extraction.py | テーブルデータの抽出、結合セル処理、ハイパーリンク処理 |
| table_formatting.py | 出力形式（コード/Mermaid/テキスト/ネスト/テーブル）の判定と生成 |
| mermaid_generator.py | Mermaidコードの生成（図形検出、ヒューリスティック検出対応） |
| image_extraction.py | DrawingML解析による画像抽出、ファイル出力 |
| csv_export.py | CSV形式Markdown出力、検証用メタデータ生成 |
| output.py | 警告・情報メッセージの標準エラー出力 |

---

## 3. 入出力仕様

### 3.1 入力

#### 3.1.1 対応ファイル形式

| 形式 | 拡張子 | 備考 |
|------|--------|------|
| Excel Open XML Workbook | .xlsx | 標準形式 |
| Excel Open XML Macro-Enabled Workbook | .xlsm | マクロ有効ブック |

#### 3.1.2 主要設定オプション

**印刷領域・範囲制御**

| オプション | 既定値 | 説明 |
|-----------|--------|------|
| no_print_area_mode | used_range | 印刷領域未設定時の処理（used_range / entire_sheet_range / skip_sheet） |
| max_sheet_count | 無制限 | 処理対象シート数の上限 |
| max_cells_per_table | 200000 | テーブルあたりの最大セル数 |
| read_only | false | 読み取り専用モード（スタイル情報制限あり） |

**セル値処理**

| オプション | 既定値 | 説明 |
|-----------|--------|------|
| value_mode | display | 値取得モード（display / formula / both） |
| merge_policy | top_left_only | 結合セル処理（top_left_only / repeat / expand / warn） |
| strip_whitespace | true | 前後空白の除去 |
| date_format_override | なし | 日付フォーマット強制指定 |
| date_default_format | YYYY-MM-DD | 日付の既定フォーマット |

**数値処理**

| オプション | 既定値 | 説明 |
|-----------|--------|------|
| numeric_thousand_sep | keep | 千単位区切り（keep / remove） |
| percent_format | keep | パーセント表記（keep / numeric） |
| currency_symbol | keep | 通貨記号（keep / strip） |

**ハイパーリンク**

| オプション | 既定値 | 説明 |
|-----------|--------|------|
| hyperlink_mode | inline | リンク出力形式（inline / inline_plain / footnote / both / text_only） |

**テーブル検出・表示**

| オプション | 既定値 | 説明 |
|-----------|--------|------|
| header_detection | true | 先頭行をヘッダーとして検出 |
| align_detection | true | 数値列の右寄せ検出 |
| numbers_right_threshold | 0.8 | 右寄せ判定の数値比率閾値 |
| hidden_policy | ignore | 隠れ行・列の処理（ignore / include / exclude） |

**Mermaid関連**

| オプション | 既定値 | 説明 |
|-----------|--------|------|
| mermaid_enabled | false | Mermaid生成の有効化 |
| mermaid_detect_mode | none | 検出モード（none / column_headers / heuristic / shapes） |
| mermaid_diagram_type | flowchart | 図の種類（flowchart / sequence / state） |
| mermaid_direction | TD | フロー方向（TD / LR / BT / RL） |
| mermaid_keep_source_table | true | 元テーブルも出力 |

**CSV Markdown関連**

| オプション | 既定値 | 説明 |
|-----------|--------|------|
| csv_markdown_enabled | true | CSV Markdown形式出力の有効化 |
| csv_output_dir | なし | CSV出力ディレクトリ |
| csv_include_metadata | true | 検証用メタデータの付加 |
| csv_include_description | true | 説明セクションの付加 |

**画像抽出**

| オプション | 既定値 | 説明 |
|-----------|--------|------|
| image_extraction | true | 画像抽出の有効化 |

**その他**

| オプション | 既定値 | 説明 |
|-----------|--------|------|
| footnote_scope | book | 脚注番号のスコープ（book / sheet） |
| split_by_sheet | false | シートごとに分割出力 |

### 3.2 出力

#### 3.2.1 CSV Markdown（既定）

csv_markdown_enabled=true（既定）時の出力形式。

**構成**
1. 概要セクション
   - ファイル情報
   - CSV生成方法の説明
   - CSV形式の仕様説明
   - 検証メタデータの説明
2. CSVデータセクション（シートごと）
   - シート見出し
   - CSVコードブロック
   - 画像リンク（該当する場合）
3. 検証用メタデータ（オプション）

**CSV形式仕様**
- 区切り文字: カンマ（,）
- 引用符: RFC 4180準拠
- セル内改行: スペースに変換（1行=1レコード）
- エンコーディング: UTF-8
- 空セル: 空文字列

#### 3.2.2 通常Markdown

csv_markdown_enabled=false 時の出力形式。

**構成**
1. ドキュメントヘッダー（ファイル名、仕様バージョン、シート情報）
2. シートセクション（シートごとの見出しとテーブル）
3. 脚注セクション（ハイパーリンク等の脚注）

**テーブル表現**
- GFM（GitHub Flavored Markdown）テーブル形式
- ヘッダー行、区切り行、データ行で構成
- 列アラインメント（左寄せ/右寄せ）対応

---

## 4. 処理フロー

### 4.1 全体処理フロー

```
[入力] Excelファイル + オプション
         ↓
    CLI引数解析
         ↓
    ワークブック読み込み
         ↓
    ┌─────────────────┐
    │ シート単位ループ │
    └────────┬────────┘
             ↓
        保護状態チェック
             ↓
    ┌──────────────────────┐
    │ shapes検出モード時    │→ DrawingML図形からMermaid生成
    └──────────────────────┘
             ↓
        印刷領域取得
             ↓
        矩形和集合計算
             ↓
    ┌──────────────────────┐
    │ CSV Markdown有効時    │→ 画像抽出
    └──────────────────────┘
             ↓
    ┌───────────────────────┐
    │ 矩形・テーブル単位ループ │
    └─────────┬─────────────┘
              ↓
         結合セルマップ作成
              ↓
         テーブル分割検出
              ↓
    ┌─────────────────────┐
    │ 各テーブル単位ループ  │
    └────────┬────────────┘
             ↓
        テーブルタイトル検出
             ↓
        テーブル抽出
        ・セル値取得・正規化
        ・ハイパーリンク処理
        ・Markdownエスケープ
        ・脚注生成
             ↓
        テーブル形式判定・出力
        ① コード形式判定
        ② Mermaidフロー判定
        ③ 単一行テキスト判定
        ④ ネスト形式判定
        ⑤ 通常テーブル形式
             ↓
    脚注処理（footnote_scope設定に基づく）
         ↓
[出力] CSV Markdown / 通常Markdown
```

### 4.2 テーブル検出フロー

```
印刷領域/使用範囲
       ↓
  非空グリッド構築
  ・セル値の空性判定
  ・背景色の判定
  ・結合セルの考慮
       ↓
  連結成分検出（BFS）
  ・空行/空列をテーブル境界として認識
       ↓
  矩形分解（Histogram Algorithm）
  ・最大矩形の反復抽出
       ↓
  テーブルリスト生成
  ・各テーブルの矩形リスト
  ・バウンディングボックス
  ・セル座標集合
```

### 4.3 テーブル形式判定フロー

抽出されたテーブルは、以下の優先順位で形式が判定される。

| 優先度 | 形式 | 判定条件 |
|-------|------|---------|
| 1 | コード形式 | コードキーワード（public, class, def等）、コードシンボル（{}, ;, //等）の検出 |
| 2 | Mermaid形式 | mermaid_enabled=true かつ検出モードの条件に合致 |
| 3 | 単一行テキスト | 先頭行に1つの値のみ、枠線なし、他セルすべて空 |
| 4 | ネスト形式 | 各行の最初の非空セルにインデントがある |
| 5 | 通常テーブル | 上記いずれにも該当しない場合 |

---

## 5. セル・テーブル処理規則

### 5.1 セル値取得

セル値は以下の手順で取得・正規化される。

1. **値取得**: value_mode設定に基づき、表示値/数式/両方を取得
2. **日付判定**: 日付型の場合、指定フォーマットで文字列化
3. **Unicode正規化**: NFC形式に正規化
4. **制御文字除去**: 不可視制御文字の除去
5. **空白処理**: strip_whitespace=true時、前後空白を除去
6. **数値正規化**: 通貨記号、千単位区切り、パーセントの処理

### 5.2 制御文字除去対象

| 範囲 | 説明 |
|------|------|
| U+0000-U+0008 | NULL～バックスペース前 |
| U+000B-U+000C | 垂直タブ、フォームフィード |
| U+000E-U+001F | シフトアウト～ユニットセパレータ |
| U+007F | DELETE |
| U+0080-U+009F | C1制御文字 |
| U+00AD | 軟ハイフン |
| U+200B-U+200D | ゼロ幅文字 |
| U+2060 | ワード結合子 |
| U+2066-U+2069 | 方向性制御文字 |

### 5.3 結合セル処理

| ポリシー | 動作 |
|---------|------|
| top_left_only | 左上セルのみ値を保持、他セルは空 |
| expand | 左上セルの値を全セルに展開 |
| repeat | expandと同等 |
| warn | expandしつつ警告を出力 |

### 5.4 セル空性の判定

セルが「空」と判定される条件：
- セル値が空文字列または空白のみ
- かつ、背景色が無色または白色

白色と見なされるフィル：
- RGB: #FFFFFF
- インデックスカラー: 64（白色インデックス）
- 読み取り専用モード時: フィル情報不可のため無色と仮定

---

## 6. Markdown生成規約

### 6.1 テーブル生成

**ヘッダー処理**
- header_detection=true: 先頭行をヘッダーとして扱う
- header_detection=false: ヘッダーなし（全行がデータ行）

**列アラインメント**
- align_detection=true: 各列の数値比率が閾値（既定80%）以上なら右寄せ
- 右寄せ: 区切り行を `---:` で表現
- 左寄せ: 区切り行を `---` で表現

### 6.2 エスケープ処理

**エスケープ対象文字**

Markdownの予約文字: `\` `|` `*` `_` `~` `#` `>` `[` `]` `(` `)` `{` `}` `+` `-` `.` `!` `` ` ``

**エスケープレベル**

| レベル | 動作 |
|-------|------|
| safe（既定） | 上記予約文字にバックスラッシュを付加 |
| minimal | パイプ記号のみエスケープ |
| aggressive | すべての文字をエスケープ（非推奨） |

**改行処理**

セル内の改行（`\r\n`, `\r`, `\n`）は `<br>` タグに変換される。

### 6.3 ハイパーリンク出力形式

| モード | 出力例 | 用途 |
|-------|--------|------|
| inline | `[テキスト](URL)` | 標準的なMarkdownリンク |
| inline_plain | `テキスト (URL)` | 平文でのURL埋め込み |
| footnote | `テキスト[^1]` + 脚注 | 参考文献スタイル |
| both | inline形式 + 脚注 | 両方を出力 |
| text_only | `テキスト` | リンク情報を除去 |

### 6.4 脚注スコープ

| スコープ | 動作 |
|---------|------|
| book | ワークブック全体で連番（[^1], [^2], ...） |
| sheet | シートごとに番号をリセット |

---

## 7. Mermaid生成規約

### 7.1 検出モード

#### 7.1.1 shapes検出モード

DrawingMLの図形（Shape）とコネクタ（Connector）を解析してMermaidコードを生成する。

**検出対象**
- xdr:sp（Shape）: フローチャートの処理ボックス
- xdr:cxnSp（Connector）: ノード間の接続線

**図形タイプマッピング**

| Excel図形 | Mermaidノード形式 | 意味 |
|----------|------------------|------|
| flowChartDecision, diamond | `{"テキスト"}` | 判断 |
| flowChartTerminator, ellipse | `(["テキスト"])` | 開始/終了 |
| flowChartInputOutput, trapezoid | `[("テキスト")]` | 入出力 |
| flowChartPreparation, hexagon | `[{"テキスト"}]` | 準備 |
| flowChartManualOperation | `[("テキスト")]` | 手動操作 |
| flowChartDocument | `[("テキスト")]` | 文書 |
| その他 | `["テキスト"]` | 処理 |

#### 7.1.2 column_headers検出モード

テーブルのヘッダー行を解析し、フロー定義列（From, To, Label等）を特定する。

**列マッピング既定値**
- from列: "From"
- to列: "To"
- label列: "Label"
- group列: なし
- note列: なし

#### 7.1.3 heuristic検出モード

テーブルのパターンを統計的に分析してフローテーブルを検出する。

**判定条件**
1. データ行数が最小行数（既定3行）以上
2. 先頭2列に値がある行が最小行数以上
3. 矢印記号（->, →, ⇒）を含む行の比率が閾値（既定30%）以上
4. 1列目と2列目の値の長さの中央値比率が許容範囲（既定0.4〜2.5）内

### 7.2 Mermaid生成規則

**図の種類**
- flowchart: フローチャート
- sequence: シーケンス図
- state: 状態遷移図

**フロー方向**
- TD: 上から下
- LR: 左から右
- BT: 下から上
- RL: 右から左

**ノードID生成**
- auto: テキストから自動生成（サニタイズ処理）
- shape_id: 図形IDを使用
- explicit: 明示的に指定されたIDを使用

**エッジ重複排除**
mermaid_dedupe_edges=true時、同一のFrom-Toペアは1つのエッジにまとめられる。

**サブグラフ**
group列が指定され、mermaid_group_column_behavior="subgraph"の場合、グループごとにサブグラフとして出力される。

---

## 8. 画像抽出規約

### 8.1 抽出方式

**優先方式: DrawingML抽出**
1. Excelファイルをzipアーカイブとして開く
2. workbook.xmlからシートIDを取得
3. worksheets/_rels/sheet{ID}.xml.relsからDrawing参照を取得
4. drawings/drawing{N}.xmlから画像参照を抽出
5. mediaディレクトリから画像データを取得

**フォールバック: openpyxl抽出**
DrawingML抽出で画像が取得できない場合、openpyxlのws._imagesを使用する。

### 8.2 出力仕様

**出力ディレクトリ**
```
{output_dir}/{ファイル名}_images/
```

**ファイル名形式**
```
{サニタイズ済シート名}_img_{連番}.{拡張子}
```

**対応画像形式**
- PNG
- JPEG
- GIF
- BMP
- その他（magic bytesによる判定）

### 8.3 Markdownでの参照

画像はMarkdownの画像リンク形式で参照される。

- alt text: セル値がある場合はその値、なければ "Image at {セル参照}"
- パス: 出力ディレクトリからの相対パス

---

## 9. エラーハンドリング

### 9.1 基本方針

**処理継続優先**: エラーが発生しても可能な限り処理を継続し、部分的な出力を確保する。

### 9.2 ログ出力

| レベル | 形式 | 出力先 |
|-------|------|--------|
| 警告 | [WARN] メッセージ | 標準エラー出力 |
| 情報 | [INFO] メッセージ | 標準エラー出力 |

### 9.3 例外処理ポリシー

| 例外発生箇所 | 処理 | ポリシー |
|-------------|------|---------|
| Excelファイル読み込み失敗 | 終了コード2で終了 | 即時終了 |
| シート処理エラー | 警告出力、次シートへ | 継続 |
| CSV抽出エラー | 警告出力、次シートへ | 継続 |
| Mermaid生成エラー | 警告出力、通常テーブルにフォールバック | 継続 |
| 画像抽出エラー | 警告出力、次画像へ | 継続 |
| メタデータ付記エラー | 警告出力、メタデータなしで出力 | 継続 |

### 9.4 代表的な警告メッセージ

| 状況 | メッセージ例 |
|------|-------------|
| 無効なURL | Invalid URL detected at C5: invalid://url |
| 印刷領域超過 | Print area max_row (100) exceeds sheet max_row (50), limiting to 50 |
| 画像抽出失敗 | Failed to extract image from media/image1.png: ... |
| 非対応検出モード | mermaid_detect_mode='column_headers' is not supported for CSV markdown |

---

## 10. 用語集

### 10.1 Excel関連

| 用語 | 定義 |
|------|------|
| 印刷領域 | ユーザーが定義した印刷対象範囲。未設定の場合はno_print_area_modeに従う |
| 使用範囲 | 値を持つセルの最小バウンディングボックス |
| 結合セル | 複数セルをマージして1つのセルとして扱うExcelの機能 |
| 結合トップレフト | 結合セル範囲の左上セル。値はここにのみ格納される |
| DrawingML | Office Open XMLで図形や画像を表現するためのマークアップ言語 |

### 10.2 テーブル検出関連

| 用語 | 定義 |
|------|------|
| 非空セル | 値があるか、または背景色が設定されているセル |
| 連結成分 | 空行・空列で分断されていない、隣接するセルの集合 |
| 矩形 | 連続した行・列で構成される矩形領域。テーブルの基本単位 |
| バウンディングボックス | テーブルを囲む最小の矩形 |

### 10.3 Mermaid関連

| 用語 | 定義 |
|------|------|
| ノード | フローチャートの処理ボックス。図形タイプにより表現が異なる |
| エッジ | ノード間の接続を表す矢印 |
| サブグラフ | 複数のノードをグループ化した領域 |

### 10.4 出力形式関連

| 用語 | 定義 |
|------|------|
| CSV Markdown | CSVコードブロックと検証用メタデータを含むMarkdown出力（既定の出力形式） |
| 通常Markdown | GFMテーブル形式を基本とした標準的なMarkdown出力 |
| 脚注 | ドキュメント末尾にまとめて配置される参照情報 |
| インラインリンク | テキスト中に埋め込まれた形式のハイパーリンク |

---

## 11. テスト仕様

### 11.1 概要

本章は、excel2mdパッケージの単体テスト仕様を定義する。テストはpytestを使用し、openpyxlのモックまたはメモリ内Workbookを活用して、実際のExcelファイルなしでテスト可能とする。

#### 11.1.1 テストファイル構成

```
excel2md/
├── __init__.py
├── cli.py
├── runner.py
├── ...（その他モジュール）
└── tests/
    ├── __init__.py
    ├── conftest.py              # pytest fixtures
    ├── test_cell_utils.py       # セル処理関連
    ├── test_table_detection.py  # テーブル検出関連
    ├── test_markdown_output.py  # Markdown出力関連
    ├── test_csv_markdown.py     # CSVマークダウン出力関連
    ├── test_mermaid.py          # Mermaid変換関連
    ├── test_hyperlink.py        # ハイパーリンク処理関連
    └── test_integration.py      # 統合テスト
```

#### 11.1.2 テストの原則

1. **モック優先**: openpyxlのWorkbook/Worksheetオブジェクトをメモリ内で生成
2. **独立性**: 各テストは他のテストに依存しない
3. **再現性**: ファイルシステムや外部リソースに依存しない
4. **網羅性**: 正常系・異常系・境界値を網羅

#### 11.1.3 テスト実行方法

```bash
cd v2.0
uv run pytest tests/ -v                    # 全テスト実行
uv run pytest tests/test_cell_utils.py -v  # 特定ファイルのみ
uv run pytest tests/ --cov=. --cov-report=html  # カバレッジ付き
```

### 11.2 テスト方針

#### 11.2.1 セル処理（test_cell_utils.py）

- **空性判定**: 値なし、空文字、空白のみ、数値0、背景色の有無を組み合わせて検証
- **数値判定**: 整数、負数、桁区切り、通貨記号、パーセント、括弧表記の正常系と、不正形式の異常系を網羅
- **エスケープ**: 各エスケープレベル（safe/minimal/aggressive）での変換結果を検証

#### 11.2.2 テーブル検出（test_table_detection.py）

- **グリッド構築**: 空行・空列によるテーブル分割、結合セルの影響を検証
- **連結成分**: 矩形、L字型、離散領域など多様な形状でのテーブル検出を検証
- **矩形分解**: 非矩形領域の長方形分解が正しく行われることを検証

#### 11.2.3 Markdown出力（test_markdown_output.py）

- **テーブル生成**: ヘッダー検出、列アラインメント、空列除去を検証
- **テーブル抽出**: 結合セル処理、ハイパーリンク処理、脚注生成を検証

#### 11.2.4 CSV出力（test_csv_markdown.py）

- **データ抽出**: 結合セル、改行、ハイパーリンクを含むセルの処理を検証
- **オプション**: 各オプション（description、metadata、mermaid）の有効/無効時の出力を検証

#### 11.2.5 ハイパーリンク（test_hyperlink.py）

- **リンク種別**: 外部URL、内部参照、mailto等の各種リンク形式を検証
- **URL検証**: 有効なURL形式と無効な形式（javascript:等）の判定を検証

#### 11.2.6 Mermaid（test_mermaid.py）

- **コード判定**: プログラミング言語のキーワード・記号検出を検証
- **言語検出**: 各言語固有のパターンによる言語判定を検証
- **出力形式判定**: テキスト/ネスト/コード/テーブルの形式判定を検証

### 11.3 動作テスト項目

本節では、自動テストでカバーしにくい項目について、動作確認する試験項目を定義する。

#### 必要なテスト用Excelファイル

以下のExcelファイルを `v2.0/tests/fixtures/` ディレクトリに準備する。

**test_standard.xlsx** - 標準テスト用（複数シート構成）
- Sheet1「基本テーブル」: 通常のテーブル（ヘッダー行あり、数値・文字列混在）
- Sheet2「結合セル」: 結合セルを含むテーブル
- Sheet3「複数テーブル」: 空行で分割された複数テーブル、印刷領域設定あり
- Sheet4「ハイパーリンク」: 外部URL、内部参照（シート内・他シート）、mailto
- Sheet5「画像」: PNG画像、JPEG画像を複数配置

**test_mermaid.xlsx** - Mermaidテスト用
- Sheet1「図形フロー」: フローチャート図形（矩形、ひし形、楕円、コネクタ）
- Sheet2「テーブルフロー」: From/To/Label列を持つフロー定義テーブル

#### 11.3.1 CLI基本動作確認

```bash
uv run python v2.0/excel_to_md.py --help
uv run python v2.0/excel_to_md.py nonexistent.xlsx
```

- [ ] `--help` で全オプションの説明が表示される
- [ ] 存在しないファイル指定時、エラーメッセージと終了コード2で終了する

#### 11.3.2 test_standard.xlsx を使用した確認

**基本変換**
```bash
uv run python v2.0/excel_to_md.py test_standard.xlsx
```
- [ ] エラーなく `test_standard_csv.md` が生成される
- [ ] 5シート分のCSVコードブロックが出力される
- [ ] 概要セクションとメタデータセクションが含まれる

**出力形式オプション**
```bash
uv run python v2.0/excel_to_md.py test_standard.xlsx --no-csv-markdown-enabled -o out.md
uv run python v2.0/excel_to_md.py test_standard.xlsx --split-by-sheet
uv run python v2.0/excel_to_md.py test_standard.xlsx --csv-output-dir ./output
```
- [ ] `--no-csv-markdown-enabled -o out.md` でGFMテーブル形式のMarkdownが出力される
- [ ] `--split-by-sheet` でシートごとに個別ファイル（5ファイル）が生成される
- [ ] `--csv-output-dir ./output` で指定ディレクトリに出力される

**Sheet1「基本テーブル」の確認**
- [ ] ヘッダー行が正しく認識される
- [ ] 数値列が右寄せで出力される

**Sheet2「結合セル」の確認**
- [ ] 結合セルの値が左上セルに出力される
- [ ] 結合範囲の他セルは空で出力される

**Sheet3「複数テーブル」の確認**
- [ ] 空行を境界として複数テーブルが検出される
- [ ] 印刷領域内のみが変換対象となる

**Sheet4「ハイパーリンク」の確認**
```bash
uv run python v2.0/excel_to_md.py test_standard.xlsx --hyperlink-mode inline
uv run python v2.0/excel_to_md.py test_standard.xlsx --hyperlink-mode inline_plain
uv run python v2.0/excel_to_md.py test_standard.xlsx --hyperlink-mode footnote
```
- [ ] `inline` モードで `[テキスト](URL)` 形式で出力される
- [ ] `inline_plain` モードで `テキスト (URL)` 形式で出力される
- [ ] `footnote` モードで脚注形式 `テキスト[^1]` で出力される
- [ ] 内部リンク（シート参照）が `→シート名!セル` 形式で出力される
- [ ] mailtoリンクが正しく処理される

**Sheet5「画像」の確認**
- [ ] `test_standard_images/` ディレクトリが作成される
- [ ] PNG画像が `Sheet5_img_1.png` として抽出される
- [ ] JPEG画像が `Sheet5_img_2.jpg` として抽出される
- [ ] CSVに `![alt](test_standard_images/...)` 形式でリンクが出力される

```bash
uv run python v2.0/excel_to_md.py test_standard.xlsx --no-image-extraction
```
- [ ] `--no-image-extraction` で画像が抽出されない

#### 11.3.3 test_mermaid.xlsx を使用した確認

**Sheet1「図形フロー」の確認**
```bash
uv run python v2.0/excel_to_md.py test_mermaid.xlsx --mermaid-enabled --mermaid-detect-mode shapes
uv run python v2.0/excel_to_md.py test_mermaid.xlsx --mermaid-enabled --mermaid-detect-mode shapes --mermaid-direction LR
```
- [ ] フローチャート図形からMermaidコードが生成される
- [ ] 矩形が `[テキスト]`、ひし形が `{テキスト}`、楕円が `([テキスト])` で出力される
- [ ] コネクタが `-->` で接続される
- [ ] `--mermaid-direction LR` で `flowchart LR` が出力される

**Sheet2「テーブルフロー」の確認**
```bash
uv run python v2.0/excel_to_md.py test_mermaid.xlsx --mermaid-enabled --mermaid-detect-mode column_headers
uv run python v2.0/excel_to_md.py test_mermaid.xlsx --mermaid-enabled --mermaid-detect-mode column_headers --no-mermaid-keep-source-table
```
- [ ] From/To/Label列からMermaidコードが生成される
- [ ] `--no-mermaid-keep-source-table` で元テーブルが出力されない

