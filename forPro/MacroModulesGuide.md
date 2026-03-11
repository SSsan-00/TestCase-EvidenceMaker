# VBAマクロ4モジュール ガイド
対象モジュール:
- `BetaEvidenceGenerator.bas`
- `BetaTestCaseGenerator.bas`
- `ConditionalBranchChecker.bas`
- `EscapePartsMarking.bas`

---

## 共通の使い方
1. VBAエディタで各 `.bas` を標準モジュールとしてインポートします。
2. 実行は `RunMain` を呼び出します（例: `Call BetaEvidenceGenerator.RunMain`）。
3. 直接更新するマクロもあるため、実行前にバックアップを推奨します。

---

## 1. `BetaEvidenceGenerator.bas`

### 概要
- 参照元 `.xlsx` を読み取り、共通/個別のエビデンス `.xlsx` を新規作成します。
- `REFER` シートから入力ファイル名をキーに `referValue` と出力ファイル名要素（α/β/γ）を取得します。
- 参照元シートは以下を対象にします。
  - 共通: `【共通】` + `referValue`
  - 個別: `【個別】` + `referValue`
- 出力シートは `A1` テンプレを複製して作成します。
- 共通ヘッダー `A1-1-1` は別処理で作成し、`A3/B3` の `〇〇〇` を `baseName` に置換します。

### 主な入力
- 参照元ファイル選択（`.xlsx`）
- 入力ファイル名（例: `menu/mainmenu.php`）
- スロット行オフセット（オプションでダイアログON/OFF）
- 出力対象シート名（カンマ区切り、オプションでダイアログON/OFF）

### 走査・書き込みルール
- 参照元の `A/E/H` 列を `SOURCE_START_ROW=8` から走査します。
- `A` に値が来たら出力先シートを切り替えます。
- `E/H` は現在の出力先シートが確定している場合のみ処理します。
- スロット書き込み先は以下です。
  - `E/H` ペア確定: 出力先 `A/B`
  - `H` 単体: 出力先 `B`
  - `E` 単体確定: 出力先 `A`
- 書き込み開始行は `FIRST_DEST_ROW=3`、既定スロット間隔は `SLOT_HEIGHT=50`。

### 罫線ルール
- 確定書き込み行で、`A:AA` の上罫線を引きます。
- `3` 行目（`FIRST_DEST_ROW`）は上罫線を引きません。
- `A/B` のどちらかに値がある行のみ対象です。
- A1複製シートでは、右罫線設定が有効な場合に、最終書き込み行の `+50` 行まで指定列の右側に罫線を引きます。

### 除外・スキップ
- 出力シート名は `Like` パターンで除外できます。
- 参照元の `A/E/H` セルで、指定塗りつぶし色のセルは「未入力扱い」で読み飛ばします。

### 主要オプション（ソース書き換えで切替）
- `OPTION_TOP_BORDER_ENABLED`:
  - `True`: 上罫線適用
  - `False`: 上罫線無効
- `OPTION_SLOT_HEIGHT_PROMPT_ENABLED`:
  - `True`: オフセット入力ダイアログ表示
  - `False`: `SLOT_HEIGHT` を使用
- `OPTION_OUTPUT_SHEET_SELECTION_PROMPT_ENABLED`:
  - `True`: 出力シート選択ダイアログ表示
  - `False`: 全シート出力
- `OPTION_EXCLUDE_OUTPUT_SHEET_BY_PATTERN_ENABLED`:
  - `True`: 除外パターン有効
  - `False`: 除外なし
- `EXCLUDED_OUTPUT_SHEET_NAME_PATTERNS`:
  - 既定: `A4,A5,A1-1,A2-3-1`
  - カンマ区切りの `Like` パターン
- `OPTION_SKIP_GRAY_FILLED_SOURCE_CELL_ENABLED`:
  - `True`: 指定色セルを読み飛ばす
  - `False`: 色判定なし
- `SOURCE_SKIP_FILL_COLOR_HEX_CODES`:
  - 既定: `#f2f2f2,#d9d9d9,#bfbfbf,#a6a6a6,#808080`
- `OPTION_RIGHT_BORDER_ENABLED`:
  - `True`: 右罫線処理を実行
  - `False`: 右罫線処理を実行しない
- `RIGHT_BORDER_TARGET_COL`:
  - 右罫線を引く列番号（例: `26` = `Z`列）

### 出力ファイル名
- 共通: `<α>_【共通】<β><γ>_単体テストエビデンス_初期開発.xlsx`
- 個別: `<α>_【個別】<β><γ>_単体テストエビデンス_初期開発.xlsx`
- 同名がある場合は `_001`, `_002` を付与して回避します。

---

## 2. `BetaTestCaseGenerator.bas`

### 概要
- 機能連番を入力し、`REFER` を検索してテストケース用 `.xlsx` を新規作成します。
- `REFER` の `J` 列（α）で「機能連番を含む行」をヒット対象にします（部分一致）。
- ヒット行の α/β/γ を使ってシートを複製・埋め込みします。

### `REFER` 列定義
- α: `J` 列（拡張子除去）
- β: `F` 列
- γ: `E` 列

### 作成シート
- ヒット件数ぶん作成:
  - `【共通】<β>`
  - `【個別】<β>`
  - `現行ソース（PHP）<β>`
- 1枚のみ作成:
  - `⇒参考`
  - `現行画面`

### セル埋め
- `【共通】<β>` / `【個別】<β>`:
  - `BD1 = β`
  - `BD3 = 入力した機能連番`
- `現行ソース（PHP）<β>`:
  - `C4 = γ`

### 出力ファイル
- `<α>_単体テストケース_初期開発.xlsx`
- 同名時は `_001`, `_002` を付与
- `ThisWorkbook` が保存済みなら同フォルダ、未保存なら SaveAs ダイアログ

### テンプレ前提
- `【共通】機能名`
- `【個別】機能名`
- `⇒参考`
- `現行ソース（PHP）`
- `現行画面`
- `REFER`

---

## 3. `ConditionalBranchChecker.bas`

### 概要
- 対象 `.xlsx` の「現行ソース」シートを解析し、B列マーキングと個別シート出力を行います。
- 対象ブックは直接更新して保存します（別名出力ではありません）。

### 対象シート判定
- 現行ソースシート:
  - シート名に `現行ソース` を含む候補を収集
  - 候補1件なら採用
  - 候補複数なら、入力機能名と完全一致するものを採用
- 個別シート:
  - `【個別】` + 機能名（完全一致）

### 現行ソースB列マーキング
- 解析列: C列
- コメント行（先頭 `#` / `//`）は除外
- `<!doctype` / `<html` / `<head` を検出したら以降は解析停止
- `function` 行: `Bn`
- それ以外の対象構文（if/else if/elseif/else/三項/for/foreach/while/switch/case/default）: `Bn-`

### 個別シート出力
- 構文イベント（FUNCTION, SWITCH, IF, TERNARY, FOR, FOREACH, WHILE）を抽出して所定セルへ記入
- `switch` は case/default を収集して分岐行を出力
- 行不足時は、個別シート `15` 行目から `10` 行テンプレを退避し、必要に応じて `50` 行単位で挿入

### オプション
- `LEADING_FUNCTION_STARTS_FROM_B1`:
  - `True`: 先頭の判定対象が function のとき `B1` 開始
  - `False`: 従来どおり `B2` 開始

---

## 4. `EscapePartsMarking.bas`

### 概要
- 選択したExcelの、シート名に `A1-1-1` を含むシートだけを処理します。
- B列の `prefix(...)` 部分を赤字・太字にします（4行目以降）。
- ヒット行のC列に固定メッセージ（既定: `SQLインジェクション対策済み`）を赤字で設定します。
- 対象ブックを上書き保存します。

### 追加されたA列のみ入力行の塗りつぶし
- 条件:
  - 行番号 `4` 以上
  - A列に値あり
  - B列が空
- 処理:
  - `OPTION_ONLY_A_VALUE_ROW_FILL_TARGET` に応じて塗りつぶし列を切り替え
  - `None`: 塗りつぶししない
  - `Left`: `A` 列のみ
  - `Right`: `B` 列のみ
  - `Both`: `A/B` 列

### オプション・色設定
- `OPTION_ONLY_A_VALUE_ROW_FILL_TARGET`:
  - `None` / `Left` / `Right` / `Both`
  - 既定: `Both`
- `ONLY_A_VALUE_ROW_FILL_COLOR_HEX`:
  - 既定: `#a6a6a6`
  - `#RRGGBB` で変更可能

### prefix設定
- モジュール先頭の `ESCAPE_TARGET_PREFIXES_CSV` を編集して追加・変更します（カンマ区切り）。
- 既定: `sqlS,sqlN`

---

## 変更時の確認ポイント
- テンプレシート名や `REFER` 列定義を変更する場合は、同名定数を優先して更新してください。
- すべてのオプションはソース内 `Const` で管理しています。運用ルールに合わせて `True/False` または文字列定数を編集してください。
