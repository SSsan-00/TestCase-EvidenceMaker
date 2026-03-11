# VBAマクロ ガイド（UserForm連携対応）
対象モジュール:
- `BetaEvidenceGenerator.bas`
- `BetaTestCaseGenerator.bas`
- `ConditionalBranchChecker.bas`
- `EscapePartsMarking.bas`
- `MacroUserFormBridge.bas`
- `MacroToolsUserFormInstaller.bas`

---

## 共通の使い方
1. VBAエディタで各 `.bas` を標準モジュールとしてインポートします。
2. 直接実行する場合は従来どおり `RunMain` を呼びます。
3. UserForm から実行する場合は、`RunMainWithUiOptions` を使います。
4. UserForm 側のコードを短くしたい場合は `MacroUserFormBridge.bas` のラッパーSubを使います。
5. 自動で UserForm を作る場合は `InstallMacroToolsUserForm` を実行します。

---

## UserForm連携の基本

### 方針
- 各モジュールに `Public Type ...UiOptions` を用意しています。
- 各モジュールにフォーム初期値用の公開API（`Create...UiOptionsForForm()` または `Initialize...UiOptionsForForm(...)`）を用意しています。
- フォーム初期値を公開APIで取得/初期化し、入力値で上書きして `RunMainWithUiOptions` へ渡します。

### 重要な挙動
- `MacroUserFormBridge.bas` では主要必須項目を事前チェックします。
- 必須チェックの内容:
  - `RunBetaEvidenceFromForm`: 参照元ブックパス / 入力ファイル名
  - `RunBetaTestCaseFromForm`: 機能連番
  - `RunConditionalBranchCheckerFromForm`: 機能名 / 対象ブックパス
  - `RunEscapePartsMarkingFromForm`: 対象ブックパス
- `frmMacroTools` 生成コードでは、テストケース生成を `UseOutputPath=False` で実行します（フォームに出力パス入力欄はありません）。

### UserFormボタンからの呼び出し例（推奨: ブリッジ経由）
```vb
Private Sub btnRunEvidence_Click()
    MacroUserFormBridge.RunBetaEvidenceFromForm _
        sourceWorkbookPath:=Trim$(Me.txtSourcePath.Value), _
        inputFileName:=Trim$(Me.txtInputFileName.Value), _
        useSlotHeight:=True, _
        slotHeight:=CLng(Me.txtSlotHeight.Value), _
        useOutputSheetFilter:=True, _
        outputSheetFilterText:=Trim$(Me.txtOutputFilter.Value), _
        topBorderEnabled:=CBool(Me.chkTopBorder.Value), _
        excludeOutputSheetByPatternEnabled:=CBool(Me.chkExcludePattern.Value), _
        excludedOutputSheetNamePatterns:=Trim$(Me.txtExcludePatterns.Value), _
        skipGrayFilledSourceCellEnabled:=CBool(Me.chkSkipGray.Value), _
        sourceSkipFillColorHexCodes:=Trim$(Me.txtSkipColors.Value), _
        rightBorderEnabled:=CBool(Me.chkRightBorder.Value), _
        useRightBorderTargetCol:=CBool(Me.chkUseRightBorderCol.Value), _
        rightBorderTargetCol:=ReadColumnIndexOrDefault(CStr(Me.txtRightBorderCol.Value), 17)
End Sub
```

---

## 1. `BetaEvidenceGenerator.bas`

### 概要
- 参照元 `.xlsx` を読み取り、共通/個別エビデンス `.xlsx` を新規作成します。
- `REFER` シートを使って `referValue` と出力名要素を解決します。
- 共通ヘッダ `A1-1-1` と本体テンプレ `A1` を使って出力します。

### UserForm向け公開API
- `Public Type BetaEvidenceUiOptions`
- `Public Sub InitializeBetaEvidenceUiOptionsForForm(ByRef options As BetaEvidenceUiOptions)`
- `Public Sub RunMainWithUiOptions(ByRef options As BetaEvidenceUiOptions)`

### 主なフォーム入力項目
- 参照元ブックパス
- 入力ファイル名
- スロット行オフセット
- 出力対象シートフィルタ（カンマ区切り）
- 上罫線ON/OFF
- 除外パターンON/OFF + パターン文字列
- 灰色セル読み飛ばしON/OFF + 対象色
- 右罫線ON/OFF + 右罫線列名（例: `Q`）

### 主要オプション（定数）
- `OPTION_TOP_BORDER_ENABLED`
- `OPTION_SLOT_HEIGHT_PROMPT_ENABLED`
- `OPTION_OUTPUT_SHEET_SELECTION_PROMPT_ENABLED`
- `OPTION_EXCLUDE_OUTPUT_SHEET_BY_PATTERN_ENABLED`
- `EXCLUDED_OUTPUT_SHEET_NAME_PATTERNS`
- `OPTION_SKIP_GRAY_FILLED_SOURCE_CELL_ENABLED`
- `SOURCE_SKIP_FILL_COLOR_HEX_CODES`
- `OPTION_RIGHT_BORDER_ENABLED`
- `RIGHT_BORDER_TARGET_COL`（既定: `17` = `Q` 列）

### 補足
- `frmMacroTools` では右罫線列入力の既定値は `Q` です。
- `frmMacroTools` 生成コードの `ReadColumnIndexOrDefault` で列名を列番号へ変換して渡します（互換のため数値入力も受け付け）。

---

## 2. `BetaTestCaseGenerator.bas`

### 概要
- 機能連番をキーに `REFER` からヒット行を抽出し、テストケース用 `.xlsx` を作成します。

### UserForm向け公開API
- `Public Type BetaTestCaseUiOptions`
- `Public Function CreateBetaTestCaseUiOptionsForForm() As BetaTestCaseUiOptions`
- `Public Sub RunMainWithUiOptions(ByRef options As BetaTestCaseUiOptions)`

### 主要フォーム入力
- 機能連番（必須）

### 補足
- API上は `UseOutputPath` / `OutputPath` を受け付けます。
- `frmMacroTools` ではこの入力欄を持たず、常に `UseOutputPath=False` で実行します。

---

## 3. `ConditionalBranchChecker.bas`

### 概要
- 現行ソースシートの条件分岐構文を解析し、B列マーキングと個別シート出力を行います。

### UserForm向け公開API
- `Public Type ConditionalBranchCheckerUiOptions`
- `Public Function CreateConditionalBranchCheckerUiOptionsForForm() As ConditionalBranchCheckerUiOptions`
- `Public Sub RunMainWithUiOptions(ByRef options As ConditionalBranchCheckerUiOptions)`

### 主要フォーム入力
- 機能名（必須）
- 対象ブックパス（必須）
- `LEADING_FUNCTION_STARTS_FROM_B1` のON/OFF

### 構文対象
- `if / else if / elseif / else`
- 三項演算子
- `for / foreach / while`
- `switch / case / default`

---

## 4. `EscapePartsMarking.bas`

### 概要
- `A1-1-1` を含むシートを対象に、B列のエスケープ関数呼び出しを赤字太字でマーキングします。
- ヒット行C列へ完了メッセージを書き込みます。
- A列のみ入力行の塗りつぶしルールを適用できます。

### UserForm向け公開API
- `Public Type EscapePartsMarkingUiOptions`
- `Public Function CreateEscapePartsMarkingUiOptionsForForm() As EscapePartsMarkingUiOptions`
- `Public Sub RunMainWithUiOptions(ByRef options As EscapePartsMarkingUiOptions)`

### 主要フォーム入力
- 対象ブックパス（必須）
- 完了メッセージ
- エスケープ関数一覧（CSV）
- 塗りつぶし対象
  - `None`
  - `Left`
  - `Right`
  - `Both`
- 塗りつぶし色HEX（例: `#a6a6a6`）

---

## 5. `MacroUserFormBridge.bas`

### 目的
- UserForm側のイベントコードを短くするための呼び出しラッパーです。

### 提供Sub
- `RunBetaEvidenceFromForm(...)`
- `RunBetaTestCaseFromForm(...)`
- `RunConditionalBranchCheckerFromForm(...)`
- `RunEscapePartsMarkingFromForm(...)`

### 提供Function
- `GetEscapeOnlyAValueRowFillTargetOptions()`
  - `Array("None", "Left", "Right", "Both")` を返します。

---

## 6. `MacroToolsUserFormInstaller.bas`

### 目的
- UserForm (`frmMacroTools`) と起動用標準モジュール (`modMacroToolsFormEntry`) を自動生成します。
- 生成フォームに4機能（Evidence/TestCase/Conditional/Escape）の入力UIをまとめます。

### 生成フォーム仕様（現行）
- 画面に収まらない高さは縦スクロールで操作します。
- 右上の `×` と、フォーム下部の `閉じる` ボタンのどちらでも終了できます。
- 参照元/対象ブックの `...` ボタンは `GetOpenFilename` でファイル選択します。
- エビデンス生成の右罫線列は列名入力（既定 `Q`）です。

### 実行手順
1. `MacroToolsUserFormInstaller.bas` をインポート
2. `InstallMacroToolsUserForm` を実行
3. `OpenMacroToolsForm` を実行してフォームを表示

### 前提
- Excel の設定で以下を有効化
  - `トラスト センター > マクロの設定 > VBA プロジェクト オブジェクト モデルへのアクセスを信頼する`

### 再生成
- `InstallMacroToolsUserForm` は既存の `frmMacroTools` / `modMacroToolsFormEntry` を削除して再作成します。

---

## 運用メモ
- 直接編集で既定値を変える場合は各モジュール先頭の `Const` を変更してください。
- UserFormで値を保持したい場合は、フォーム初期化時に各公開API（`Create...UiOptionsForForm()` / `Initialize...UiOptionsForForm(...)`）を使ってコントロール初期値へ反映する運用が安全です。
- 直接更新系（`ConditionalBranchChecker` / `EscapePartsMarking`）は実行前バックアップを推奨します。
