# VBAマクロ導入・運用ガイド

このガイドは、単一ファイル運用を前提にした導入手順と、各マクロの使い方をまとめたものです。

## 対象ファイル

通常配布・導入で使うファイルは次の 1 本です。

- `MacroSuiteSingleFileInstaller.bas`

このファイルをインポートして `InstallMacroSuiteFromSingleFile` を実行すると、VBA プロジェクト内に次の 5 モジュールが展開されます。

- `BetaEvidenceGenerator.bas`
- `BetaTestCaseGenerator.bas`
- `ConditionalBranchChecker.bas`
- `EscapePartsMarking.bas`
- `MacroToolsUserFormInstaller.bas`

## 初回導入手順

1. 利用先の `.xlsm` を Excel で開きます。
2. `Alt + F11` で VBE を開きます。
3. `ファイル > ファイルのインポート` から `MacroSuiteSingleFileInstaller.bas` をインポートします。
4. Excel の設定で `VBA プロジェクト オブジェクト モデルへのアクセスを信頼する` を有効にします。
5. VBE のマクロ一覧から `InstallMacroSuiteFromSingleFile` を実行します。
6. フォームを使う場合は `InstallMacroToolsUserForm` を実行します。
7. フォームを開く場合は `OpenMacroToolsForm` を実行します。
8. CONFIG シートを使う場合は `CreateMacroConfigSheet` を実行します。
9. VBE の `Debug > Compile VBAProject` でコンパイル確認します。

信頼設定の場所:
`ファイル > オプション > トラスト センター > トラスト センターの設定 > マクロの設定 > VBA プロジェクト オブジェクト モデルへのアクセスを信頼する`

## 更新手順

`MacroSuiteSingleFileInstaller.bas` を差し替える場合は、次の順で更新します。

1. VBE で既存の `MacroSuiteSingleFileInstaller` モジュールを削除します。
2. 新しい `MacroSuiteSingleFileInstaller.bas` をインポートします。
3. `InstallMacroSuiteFromSingleFile` を実行します。
4. フォームを使う場合は `InstallMacroToolsUserForm` を再実行します。
5. CONFIG シートを使う場合は `CreateMacroConfigSheet` を再実行します。
6. `Debug > Compile VBAProject` を実行します。

`InstallMacroSuiteFromSingleFile` は展開先の同名モジュールを置き換えます。既存の UserForm や CONFIG シートは自動更新されないため、必要に応じて再生成してください。

## 単一ファイルの再生成

開発側で各モジュールを編集した後、単一ファイルを作り直す場合は `RegenerateMacroSuiteSingleFileInstaller` を実行します。

- 出力先は `ThisWorkbook.Path\MacroSuiteSingleFileInstaller.bas` です。
- 現在の VBA プロジェクト内にある 5 モジュールを元に、埋め込み済みの単一ファイルを再生成します。

## 利用方法の選択

利用方法は 3 通りあります。

- UserForm: `OpenMacroToolsForm` から画面操作で実行します。
- CONFIG シート: `CreateMacroConfigSheet` で作成した設定シートから実行します。
- 直接実行: 展開後の各モジュールの `RunMain` を直接実行します。

通常運用では UserForm または CONFIG シートを推奨します。直接実行は、最小限の入力だけを求めるプロ向けの運用です。

## UserForm の使い方

1. `InstallMacroToolsUserForm` を実行します。
2. `OpenMacroToolsForm` を実行します。
3. 必要な入力欄を埋めます。
4. 各処理の実行ボタンを押します。

フォームはモデルレス表示です。フォームを開いたまま、同じ Excel インスタンス内のブックやシートを操作できます。

## CONFIG シートの使い方

1. `CreateMacroConfigSheet` を実行します。
2. 作成された `CONFIG` シートの B 列に設定値を入力します。
3. シート上の実行ボタン、または `Run...FromConfigSheet` マクロを実行します。

CONFIG シートの基本構成:

- A列: 項目名
- B列: 入力値
- C列: 説明
- E列以降: 実行ボタン、参照ボタン

CONFIG シートでは、ウィンドウ枠の固定は行いません。`CONFIG を更新` ボタンも置きません。

## CONFIG シートの主な入力仕様

- `C1` は空欄です。
- `参照元ブックパス` / `対象ブックパス` の説明は `必須。` です。
- `行オフセット` は `1以上の整数。` です。
- `新側列数` は `3以上の整数。` です。
- `出力シート絞り込み` は `空欄で全て出力。カンマ(,)で個別選択、コロン(:)で範囲選択。` です。
- `出力シート絞り込み` / `除外パターン` / `読み飛ばし色` / `エスケープ関数一覧` はカンマ区切りで直接入力します。
- `読み飛ばし色` の既定値は `#f2f2f2,#d9d9d9,#bfbfbf,#a6a6a6,#808080` です。
- `グレーアウト対象` は `なし / A列のみ / B列のみ / A,B列` から選択します。
- 色入力は `#RRGGBB / 0xRRGGBB / RRGGBB` 形式のみ許可します。

## 1. BetaEvidenceGenerator

### 目的

参照元 `.xlsx` を読み取り、共通/個別エビデンス `.xlsx` を出力します。

### 主な入力

- 参照元ブックパス
- 対象ファイル名
- 行オフセット
- 新側列数
- 出力シート絞り込み
- 横罫線 ON/OFF
- 縦罫線 ON/OFF
- 現行ラベルを削除する ON/OFF
- 現行側の列数も追従する ON/OFF
- 除外パターン ON/OFF
- 除外パターン
- グレーで塗りつぶしたセルを読み飛ばす ON/OFF
- 読み飛ばし色一覧

### 出力ブックの扱い

- 想定ファイル名のブックが存在しない場合は新規作成します。
- 想定ファイル名のブックが存在し、今回作成予定のシートと衝突しなければ、その既存ブックへ追加します。
- 想定ファイル名のブックに同名シートが 1 つでもある場合は、連番付きの新規ブックへそのモード全体を出力します。
- 既存ブックが開かれている場合は、その開いているブックへ追加して保存します。自動では閉じません。
- 保存後に開いたときのアクティブシートは先頭シートです。

### 出力シート絞り込み

- `A1,A2,A3` のようなカンマ区切り指定に対応します。
- `A1:A3` のような範囲指定に対応します。
- `:A2` は `A1,A2` として扱います。
- `A3:` は参照元に存在する最大番号まで展開します。
- `A1:A3,B1` のように、範囲指定と単独指定を併用できます。
- `A1:B3` のように、1 つの範囲指定内で共通/個別をまたぐ指定はエラーです。

### 除外パターン

`A2-3-1` のように指定した場合、`A2-3-1` だけを除外します。`A2` 本体は除外しません。

`Like` 判定を使うため、`B3-*` のようなワイルドカード指定も可能です。

### 罫線と列構成

- 横罫線と縦罫線の位置は、列数変更後の構成に追従します。
- 縦罫線は、最後の書き込み行に対して行オフセット分だけ延長します。
- 横罫線 ON の場合、最後の縦罫線終端行に下罫線を引いて閉じます。
- 縦罫線を引き直す前に、雛形由来の余分な縦罫線を UsedRange の最終行まで消します。
- 列構成変更の対象は `A1` 複製シートです。`A1-1-1` は変更しません。
- 共通モード先頭シートでは、`A1-1-1` の `A3/B3` にある `○○○` を `baseName` へ置換します。

### 直接実行

`RunMain` を直接実行した場合は、参照元ブックの選択と対象ファイル名の入力を求めます。行オフセットや出力シート絞り込みの直接入力ダイアログは既定で OFF です。

## 2. BetaTestCaseGenerator

### 目的

テストケースブックを新規作成します。保存後に開いたときのアクティブシートは先頭シートです。

### 入力仕様

- 機能連番は `S99-999-99` 形式のみ許可します。
- `S` は大文字固定です。
- 数字部は半角 `0-9` のみ許可します。
- 形式不正時はエラーで中断します。

### 直接実行

`RunMain` を直接実行した場合は、機能連番の入力を求めます。

## 3. ConditionalBranchChecker

### 目的

対象ブックを走査し、条件分岐の一覧化、個別シートへの出力、現行ソースシートへのマーキングを行います。

### 主な入力

- 機能名
- 対象ブックパス
- 先頭FunctionをB1開始にする ON/OFF
- 非Function行に `B1-` 形式を書き込む ON/OFF
- マーキングセルを塗りつぶす ON/OFF
- 塗りつぶし色

### 現行ソースシートのマーキング

- `function` 宣言行は A 列へ `★` を書き込みます。
- `function` 宣言行は B 列へ `B1` / `B2` ... を書き込みます。
- A 列に既存値がある場合も、`function` 宣言行では `★` で上書きします。
- 前回実行で残った A 列の `★` は、今回の再判定前に消します。
- 非Function行は、オプション ON 時のみ B 列へ `B1-` 形式を書き込みます。
- マーキングセルの塗りつぶしは B 列セルに適用します。
- 再度ブックを開いたときのアクティブシートは、走査した現行ソースシートです。

### 個別シート出力

- `switch` 本体を書いた直後に、`case/default` 内の `if / for / while / switch` などをソース出現順で続けて書き込みます。
- ネストした `switch` も、親 `switch` の直後で深さ優先に展開します。
- `elseif` は判定対象に含みます。

### 直接実行

`RunMain` を直接実行した場合は、機能名の入力と対象ブックの選択を求めます。

## 4. EscapePartsMarking

### 目的

対象ブックを走査し、エスケープ対象関数のマーキングを行います。

### 主な入力

- 対象ブックパス
- 完了メッセージ
- エスケープ関数一覧
- グレーアウト対象
- 塗りつぶし色

### 既定のエスケープ対象関数

既定値は次の通りです。

```text
pg_escape_string,sqlS,sqlN,sqlLS,sqlC,sqlNZ,sqlInN,sqlF,sqlChk,sqlLikeStr,sqlNum,sqlNum0,sqlStr
```

### グレーアウト対象

`A列にしか値が入っていない行` について、次のいずれかを選択できます。

- `なし`
- `A列のみ`
- `B列のみ`
- `A,B列`

内部値は `None / Left / Right / Both` です。

### 直接実行

`RunMain` を直接実行した場合は、対象ブックの選択を求めます。

## 5. MacroToolsUserFormInstaller

### 役割

- `frmMacroTools` と `modMacroToolsFormEntry` を自動生成します。
- UserForm 実行ブリッジを提供します。
- CONFIG シートの作成、補助 UI、実行マクロを提供します。

### 提供マクロ

- `InstallMacroToolsUserForm`
- `OpenMacroToolsForm`
- `CreateMacroConfigSheet`
- `OpenMacroConfigSheet`
- `RunBetaEvidenceFromConfigSheet`
- `RunBetaTestCaseFromConfigSheet`
- `RunConditionalBranchCheckerFromConfigSheet`
- `RunEscapePartsMarkingFromConfigSheet`

## 入力検証

次の色入力は、実行前に `#RRGGBB / 0xRRGGBB / RRGGBB` 形式か検証します。

- 条件分岐チェックの塗りつぶし色
- エスケープ箇所マーキングの塗りつぶし色
- エビデンス生成の読み飛ばし色一覧

不正な値がある場合は、処理を開始せずにエラーを表示します。CONFIG シート実行時は `CONFIG!Bxx` 付きで対象セルを示します。

## よくあるエラーと対処

### `1004` エラー

`VBA プロジェクト オブジェクト モデルへのアクセスを信頼する` が OFF の可能性があります。トラストセンターの設定を確認してください。

### `75` エラー

パスや一時ファイル作成に失敗している可能性があります。ブックが保存済みか、対象フォルダに書き込み権限があるかを確認してください。

### `438` エラー

古いフォームや古いモジュールが残っている可能性があります。`InstallMacroSuiteFromSingleFile` と `InstallMacroToolsUserForm` を再実行してください。

### コンパイルエラー

古いモジュールが残っている可能性があります。単一ファイルを再インポートし、`InstallMacroSuiteFromSingleFile` を実行してからコンパイルしてください。

## 運用メモ

- 既定値を変える場合は、各本体モジュール先頭の `Const` を編集してください。
- UserForm と CONFIG は同じ実行ブリッジを通すため、入力検証の挙動は揃います。
- 対象ブックを直接更新するマクロを実行する前は、バックアップを取る運用を推奨します。
- 単一ファイル運用では、通常 `MacroSuiteSingleFileInstaller.bas` だけを配布・導入対象にします。