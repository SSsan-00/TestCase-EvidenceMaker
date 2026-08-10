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

## まず結論

フォームを使い始めるために、UserForm を自分で挿入したり、フォーム名を自分で変更したりする必要はありません。

通常は次の流れだけで使い始めます。

1. `MacroSuiteSingleFileInstaller.bas` をインポートします。
2. `InstallMacroSuiteFromSingleFile` を実行します。
3. `InstallMacroToolsUserForm` を実行します。
4. `OpenMacroToolsForm` を実行します。

`InstallMacroToolsUserForm` が、`frmMacroTools` という UserForm と、フォームを開くための `modMacroToolsFormEntry` を自動で作成します。

手動で UserForm を挿入する必要があるのは、`InstallMacroToolsUserForm` 実行時に `UserForm component could not be created` と表示された例外時だけです。通常運用では手動挿入しません。

## マクロ一覧に表示されるタイミング

導入直後は、すべてのマクロが最初から見えているわけではありません。表示される順番は次の通りです。

| 状態 | マクロ一覧で実行するもの | 実行後に増えるもの |
| --- | --- | --- |
| `MacroSuiteSingleFileInstaller.bas` をインポートした直後 | `InstallMacroSuiteFromSingleFile` | 本体 5 モジュール |
| `InstallMacroSuiteFromSingleFile` 実行後 | `InstallMacroToolsUserForm` | `frmMacroTools` と `modMacroToolsFormEntry` |
| `InstallMacroToolsUserForm` 実行後 | `OpenMacroToolsForm` | フォーム画面を開ける状態 |

つまり、`OpenMacroToolsForm` が最初から見えないのは正常です。先に `InstallMacroSuiteFromSingleFile`、次に `InstallMacroToolsUserForm` を実行してください。

## 導入前に確認すること

初回導入や再展開を行う前に、次を確認してください。

- 作業先ブックは `.xlsm` 形式で保存済みにします。
- マクロを有効化します。
- VBE のプロジェクトがパスワードロックされていない状態にします。
- フォームを再展開する場合は、開いている `frmMacroTools` を閉じます。
- `MacroSuiteSingleFileInstaller.bas` は通常、この 1 ファイルだけをインポートします。`forPro` 配下の個別 `.bas` は通常運用では直接インポートしません。

信頼設定の場所:
`ファイル > オプション > トラスト センター > トラスト センターの設定 > マクロの設定 > VBA プロジェクト オブジェクト モデルへのアクセスを信頼する`

この信頼設定が OFF の場合、フォーム作成やモジュール展開で `1004` エラーになることがあります。

## 導入用マクロの役割

導入時によく使うマクロは次の 5 つです。

- `InstallMacroSuiteFromSingleFile`: 単一ファイル内に埋め込まれている本体 5 モジュールを VBA プロジェクトへ展開します。
- `InstallMacroToolsUserForm`: `frmMacroTools` と `modMacroToolsFormEntry` を作成または再作成します。フォームの初回導入・再展開はこのマクロを使います。
- `OpenMacroToolsForm`: 作成済みのフォームを開きます。フォーム未導入の状態では使えません。
- `CreateMacroConfigSheet`: `CONFIG` シートを作成または作り直します。フォームだけ使う場合は必須ではありません。
- `RegenerateMacroSuiteSingleFileInstaller`: 開発者向けです。VBA プロジェクト内の本体モジュールから、配布用の `MacroSuiteSingleFileInstaller.bas` を再生成します。

通常の利用者が最初に実行する順番は、`InstallMacroSuiteFromSingleFile`、`InstallMacroToolsUserForm`、`OpenMacroToolsForm` です。

## フォーム初回導入手順

初めてフォームを使えるようにする場合は、次の順番で作業します。

1. 利用先の `.xlsm` を Excel で開きます。
2. Excel の信頼設定で `VBA プロジェクト オブジェクト モデルへのアクセスを信頼する` を ON にします。
3. `Alt + F11` で VBE を開きます。
4. VBE の `ファイル > ファイルのインポート` から `MacroSuiteSingleFileInstaller.bas` をインポートします。
5. Excel 側で `Alt + F8` を押します。
6. マクロ一覧から `InstallMacroSuiteFromSingleFile` を選択して実行します。
7. 完了メッセージが出たら、もう一度 `Alt + F8` を押します。
8. マクロ一覧から `InstallMacroToolsUserForm` を選択して実行します。
9. 完了メッセージが出たら、もう一度 `Alt + F8` を押します。
10. マクロ一覧から `OpenMacroToolsForm` を選択して実行します。
11. フォーム画面が表示されたら導入完了です。
12. `.xlsm` を保存します。
13. VBE の `Debug > Compile VBAProject` でコンパイル確認します。

この通常手順では、UserForm の手動挿入も、フォーム名の手動変更も行いません。`InstallMacroToolsUserForm` が自動で `frmMacroTools` を作成します。

VBE 側で確認したい場合は、次の状態になっていれば正常です。

- 標準モジュールに `BetaEvidenceGenerator`、`BetaTestCaseGenerator`、`ConditionalBranchChecker`、`EscapePartsMarking`、`MacroToolsUserFormInstaller` がある。
- フォームに `frmMacroTools` がある。
- 標準モジュールに `modMacroToolsFormEntry` がある。

`OpenMacroToolsForm` がマクロ一覧に出てこない場合は、まだ `InstallMacroToolsUserForm` が成功していません。先に `InstallMacroSuiteFromSingleFile`、次に `InstallMacroToolsUserForm` を実行してください。

## フォーム再展開手順

フォームの見た目、ラベル、入力欄、ボタン配置、フォーム連携処理を更新したい場合は、フォームを再展開します。

1. 開いている `frmMacroTools` を閉じます。
2. `InstallMacroToolsUserForm` を実行します。
3. 完了メッセージを確認します。
4. `OpenMacroToolsForm` を実行します。
5. 表示内容が更新されていることを確認します。
6. `.xlsm` を保存します。
7. VBE の `Debug > Compile VBAProject` でコンパイル確認します。

`InstallMacroToolsUserForm` は何度実行しても構いません。既存の `frmMacroTools` はレイアウトとコードが作り直され、既存の `modMacroToolsFormEntry` は削除して再作成されます。

フォーム再展開で更新されるもの:

- フォームのラベル文言
- 入力欄、チェックボックス、ラジオボタン、ボタンの配置
- フォームから各マクロを呼び出す処理
- フォームを開くための `OpenMacroToolsForm`

フォーム再展開だけでは更新されないもの:

- `BetaEvidenceGenerator` などの本体モジュール
- `CONFIG` シートの入力値
- 既に作成済みのエビデンス、テストケース、マーキング済みブック

本体マクロも更新したい場合は、次の「単一ファイル差し替え時の更新手順」を実行してください。

## 単一ファイル差し替え時の更新手順

新しい `MacroSuiteSingleFileInstaller.bas` を受け取った場合は、次の順で更新します。

1. 開いている `frmMacroTools` を閉じます。
2. VBE で既存の `MacroSuiteSingleFileInstaller` モジュールを削除します。
3. VBE の `ファイル > ファイルのインポート` から新しい `MacroSuiteSingleFileInstaller.bas` をインポートします。
4. `InstallMacroSuiteFromSingleFile` を実行します。
5. フォームを使う場合は `InstallMacroToolsUserForm` を実行します。
6. フォームを開く場合は `OpenMacroToolsForm` を実行します。
7. CONFIG シートを使う場合は `CreateMacroConfigSheet` を実行します。
8. `.xlsm` を保存します。
9. VBE の `Debug > Compile VBAProject` でコンパイル確認します。

`InstallMacroSuiteFromSingleFile` は展開先の同名モジュールを置き換えます。既存の UserForm や CONFIG シートは自動更新されないため、フォームを使う場合は必ず `InstallMacroToolsUserForm` を再実行してください。

## CONFIG シート再作成手順

CONFIG シートの項目、説明、ボタン配置を更新したい場合は、`CreateMacroConfigSheet` を再実行します。

1. 必要であれば既存の `CONFIG` シートの入力値を控えます。
2. `CreateMacroConfigSheet` を実行します。
3. 作成された `CONFIG` シートの B 列に設定値を入力します。
4. `.xlsm` を保存します。

`CreateMacroConfigSheet` は CONFIG シートを作り直すため、既存の入力値を残したい場合は事前に退避してください。

## 単一ファイルの再生成

この手順は開発者向けです。利用者がフォームを使い始めるための作業ではありません。

各本体モジュールを編集した後、配布用の単一ファイルを作り直す場合は `RegenerateMacroSuiteSingleFileInstaller` を実行します。

- 出力先は `ThisWorkbook.Path\MacroSuiteSingleFileInstaller.bas` です。
- 現在の VBA プロジェクト内にある 5 モジュールを元に、埋め込み済みの単一ファイルを再生成します。
- 再生成後、その新しい `MacroSuiteSingleFileInstaller.bas` を利用者へ配布します。

## 利用方法の選択

利用方法は 3 通りあります。

- UserForm: `OpenMacroToolsForm` から画面操作で実行します。
- CONFIG シート: `CreateMacroConfigSheet` で作成した設定シートから実行します。
- 直接実行: 展開後の各モジュールの `RunMain` を直接実行します。

通常運用では UserForm または CONFIG シートを推奨します。直接実行は、最小限の入力だけを求めるプロ向けの運用です。

## UserForm の使い方

初回導入または再展開が完了している場合は、次の手順でフォームを使います。

1. `OpenMacroToolsForm` を実行します。
2. 必要な入力欄を埋めます。
3. 各処理の実行ボタンを押します。
4. 作業が終わったらフォーム右上の `×` で閉じます。

フォームはモデルレス表示です。フォームを開いたまま、同じ Excel インスタンス内のブックやシートを操作できます。

フォームの表示内容が古い場合は、フォームを閉じてから `InstallMacroToolsUserForm` を実行し直してください。

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
- 個別シートへ書き込む ON/OFF
- 非Function行に `B1-` 形式を書き込む ON/OFF
- マーキングセルを塗りつぶす ON/OFF
- 塗りつぶし色

### 現行ソースシートのマーキング

- `function` 宣言行は A 列へ `★` を書き込みます。
- `function` 宣言行は B 列へ `B1` / `B2` ... を書き込みます。
- 最初の判定対象が `function` の場合、先頭Functionは常に `B1` から開始します。
- Functionより前に条件分岐がある場合は `MAIN` が `B1`、最初のFunctionが `B2` になります。
- A 列に既存値がある場合も、`function` 宣言行では `★` で上書きします。
- 前回実行で残った A 列の `★` は、今回の再判定前に消します。
- 非Function行は、オプション ON 時のみ B 列へ `B1-` 形式を書き込みます。
- マーキングセルの塗りつぶしは B 列セルに適用します。
- 再度ブックを開いたときのアクティブシートは、走査した現行ソースシートです。

### 個別シート出力

- 「個別シートへ書き込む」の既定値は ON です。
- OFFの場合も現行ソースへのマーキングは実行しますが、個別シートの検索、解析イベント収集、書き込み、行挿入、テンプレート退避は行いません。
- OFFの場合は対象の個別シートが存在しなくても処理できます。
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

### エスケープ関数マーキング

- `A1-1-1` を含むシートの B列を 4行目から走査します。
- `sqlS(...)` など、エスケープ関数一覧に含まれる `prefix(...)` を赤字・太字にします。
- 関数名は識別子の境界を含めて判定するため、`mySqlS(...)` の一部を `sqlS(...)` として誤検出しません。
- 文字列内、および PHP の行コメント（`//`、`#`）・ブロックコメント（`/* ... */`）内の記述は対象外です。
- `DbHelper.sqlS(...)` のように `.` 経由で呼び出す場合は、左側の識別子も含めて装飾します。
- 括弧のネストを解析するため、`sqlS(xxx + trim(yyy) + "zzz")` も外側の `sqlS(...)` 全体を装飾します。
- 文字列やコメント内の括弧は対応する閉じ括弧として数えません。
- B列の複数行に跨る `sqlS(` / `XXX` / `)` のような記述も、対応する閉じ括弧までを装飾します。
- 複数行に跨ってヒットした場合は、装飾された B列セルと同じ行の C列に完了メッセージを書きます。

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

### `OpenMacroToolsForm` が見つからない

`InstallMacroToolsUserForm` がまだ成功していない可能性があります。先に `InstallMacroSuiteFromSingleFile`、次に `InstallMacroToolsUserForm` を実行してください。

### `InstallMacroToolsUserForm` が見つからない

本体モジュールの展開が完了していない可能性があります。`MacroSuiteSingleFileInstaller.bas` をインポートしただけではフォーム作成マクロはまだ使えません。`InstallMacroSuiteFromSingleFile` を実行してください。

### フォームのラベルや配置が古い

古い `frmMacroTools` が残っています。フォームを閉じてから `InstallMacroToolsUserForm` を再実行し、その後 `OpenMacroToolsForm` で開き直してください。

### フォームを閉じられない、または再展開できない

フォームが開いたまま再展開しようとしている可能性があります。フォーム右上の `×` で閉じてから `InstallMacroToolsUserForm` を実行してください。閉じられない場合は、作業中のブックを保存して Excel を開き直してから再実行してください。

### `1004` エラー

`VBA プロジェクト オブジェクト モデルへのアクセスを信頼する` が OFF の可能性があります。トラストセンターの設定を確認してください。

### `75` エラー

パスや一時ファイル作成に失敗している可能性があります。ブックが保存済みか、対象フォルダに書き込み権限があるかを確認してください。

### `438` エラー

古いフォームや古いモジュールが残っている可能性があります。`InstallMacroSuiteFromSingleFile` と `InstallMacroToolsUserForm` を再実行してください。

### `UserForm component could not be created` と表示される

通常は UserForm の手動挿入は不要です。このエラーが出た場合だけ、環境によって UserForm の自動追加がブロックされている可能性があります。例外対応として、VBE で `挿入 > ユーザーフォーム` から UserForm を 1 つ手動追加し、プロパティウィンドウの `(Name)` を `frmMacroTools` に変更してから、`InstallMacroToolsUserForm` を再実行してください。

### コンパイルエラー

古いモジュールが残っている可能性があります。単一ファイルを再インポートし、`InstallMacroSuiteFromSingleFile`、`InstallMacroToolsUserForm` の順に実行してからコンパイルしてください。

## 運用メモ

- 既定値を変える場合は、各本体モジュール先頭の `Const` を編集してください。
- UserForm と CONFIG は同じ実行ブリッジを通すため、入力検証の挙動は揃います。
- 対象ブックを直接更新するマクロを実行する前は、バックアップを取る運用を推奨します。
- 単一ファイル運用では、通常 `MacroSuiteSingleFileInstaller.bas` だけを配布・導入対象にします。
