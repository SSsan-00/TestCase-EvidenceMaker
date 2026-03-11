Attribute VB_Name = "BetaEvidenceGenerator"
Option Explicit

' ============================================================
' 単体テストエビデンス シート生成マクロ
' ------------------------------------------------------------
' このモジュールは、マクロブック（ThisWorkbook）にある雛形シートを使って、
' ユーザーが選択した参照元ブック（xlsx）から新規出力ブックを作成し、そこへエビデンスシートを生成する
'
' 主な流れ
' 1. 参照元xlsxファイルを選択する
' 2. 入力ファイル名（例: foo.php）を入力する
' 3. マクロブックのREFERシートを参照して referValue を取得する
' 4. 参照元ブックの【共通】/【個別】参照元シートを走査する
' 5. 共通/個別それぞれの出力xlsxを新規作成する（同名時は連番）
' 6. 出力対象シート名を任意入力し、空欄なら全シートを出力する
' 7. 共通モードでは A1-1-1（指定時のみ）を独立シートとして出力する
' 8. エビデンスシートは A1 テンプレを複製して作成する
' 9. A1-1-1 の A3/B3 の 〇〇〇 を baseName に置換し、E/H列を規則で書き込む
' ============================================================

' ===== マクロブック内の固定シート名 =====
Private Const TEMPLATE_HEADER_SHEET_NAME As String = "A1-1-1" ' 共通モード先頭シート用テンプレ
Private Const TEMPLATE_BODY_SHEET_NAME As String = "A1"       ' エビデンスシート本体テンプレ
Private Const REFER_SHEET_NAME As String = "REFER"            ' 入力ファイル名 -> referValue 参照用

' ===== REFERシートの列定義（列記号 -> 列番号へ変換して使用） =====
' 注意:
' Cells(row, col) の第2引数に "E" のような列記号文字列を直接渡すと
' 実行時エラーになる環境があるため、必ず列番号へ変換してから使う
Private Const REFER_KEY_COL_LETTER As String = "E"   ' 入力ファイル名のキー列（完全一致）
Private Const REFER_VALUE_COL_LETTER As String = "F" ' referValue を取得する列
Private Const REFER_BETA_COL_LETTER As String = "D"  ' β（00形式にする番号）を取得する列
Private Const REFER_ALPHA_COL_LETTER As String = "J" ' α（拡張子なし元文字列）を取得する列

' ===== 参照元シートの走査条件 =====
Private Const SOURCE_START_ROW As Long = 8          ' 仕様にある開始行
Private Const SOURCE_COL_A As Long = 1              ' A列: 作成するエビデンスシート名
Private Const SOURCE_COL_B As Long = 5              ' E列: pendingB 用
Private Const SOURCE_COL_C As Long = 8              ' H列: 確定トリガ
Private Const EMPTY_STREAK_STOP_COUNT As Long = 100  ' A/E/H空行が連続したら走査終了

' ===== エビデンスシートへの書き込み（スロット） =====
Private Const FIRST_DEST_ROW As Long = 3 ' slot0 の書き込み開始行
Private Const SLOT_HEIGHT As Long = 50   ' 既定値: 50行刻み
Private Const DEST_COL_A As Long = 1     ' 書き込み先 A列
Private Const DEST_COL_B As Long = 2     ' 書き込み先 B列
Private Const BORDER_END_COL As Long = 32 ' 上罫線の終端（AF列）
Public Const OPTION_RIGHT_BORDER_ENABLED As Boolean = True ' True: 右罫線を適用 / False: 右罫線を適用しない
Private Const RIGHT_BORDER_TARGET_COL As Long = 17 ' 右罫線を引く対象列（既定: Q列）
Private Const RIGHT_BORDER_EXTRA_ROWS As Long = 50 ' 最終書き込み行から下方向へ延長する行数

' ===== 共通モード先頭シートのヘッダ置換 =====
Private Const HEADER_PLACEHOLDER As String = "〇〇〇"

' ===== Office定数を数値で扱う（参照設定に依存しにくくするため） =====
Private Const FILE_DIALOG_PICKER As Long = 3 ' msoFileDialogFilePicker
Public Const OPTION_TOP_BORDER_ENABLED As Boolean = True ' True: A:AAに上罫線を適用 / False: 上罫線を適用しない
Public Const OPTION_SLOT_HEIGHT_PROMPT_ENABLED As Boolean = True ' True: 行オフセット入力を表示 / False: SLOT_HEIGHTを使用
Public Const OPTION_OUTPUT_SHEET_SELECTION_PROMPT_ENABLED As Boolean = True ' True: 作成シート選択入力を表示 / False: 全シート出力
' 出力対象から除外したいシート名/パターンを Like 形式で指定
' 例: A4,A5,A1-1,A2-3-1,B3-*
Public Const OPTION_EXCLUDE_OUTPUT_SHEET_BY_PATTERN_ENABLED As Boolean = True ' True: 除外パターン一致シートを作成しない / False: すべて作成対象
Private Const EXCLUDED_OUTPUT_SHEET_NAME_PATTERNS As String = "A4,A5,A1-1,A2-3-1" ' Likeパターンをカンマ区切りで指定
' 参照元セルの塗りつぶし色が一致した場合、そのセル値を未入力扱いでスキップする
' 例: #f2f2f2,#d9d9d9,#bfbfbf,#a6a6a6,#808080
Public Const OPTION_SKIP_GRAY_FILLED_SOURCE_CELL_ENABLED As Boolean = True ' True: 灰色塗りつぶしセルを読み飛ばす / False: 色判定を行わない
Private Const SOURCE_SKIP_FILL_COLOR_HEX_CODES As String = "#f2f2f2,#d9d9d9,#bfbfbf,#a6a6a6,#808080" ' 比較対象カラーコード（#RRGGBB）
Public Type BetaEvidenceUiOptions
    Enabled As Boolean
    SourceWorkbookPath As String
    InputFileName As String

    UseSlotHeight As Boolean
    SlotHeight As Long

    UseOutputSheetFilter As Boolean
    OutputSheetFilterText As String

    OverrideTopBorderEnabled As Boolean
    TopBorderEnabled As Boolean

    OverrideSlotHeightPromptEnabled As Boolean
    SlotHeightPromptEnabled As Boolean

    OverrideOutputSheetSelectionPromptEnabled As Boolean
    OutputSheetSelectionPromptEnabled As Boolean

    OverrideExcludeOutputSheetByPatternEnabled As Boolean
    ExcludeOutputSheetByPatternEnabled As Boolean

    UseExcludedOutputSheetNamePatterns As Boolean
    ExcludedOutputSheetNamePatterns As String

    OverrideSkipGrayFilledSourceCellEnabled As Boolean
    SkipGrayFilledSourceCellEnabled As Boolean

    UseSourceSkipFillColorHexCodes As Boolean
    SourceSkipFillColorHexCodes As String

    OverrideRightBorderEnabled As Boolean
    RightBorderEnabled As Boolean

    UseRightBorderTargetCol As Boolean
    RightBorderTargetCol As Long
End Type

Private mUiOptions As BetaEvidenceUiOptions
Private mSlotHeight As Long ' スロット行オフセット（未指定時は既定値を使用）
Private mSkipSourceFillColorMap As Object ' 参照元塗りつぶしスキップ色マップ

Public Sub RunMainWithUiOptions(ByRef options As BetaEvidenceUiOptions)
    ClearUiOptions
    mUiOptions = options
    mUiOptions.Enabled = True

    RunMain

    ClearUiOptions
End Sub

Public Function CreateBetaEvidenceUiOptionsForForm() As BetaEvidenceUiOptions
    Dim options As BetaEvidenceUiOptions

    InitializeBetaEvidenceUiOptionsForForm options
    CreateBetaEvidenceUiOptionsForForm = options
End Function

Public Sub InitializeBetaEvidenceUiOptionsForForm(ByRef options As BetaEvidenceUiOptions)
    options.Enabled = True
    options.SourceWorkbookPath = vbNullString
    options.InputFileName = vbNullString

    options.UseSlotHeight = True
    options.SlotHeight = SLOT_HEIGHT

    options.UseOutputSheetFilter = True
    options.OutputSheetFilterText = vbNullString

    options.OverrideTopBorderEnabled = True
    options.TopBorderEnabled = OPTION_TOP_BORDER_ENABLED

    options.OverrideSlotHeightPromptEnabled = True
    options.SlotHeightPromptEnabled = False

    options.OverrideOutputSheetSelectionPromptEnabled = True
    options.OutputSheetSelectionPromptEnabled = False

    options.OverrideExcludeOutputSheetByPatternEnabled = True
    options.ExcludeOutputSheetByPatternEnabled = OPTION_EXCLUDE_OUTPUT_SHEET_BY_PATTERN_ENABLED

    options.UseExcludedOutputSheetNamePatterns = True
    options.ExcludedOutputSheetNamePatterns = EXCLUDED_OUTPUT_SHEET_NAME_PATTERNS

    options.OverrideSkipGrayFilledSourceCellEnabled = True
    options.SkipGrayFilledSourceCellEnabled = OPTION_SKIP_GRAY_FILLED_SOURCE_CELL_ENABLED

    options.UseSourceSkipFillColorHexCodes = True
    options.SourceSkipFillColorHexCodes = SOURCE_SKIP_FILL_COLOR_HEX_CODES

    options.OverrideRightBorderEnabled = True
    options.RightBorderEnabled = OPTION_RIGHT_BORDER_ENABLED

    options.UseRightBorderTargetCol = True
    options.RightBorderTargetCol = RIGHT_BORDER_TARGET_COL
End Sub

Private Sub ClearUiOptions()
    mUiOptions.Enabled = False
    mUiOptions.SourceWorkbookPath = vbNullString
    mUiOptions.InputFileName = vbNullString

    mUiOptions.UseSlotHeight = False
    mUiOptions.SlotHeight = 0

    mUiOptions.UseOutputSheetFilter = False
    mUiOptions.OutputSheetFilterText = vbNullString

    mUiOptions.OverrideTopBorderEnabled = False
    mUiOptions.TopBorderEnabled = False

    mUiOptions.OverrideSlotHeightPromptEnabled = False
    mUiOptions.SlotHeightPromptEnabled = False

    mUiOptions.OverrideOutputSheetSelectionPromptEnabled = False
    mUiOptions.OutputSheetSelectionPromptEnabled = False

    mUiOptions.OverrideExcludeOutputSheetByPatternEnabled = False
    mUiOptions.ExcludeOutputSheetByPatternEnabled = False

    mUiOptions.UseExcludedOutputSheetNamePatterns = False
    mUiOptions.ExcludedOutputSheetNamePatterns = vbNullString

    mUiOptions.OverrideSkipGrayFilledSourceCellEnabled = False
    mUiOptions.SkipGrayFilledSourceCellEnabled = False

    mUiOptions.UseSourceSkipFillColorHexCodes = False
    mUiOptions.SourceSkipFillColorHexCodes = vbNullString

    mUiOptions.OverrideRightBorderEnabled = False
    mUiOptions.RightBorderEnabled = False

    mUiOptions.UseRightBorderTargetCol = False
    mUiOptions.RightBorderTargetCol = 0
End Sub

' ============================================================
' エントリポイント
' ============================================================

Public Sub RunMain()
    On Error GoTo ErrorHandler

    Dim macroWb As Workbook
    Dim sourceWb As Workbook
    Dim targetWb As Workbook
    Dim referWs As Worksheet
    Dim templateBodyWs As Worksheet
    Dim templateHeaderWs As Worksheet

    Dim targetPath As String
    Dim inputFileName As String
    Dim baseName As String
    Dim referValue As String
    Dim expectedCommonWorkbookName As String
    Dim expectedIndividualWorkbookName As String

    Dim commonSourceSheetName As String
    Dim commonSourceWs As Worksheet
    Dim individualSourceWs As Worksheet

    Dim commonSummary As String
    Dim individualSummary As String
    Dim finalMessage As String
    Dim processedAnyMode As Boolean
    Dim sourceWbWasAlreadyOpen As Boolean

    Dim commonOutputPath As String
    Dim individualOutputPath As String
    Dim seedSheetName As String
    Dim outputSheetFilter As Object
    Dim outputSheetFilterLabel As String
    Dim commonCreatedSheetCount As Long
    Dim individualCreatedSheetCount As Long
    Dim commonHeaderCreated As Boolean

    ' Application状態は、エラー時でも必ず元に戻す
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim appStateCaptured As Boolean

    Set macroWb = ThisWorkbook

    ' まず必要なテンプレ/REFERシートが存在するか確認して、以降の処理を分かりやすく失敗させる
    Set templateBodyWs = GetWorksheetOrRaise(macroWb, TEMPLATE_BODY_SHEET_NAME, "雛形シート（本体）")
    Set templateHeaderWs = GetWorksheetOrRaise(macroWb, TEMPLATE_HEADER_SHEET_NAME, "雛形シート（ヘッダー）")
    Set referWs = GetWorksheetOrRaise(macroWb, REFER_SHEET_NAME, "REFERシート")
    ' 参照元になるxlsxファイルを選択する
    targetPath = SelectTargetWorkbookPath()
    If Len(targetPath) = 0 Then
        MsgBox "処理をキャンセルしました（参照元ブックが未選択です）。", vbInformation
        Exit Sub
    End If

    ' REFER検索キーになる入力ファイル名を受け取る（例: foo.php）。
    inputFileName = PromptInputFileName()
    If Len(inputFileName) = 0 Then
        MsgBox "処理をキャンセルしました（入力ファイル名が未入力です）。", vbInformation
        Exit Sub
    End If

    If mUiOptions.Enabled Then
        If mUiOptions.UseSlotHeight And mUiOptions.SlotHeight > 0 Then
            mSlotHeight = mUiOptions.SlotHeight
        Else
            mSlotHeight = SLOT_HEIGHT
        End If
    ElseIf IsSlotHeightPromptEnabled() Then
        mSlotHeight = PromptSlotHeightOrDefault(SLOT_HEIGHT)
    Else
        mSlotHeight = SLOT_HEIGHT ' 入力ダイアログOFF時は既定オフセットをそのまま使う
    End If

    If mUiOptions.Enabled Then
        If mUiOptions.UseOutputSheetFilter Then
            Set outputSheetFilter = ParseOutputSheetFilter(mUiOptions.OutputSheetFilterText)
        Else
            Set outputSheetFilter = Nothing
        End If
    ElseIf IsOutputSheetSelectionPromptEnabled() Then
        Set outputSheetFilter = PromptOutputSheetFilter()
    Else
        Set outputSheetFilter = Nothing ' 入力ダイアログOFF時は全シートを出力対象にする
    End If

    outputSheetFilterLabel = BuildOutputSheetFilterLabel(outputSheetFilter)
    Set mSkipSourceFillColorMap = BuildSkipSourceFillColorMap()

    ' 後続処理で共通/個別シート名や置換に使うため、拡張子なし名を作成する
    baseName = RemoveExtension(inputFileName)
    If Len(baseName) = 0 Then
        Err.Raise vbObjectError + 2001, "RunMain", _
                  "入力ファイル名から拡張子なしの名前を取得できませんでした。"
    End If

    ' 参照元ブックを開く（既に開いていればそのインスタンスを再利用）。
    sourceWbWasAlreadyOpen = IsWorkbookAlreadyOpen(targetPath)
    Set sourceWb = OpenTargetWorkbook(targetPath, True)
    If sourceWb Is Nothing Then
        Err.Raise vbObjectError + 2002, "RunMain", _
                  "参照元ブックを開けませんでした。"
    End If

    ' REFERシートから referValue を取得する（キーは入力ファイル名）。
    referValue = GetReferValueFromReferSheet(referWs, inputFileName)

    ' REFERシートから出力ファイル名を構築する。
    BuildEvidenceWorkbookNamesFromRefer referWs, inputFileName, _
                                       expectedCommonWorkbookName, expectedIndividualWorkbookName

    ' 参照元の対象シートを取得する（この時点では sourceWb を見る）。
    commonSourceSheetName = "【共通】" & referValue
    Set commonSourceWs = FindWorksheetExact(sourceWb, commonSourceSheetName)
    Set individualSourceWs = FindIndividualSourceSheet(sourceWb, referValue)

    ' 速度改善のため、画面更新や再計算を一時的に止める
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    appStateCaptured = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' -------------------------
    ' 共通モード（【共通】）
    ' -------------------------
    If commonSourceWs Is Nothing Then
        commonSummary = "共通モード: スキップ（参照元シートなし: " & commonSourceSheetName & "）"
    Else
        commonOutputPath = ResolveOutputWorkbookPath(BuildOutputWorkbookPath(targetPath, expectedCommonWorkbookName))
        Set targetWb = CreateEmptyOutputWorkbook(commonOutputPath, seedSheetName)

        commonCreatedSheetCount = 0
        commonHeaderCreated = False
        If IsSheetAllowedByFilter(TEMPLATE_HEADER_SHEET_NAME, outputSheetFilter) Then
            CreateCommonHeaderSheet targetWb, templateHeaderWs, baseName
            commonHeaderCreated = True
        End If

        commonSummary = ProcessReferenceSheet( _
            sourceWs:=commonSourceWs, _
            targetWb:=targetWb, _
            templateBodyWs:=templateBodyWs, _
            templateHeaderWs:=templateHeaderWs, _
            baseName:=baseName, _
            applyHeaderOverlay:=False, _
            modeLabel:="共通", _
            outputSheetFilter:=outputSheetFilter, _
            createdSheetCountOut:=commonCreatedSheetCount)

        If commonHeaderCreated Then
            commonCreatedSheetCount = commonCreatedSheetCount + 1
        End If

        If commonCreatedSheetCount > 0 Then
            RemoveSeedSheetIfNeeded targetWb, seedSheetName
            ActivateFirstWorksheetForOpenState targetWb
            targetWb.Save
            targetWb.Close SaveChanges:=True
            Set targetWb = Nothing

            commonSummary = commonSummary & " / 出力: " & commonOutputPath
            processedAnyMode = True
        Else
            DiscardOutputWorkbookAndFile targetWb, commonOutputPath
            commonSummary = commonSummary & " / 出力対象シートなしのためファイル未出力"
        End If
    End If

    ' -------------------------
    ' 個別モード（【個別】）
    ' 参照元シート名: 【個別】referValue（完全一致）


    ' -------------------------
    If individualSourceWs Is Nothing Then
        individualSummary = "個別モード: スキップ（参照元シートなし: 【個別】" & referValue & "）"
    Else
        individualOutputPath = ResolveOutputWorkbookPath(BuildOutputWorkbookPath(targetPath, expectedIndividualWorkbookName))
        Set targetWb = CreateEmptyOutputWorkbook(individualOutputPath, seedSheetName)
        individualCreatedSheetCount = 0

        individualSummary = ProcessReferenceSheet( _
            sourceWs:=individualSourceWs, _
            targetWb:=targetWb, _
            templateBodyWs:=templateBodyWs, _
            templateHeaderWs:=templateHeaderWs, _
            baseName:=baseName, _
            applyHeaderOverlay:=False, _
            modeLabel:="個別", _
            outputSheetFilter:=outputSheetFilter, _
            createdSheetCountOut:=individualCreatedSheetCount)

        If individualCreatedSheetCount > 0 Then
            RemoveSeedSheetIfNeeded targetWb, seedSheetName
            ActivateFirstWorksheetForOpenState targetWb
            targetWb.Save
            targetWb.Close SaveChanges:=True
            Set targetWb = Nothing

            individualSummary = individualSummary & " / 出力: " & individualOutputPath
            processedAnyMode = True
        Else
            DiscardOutputWorkbookAndFile targetWb, individualOutputPath
            individualSummary = individualSummary & " / 出力対象シートなしのためファイル未出力"
        End If
    End If

    If processedAnyMode Then
        finalMessage = "処理が完了しました。" & vbCrLf & _
                       "参照元ブック: " & sourceWb.Name & vbCrLf & _
                       "入力ファイル名: " & inputFileName & vbCrLf & _
                       "baseName: " & baseName & vbCrLf & _
                       "REFER(F): " & referValue & vbCrLf & _
                       "出力シート指定: " & outputSheetFilterLabel & vbCrLf & _
                       "想定ファイル名（共通）: " & expectedCommonWorkbookName & vbCrLf & _
                       "想定ファイル名（個別）: " & expectedIndividualWorkbookName & vbCrLf & vbCrLf & _
                       commonSummary & vbCrLf & _
                       individualSummary
    Else
        finalMessage = "出力対象シートが見つからなかったため、出力は作成されませんでした。" & vbCrLf & _
                       "参照元ブック: " & sourceWb.Name & vbCrLf & _
                       "確認対象: " & commonSourceSheetName & " / 【個別】" & referValue & vbCrLf & _
                       "出力シート指定: " & outputSheetFilterLabel
    End If

    GoTo SafeExit

ErrorHandler:
    finalMessage = "エラーが発生しました。" & vbCrLf & _
                   Err.Number & " : " & Err.Description

SafeExit:
    On Error Resume Next
    Application.CutCopyMode = False

    If Not targetWb Is Nothing Then
        targetWb.Close SaveChanges:=False
    End If

    If Not sourceWb Is Nothing Then
        If Not sourceWbWasAlreadyOpen Then
            sourceWb.Close SaveChanges:=False
        End If
    End If

    Set mSkipSourceFillColorMap = Nothing
    If appStateCaptured Then
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        Application.EnableEvents = prevEnableEvents
        Application.Calculation = prevCalculation
    End If
    On Error GoTo 0

    If Len(finalMessage) > 0 Then
        If Left$(finalMessage, 6) = "エラーが発生" Then
            MsgBox finalMessage, vbExclamation
        Else
            MsgBox finalMessage, vbInformation
        End If
    End If
End Sub

' ============================================================
' 入力・ブック取得
' ============================================================

Private Function SelectTargetWorkbookPath() As String
    ' FileDialog を使って、参照元の xlsx をユーザーに選ばせる
    ' 参照設定依存を避けるため、FileDialog型ではなく Object で扱う
    Dim fd As Object

    If mUiOptions.Enabled Then
        SelectTargetWorkbookPath = Trim$(mUiOptions.SourceWorkbookPath)
        Exit Function
    End If

    On Error GoTo Fallback

    Set fd = Application.FileDialog(FILE_DIALOG_PICKER)
    With fd
        .Title = "参照元のxlsxファイルを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel ブック (*.xlsx)", "*.xlsx"

        If .Show <> -1 Then
            SelectTargetWorkbookPath = vbNullString
            Exit Function
        End If

        SelectTargetWorkbookPath = CStr(.SelectedItems(1))
    End With
    Exit Function

Fallback:
    ' 環境差で FileDialog が使えない場合に備え、GetOpenFilename にフォールバックする
    Dim selectedPath As Variant

    selectedPath = Application.GetOpenFilename( _
        FileFilter:="Excel ブック (*.xlsx),*.xlsx", _
        Title:="参照元のxlsxファイルを選択してください")

    If VarType(selectedPath) = vbBoolean Then
        SelectTargetWorkbookPath = vbNullString
    Else
        SelectTargetWorkbookPath = CStr(selectedPath)
    End If
End Function

Private Function PromptInputFileName() As String
    ' REFER検索キーになる入力ファイル名を受け取る
    ' 前後の空白は誤入力になりやすいため Trim する
    Dim s As String

    If mUiOptions.Enabled Then
        PromptInputFileName = Trim$(mUiOptions.InputFileName)
        Exit Function
    End If

    s = InputBox("入力ファイル名を入力してください（例: menu/mainmenu.php）", "入力ファイル名")
    PromptInputFileName = Trim$(s)
End Function
Private Function PromptSlotHeightOrDefault(ByVal defaultHeight As Long) As Long
    ' スロットの行オフセットを受け取る（空欄は既定値）
    Dim inputText As String
    Dim numericValue As Double

    inputText = InputBox( _
        "スロットの行オフセットを入力してください（空欄は既定値 " & CStr(defaultHeight) & "）。" & vbCrLf & _
        "例: 50", _
        "スロット行オフセット", _
        CStr(defaultHeight))

    inputText = Trim$(inputText)
    If Len(inputText) = 0 Then
        PromptSlotHeightOrDefault = defaultHeight
        Exit Function
    End If

    If Not IsNumeric(inputText) Then
        MsgBox "スロット行オフセットが数値ではないため、既定値 " & CStr(defaultHeight) & " を使用します。", vbExclamation
        PromptSlotHeightOrDefault = defaultHeight
        Exit Function
    End If

    numericValue = CDbl(inputText)
    If numericValue <= 0 Or numericValue <> Fix(numericValue) Then
        MsgBox "スロット行オフセットは1以上の整数で入力してください。既定値 " & CStr(defaultHeight) & " を使用します。", vbExclamation
        PromptSlotHeightOrDefault = defaultHeight
        Exit Function
    End If

    PromptSlotHeightOrDefault = CLng(numericValue)
End Function

Private Function PromptOutputSheetFilter() As Object
    Dim inputText As String

    inputText = InputBox( _
        "出力するシート名をカンマ区切りで入力してください（任意）。" & vbCrLf & _
        "例: A1,A2,B1" & vbCrLf & _
        "空欄の場合は全シートを出力します。", _
        "出力シート名（任意）")

    Set PromptOutputSheetFilter = ParseOutputSheetFilter(inputText)
End Function

Private Function ParseOutputSheetFilter(ByVal rawInput As String) As Object
    Dim normalizedText As String
    Dim names As Variant
    Dim nameText As String
    Dim i As Long
    Dim dict As Object

    normalizedText = Replace(rawInput, "，", ",")
    names = Split(normalizedText, ",")

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    For i = LBound(names) To UBound(names)
        nameText = Trim$(CStr(names(i)))
        If Len(nameText) > 0 Then
            If Not dict.Exists(nameText) Then
                dict.Add nameText, True
            End If
        End If
    Next i

    If dict.Count = 0 Then
        Set ParseOutputSheetFilter = Nothing
    Else
        Set ParseOutputSheetFilter = dict
    End If
End Function

Private Function BuildOutputSheetFilterLabel(ByVal outputSheetFilter As Object) As String
    Dim key As Variant
    Dim sheetNames As String

    If outputSheetFilter Is Nothing Then
        BuildOutputSheetFilterLabel = "全シート（指定なし）"
        Exit Function
    End If

    For Each key In outputSheetFilter.Keys
        If Len(sheetNames) > 0 Then
            sheetNames = sheetNames & ", "
        End If
        sheetNames = sheetNames & CStr(key)
    Next key

    If Len(sheetNames) = 0 Then
        BuildOutputSheetFilterLabel = "全シート（指定なし）"
    Else
        BuildOutputSheetFilterLabel = sheetNames
    End If
End Function

Private Function IsTopBorderEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideTopBorderEnabled Then
        IsTopBorderEnabled = mUiOptions.TopBorderEnabled
    Else
        IsTopBorderEnabled = OPTION_TOP_BORDER_ENABLED
    End If
End Function

Private Function IsSlotHeightPromptEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideSlotHeightPromptEnabled Then
        IsSlotHeightPromptEnabled = mUiOptions.SlotHeightPromptEnabled
    Else
        IsSlotHeightPromptEnabled = OPTION_SLOT_HEIGHT_PROMPT_ENABLED
    End If
End Function

Private Function IsOutputSheetSelectionPromptEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideOutputSheetSelectionPromptEnabled Then
        IsOutputSheetSelectionPromptEnabled = mUiOptions.OutputSheetSelectionPromptEnabled
    Else
        IsOutputSheetSelectionPromptEnabled = OPTION_OUTPUT_SHEET_SELECTION_PROMPT_ENABLED
    End If
End Function

Private Function IsExcludeOutputSheetByPatternEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideExcludeOutputSheetByPatternEnabled Then
        IsExcludeOutputSheetByPatternEnabled = mUiOptions.ExcludeOutputSheetByPatternEnabled
    Else
        IsExcludeOutputSheetByPatternEnabled = OPTION_EXCLUDE_OUTPUT_SHEET_BY_PATTERN_ENABLED
    End If
End Function

Private Function GetExcludedOutputSheetNamePatterns() As String
    If mUiOptions.Enabled And mUiOptions.UseExcludedOutputSheetNamePatterns Then
        GetExcludedOutputSheetNamePatterns = CStr(mUiOptions.ExcludedOutputSheetNamePatterns)
    Else
        GetExcludedOutputSheetNamePatterns = EXCLUDED_OUTPUT_SHEET_NAME_PATTERNS
    End If
End Function

Private Function IsSkipGrayFilledSourceCellEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideSkipGrayFilledSourceCellEnabled Then
        IsSkipGrayFilledSourceCellEnabled = mUiOptions.SkipGrayFilledSourceCellEnabled
    Else
        IsSkipGrayFilledSourceCellEnabled = OPTION_SKIP_GRAY_FILLED_SOURCE_CELL_ENABLED
    End If
End Function

Private Function GetSourceSkipFillColorHexCodes() As String
    If mUiOptions.Enabled And mUiOptions.UseSourceSkipFillColorHexCodes Then
        GetSourceSkipFillColorHexCodes = CStr(mUiOptions.SourceSkipFillColorHexCodes)
    Else
        GetSourceSkipFillColorHexCodes = SOURCE_SKIP_FILL_COLOR_HEX_CODES
    End If
End Function

Private Function IsRightBorderEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideRightBorderEnabled Then
        IsRightBorderEnabled = mUiOptions.RightBorderEnabled
    Else
        IsRightBorderEnabled = OPTION_RIGHT_BORDER_ENABLED
    End If
End Function

Private Function GetRightBorderTargetCol() As Long
    If mUiOptions.Enabled And mUiOptions.UseRightBorderTargetCol Then
        GetRightBorderTargetCol = mUiOptions.RightBorderTargetCol
    Else
        GetRightBorderTargetCol = RIGHT_BORDER_TARGET_COL
    End If
End Function

Private Function IsSheetAllowedByFilter( _
    ByVal sheetName As String, _
    ByVal outputSheetFilter As Object) As Boolean

    If outputSheetFilter Is Nothing Then
        IsSheetAllowedByFilter = True
    Else
        IsSheetAllowedByFilter = outputSheetFilter.Exists(sheetName)
    End If
End Function

Private Function IsExcludedByOutputSheetPattern(ByVal sheetName As String) As Boolean
    Dim normalizedName As String
    Dim rawPatterns As String
    Dim patterns As Variant
    Dim patternText As String
    Dim i As Long

    If Not IsExcludeOutputSheetByPatternEnabled() Then Exit Function

    normalizedName = Trim$(sheetName)
    If Len(normalizedName) = 0 Then Exit Function

    rawPatterns = Replace(GetExcludedOutputSheetNamePatterns(), "，", ",")
    patterns = Split(rawPatterns, ",")

    For i = LBound(patterns) To UBound(patterns)
        patternText = Trim$(CStr(patterns(i)))
        If Len(patternText) > 0 Then
            If normalizedName Like patternText Then
                IsExcludedByOutputSheetPattern = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function BuildSkipSourceFillColorMap() As Object
    Dim dict As Object
    Dim rawText As String
    Dim rawItems As Variant
    Dim normalizedHex As String
    Dim colorValue As Long
    Dim i As Long

    If Not IsSkipGrayFilledSourceCellEnabled() Then Exit Function

    rawText = Replace(GetSourceSkipFillColorHexCodes(), "，", ",")
    rawItems = Split(rawText, ",")

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    For i = LBound(rawItems) To UBound(rawItems)
        normalizedHex = NormalizeHexColorTextForFillRule(CStr(rawItems(i)))
        If Len(normalizedHex) > 0 Then
            colorValue = HexColorTextToColorLong(normalizedHex)
            If colorValue >= 0 Then
                If Not dict.Exists(CStr(colorValue)) Then
                    dict.Add CStr(colorValue), True
                End If
            End If
        End If
    Next i

    If dict.Count = 0 Then
        Set BuildSkipSourceFillColorMap = Nothing
    Else
        Set BuildSkipSourceFillColorMap = dict
    End If
End Function

Private Function NormalizeHexColorTextForFillRule(ByVal rawText As String) As String
    Dim t As String
    Dim i As Long
    Dim ch As String

    t = UCase$(Trim$(rawText))
    If Len(t) = 0 Then Exit Function

    If Left$(t, 1) = "#" Then
        t = Mid$(t, 2)
    ElseIf Left$(t, 2) = "0X" Then
        t = Mid$(t, 3)
    End If

    If Len(t) <> 6 Then Exit Function

    For i = 1 To 6
        ch = Mid$(t, i, 1)
        If InStr(1, "0123456789ABCDEF", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i

    NormalizeHexColorTextForFillRule = t
End Function

Private Function HexColorTextToColorLong(ByVal normalizedHex6 As String) As Long
    Dim redPart As Long
    Dim greenPart As Long
    Dim bluePart As Long

    On Error GoTo ConversionError

    redPart = CLng("&H" & Mid$(normalizedHex6, 1, 2))
    greenPart = CLng("&H" & Mid$(normalizedHex6, 3, 2))
    bluePart = CLng("&H" & Mid$(normalizedHex6, 5, 2))

    HexColorTextToColorLong = RGB(redPart, greenPart, bluePart)
    Exit Function

ConversionError:
    HexColorTextToColorLong = -1
End Function

Private Function ShouldSkipSourceCellByFillColor( _
    ByVal sourceWs As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal columnNumber As Long) As Boolean

    Dim colorValue As Long

    If mSkipSourceFillColorMap Is Nothing Then Exit Function

    colorValue = sourceWs.Cells(rowNumber, columnNumber).Interior.Color
    ShouldSkipSourceCellByFillColor = mSkipSourceFillColorMap.Exists(CStr(colorValue))
End Function

Private Function OpenTargetWorkbook(ByVal workbookPath As String, Optional ByVal openReadOnly As Boolean = False) As Workbook
    ' 既に同じファイルが開いている場合は再利用し、未オープンなら開く
    Dim wb As Workbook

    If Len(Trim$(workbookPath)) = 0 Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, workbookPath, vbTextCompare) = 0 Then
            Set OpenTargetWorkbook = wb
            Exit Function
        End If
    Next wb

    Set OpenTargetWorkbook = Application.Workbooks.Open( _
        Filename:=workbookPath, _
        UpdateLinks:=0, _
        ReadOnly:=openReadOnly)
End Function

Private Function BuildOutputWorkbookPath( _
    ByVal sourceWorkbookPath As String, _
    ByVal outputWorkbookName As String) As String

    Dim lastSepPos As Long
    Dim folderPath As String

    If Len(Trim$(outputWorkbookName)) = 0 Then
        Err.Raise vbObjectError + 2005, "BuildOutputWorkbookPath", "出力ファイル名が空です。"
    End If

    lastSepPos = InStrRev(sourceWorkbookPath, "\")
    If lastSepPos <= 0 Then
        BuildOutputWorkbookPath = outputWorkbookName
    Else
        folderPath = Left$(sourceWorkbookPath, lastSepPos)
        BuildOutputWorkbookPath = folderPath & outputWorkbookName
    End If
End Function

Private Function ResolveOutputWorkbookPath(ByVal desiredOutputPath As String) As String

    Dim lastSepPos As Long
    Dim lastDotPos As Long
    Dim basePath As String
    Dim extensionPart As String
    Dim candidatePath As String
    Dim seqNo As Long

    If Len(Trim$(desiredOutputPath)) = 0 Then
        Err.Raise vbObjectError + 2010, "ResolveOutputWorkbookPath", "出力先パスが空です。"
    End If

    If Len(Dir$(desiredOutputPath)) = 0 And Not IsWorkbookAlreadyOpen(desiredOutputPath) Then
        ResolveOutputWorkbookPath = desiredOutputPath
        Exit Function
    End If

    lastSepPos = InStrRev(desiredOutputPath, "\")
    lastDotPos = InStrRev(desiredOutputPath, ".")
    If lastDotPos > (lastSepPos + 1) Then
        basePath = Left$(desiredOutputPath, lastDotPos - 1)
        extensionPart = Mid$(desiredOutputPath, lastDotPos)
    Else
        basePath = desiredOutputPath
        extensionPart = vbNullString
    End If

    For seqNo = 1 To 9999
        candidatePath = basePath & "_" & Format$(seqNo, "000") & extensionPart
        If Len(Dir$(candidatePath)) = 0 And Not IsWorkbookAlreadyOpen(candidatePath) Then
            ResolveOutputWorkbookPath = candidatePath
            Exit Function
        End If
    Next seqNo

    Err.Raise vbObjectError + 2011, "ResolveOutputWorkbookPath", _
              "連番付き出力先を決定できませんでした（上限: 9999）。"
End Function

Private Function IsWorkbookAlreadyOpen(ByVal workbookPath As String) As Boolean
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, workbookPath, vbTextCompare) = 0 Then
            IsWorkbookAlreadyOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Function CreateEmptyOutputWorkbook( _
    ByVal outputPath As String, _
    ByRef seedSheetName As String) As Workbook

    Dim wb As Workbook

    If Len(Trim$(outputPath)) = 0 Then
        Err.Raise vbObjectError + 2006, "CreateEmptyOutputWorkbook", "出力先パスが空です。"
    End If

    If IsWorkbookAlreadyOpen(outputPath) Then
        Err.Raise vbObjectError + 2009, "CreateEmptyOutputWorkbook", _
                  "同名の出力ブックが既に開かれています。閉じてから再実行してください。" & vbCrLf & outputPath
    End If

    If Len(Dir$(outputPath)) > 0 Then
        Err.Raise vbObjectError + 2012, "CreateEmptyOutputWorkbook", _
                  "出力先ファイルが既に存在します（連番解決漏れ）。" & vbCrLf & outputPath
    End If

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    seedSheetName = wb.Worksheets(1).Name
    wb.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook

    Set CreateEmptyOutputWorkbook = wb
End Function

Private Sub RemoveSeedSheetIfNeeded(ByVal wb As Workbook, ByVal seedSheetName As String)
    If wb Is Nothing Then Exit Sub
    If Len(seedSheetName) = 0 Then Exit Sub
    If wb.Worksheets.Count <= 1 Then Exit Sub

    DeleteWorksheetIfExists wb, seedSheetName
End Sub

Private Sub ActivateFirstWorksheetForOpenState(ByVal wb As Workbook)
    Dim firstWs As Worksheet

    If wb Is Nothing Then Exit Sub
    If wb.Worksheets.Count = 0 Then Exit Sub

    Set firstWs = wb.Worksheets(1)
    firstWs.Activate
End Sub


Private Sub DiscardOutputWorkbookAndFile( _
    ByRef wb As Workbook, _
    ByVal outputPath As String)

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If

    If Len(Trim$(outputPath)) > 0 Then
        If Len(Dir$(outputPath)) > 0 Then
            Kill outputPath
        End If
    End If

    On Error GoTo 0
End Sub

Private Sub CreateCommonHeaderSheet( _
    ByVal targetWb As Workbook, _
    ByVal templateHeaderWs As Worksheet, _
    ByVal baseName As String)

    Dim headerWs As Worksheet

    DeleteWorksheetIfExists targetWb, TEMPLATE_HEADER_SHEET_NAME

    templateHeaderWs.Copy After:=targetWb.Worksheets(targetWb.Worksheets.Count)
    Set headerWs = targetWb.Worksheets(targetWb.Worksheets.Count)

    On Error GoTo RenameError
    If StrComp(headerWs.Name, TEMPLATE_HEADER_SHEET_NAME, vbBinaryCompare) <> 0 Then
        headerWs.Name = TEMPLATE_HEADER_SHEET_NAME
    End If
    On Error GoTo 0

    ReplaceHeaderPlaceholderInSheet headerWs, baseName
    Exit Sub

RenameError:
    Err.Raise vbObjectError + 2221, "CreateCommonHeaderSheet", _
              "共通ヘッダシート名を設定できませんでした: " & TEMPLATE_HEADER_SHEET_NAME
End Sub

' ============================================================
' REFER参照
' ============================================================

Private Function GetReferValueFromReferSheet( _
    ByVal referWs As Worksheet, _
    ByVal inputFileName As String) As String

    ' REFERシートから、入力ファイル名をキーに該当行を探し、F列の値を返す
    ' 仕様上、完全一致を前提にする
    Dim matchedRow As Long
    Dim matchCount As Long
    Dim valueColIndex As Long
    Dim referValueRaw As Variant

    matchedRow = FindRowByExactMatch(referWs, REFER_KEY_COL_LETTER, inputFileName, matchCount)

    If matchCount = 0 Then
        Err.Raise vbObjectError + 2101, "GetReferValueFromReferSheet", _
                  "REFERシートの" & REFER_KEY_COL_LETTER & "列に完全一致する値が見つかりませんでした。" & vbCrLf & _
                  "入力値: " & inputFileName
    End If

    If matchCount > 1 Then
        Err.Raise vbObjectError + 2102, "GetReferValueFromReferSheet", _
                  "REFERシートの" & REFER_KEY_COL_LETTER & "列に完全一致する値が複数あります。" & vbCrLf & _
                  "入力値: " & inputFileName & vbCrLf & _
                  "件数: " & CStr(matchCount)
    End If

    valueColIndex = ColumnLetterToIndex(REFER_VALUE_COL_LETTER, "REFER値列")
    referValueRaw = referWs.Cells(matchedRow, valueColIndex).Value

    If IsError(referValueRaw) Then
        Err.Raise vbObjectError + 2103, "GetReferValueFromReferSheet", _
                  "REFERシートの" & REFER_VALUE_COL_LETTER & "列にエラー値が入っています。"
    End If

    GetReferValueFromReferSheet = Trim$(CStr(referValueRaw))
    If Len(GetReferValueFromReferSheet) = 0 Then
        Err.Raise vbObjectError + 2104, "GetReferValueFromReferSheet", _
                  "REFERシートの" & REFER_VALUE_COL_LETTER & "列の値が空です。"
    End If
End Function

Private Sub BuildEvidenceWorkbookNamesFromRefer( _
    ByVal referWs As Worksheet, _
    ByVal inputFileName As String, _
    ByRef commonWorkbookName As String, _
    ByRef individualWorkbookName As String)

    ' 旧仕様との互換のため、REFERシートから α/β/γ を組み立てて
    ' 作成対象のファイル名（共通/個別）を決定する。
    '
    ' α = J列（拡張子なし）
    ' β = D列（00形式）
    ' γ = F列
    Dim matchedRow As Long
    Dim matchCount As Long
    Dim alphaColIndex As Long
    Dim betaColIndex As Long
    Dim gammaColIndex As Long
    Dim alphaText As String
    Dim betaText As String
    Dim gammaText As String
    Dim betaRaw As Variant

    matchedRow = FindRowByExactMatch(referWs, REFER_KEY_COL_LETTER, inputFileName, matchCount)

    If matchCount = 0 Then
        Err.Raise vbObjectError + 2131, "BuildEvidenceWorkbookNamesFromRefer", _
                  "REFERシートから想定ファイル名を組み立てるためのキーが見つかりません。" & vbCrLf & _
                  "キー(" & REFER_KEY_COL_LETTER & "列): " & inputFileName
    End If

    If matchCount > 1 Then
        Err.Raise vbObjectError + 2132, "BuildEvidenceWorkbookNamesFromRefer", _
                  "REFERシートから想定ファイル名を組み立てる対象が複数あります。" & vbCrLf & _
                  "キー(" & REFER_KEY_COL_LETTER & "列): " & inputFileName & vbCrLf & _
                  "件数: " & CStr(matchCount)
    End If

    alphaColIndex = ColumnLetterToIndex(REFER_ALPHA_COL_LETTER, "α列")
    betaColIndex = ColumnLetterToIndex(REFER_BETA_COL_LETTER, "β列")
    gammaColIndex = ColumnLetterToIndex(REFER_VALUE_COL_LETTER, "γ列")

    alphaText = RemoveExtension(GetTrimmedCellStringOrRaise( _
        referWs.Cells(matchedRow, alphaColIndex).Value, _
        "BuildEvidenceWorkbookNamesFromRefer", _
        "REFERシートの" & REFER_ALPHA_COL_LETTER & "列（α）"))
    If Len(alphaText) = 0 Then
        Err.Raise vbObjectError + 2133, "BuildEvidenceWorkbookNamesFromRefer", _
                  "REFERシートの" & REFER_ALPHA_COL_LETTER & "列（α）から拡張子なし文字列を取得できませんでした。"
    End If

    betaRaw = referWs.Cells(matchedRow, betaColIndex).Value
    If IsError(betaRaw) Then
        Err.Raise vbObjectError + 2134, "BuildEvidenceWorkbookNamesFromRefer", _
                  "REFERシートの" & REFER_BETA_COL_LETTER & "列（β）にエラー値が入っています。"
    End If
    betaText = ToTwoDigitStringStrict(betaRaw, "REFERシートの" & REFER_BETA_COL_LETTER & "列（β）")

    gammaText = GetTrimmedCellStringOrRaise( _
        referWs.Cells(matchedRow, gammaColIndex).Value, _
        "BuildEvidenceWorkbookNamesFromRefer", _
        "REFERシートの" & REFER_VALUE_COL_LETTER & "列（γ）")

    commonWorkbookName = alphaText & "_【共通】" & betaText & gammaText & "_単体テストエビデンス_初期開発.xlsx"
    individualWorkbookName = alphaText & "_【個別】" & betaText & gammaText & "_単体テストエビデンス_初期開発.xlsx"
End Sub

Private Function BuildTargetNameCheckMessage( _
    ByVal actualTargetWorkbookName As String, _
    ByVal expectedCommonWorkbookName As String, _
    ByVal expectedIndividualWorkbookName As String) As String

    If StrComp(actualTargetWorkbookName, expectedCommonWorkbookName, vbBinaryCompare) = 0 Or _
       StrComp(actualTargetWorkbookName, expectedIndividualWorkbookName, vbBinaryCompare) = 0 Then
        BuildTargetNameCheckMessage = "ターゲット名照合: OK（REFERから決まる想定名と一致）"
    Else
        BuildTargetNameCheckMessage = "ターゲット名照合: 注意（REFER想定名と不一致のまま処理を継続）"
    End If
End Function

Private Function FindRowByExactMatch( _
    ByVal ws As Worksheet, _
    ByVal targetColLetter As String, _
    ByVal searchValue As String, _
    ByRef matchCount As Long) As Long

    ' 文字列の完全一致（vbBinaryCompare）で検索
    ' 大文字/小文字や全角/半角の違いも区別
    Dim targetColIndex As Long
    Dim lastRow As Long
    Dim r As Long
    Dim cellValue As Variant
    Dim cellText As String

    matchCount = 0
    targetColIndex = ColumnLetterToIndex(targetColLetter, "検索列")
    lastRow = ws.Cells(ws.Rows.Count, targetColIndex).End(xlUp).Row

    If lastRow < 1 Then Exit Function

    For r = 1 To lastRow
        cellValue = ws.Cells(r, targetColIndex).Value

        If IsError(cellValue) Then
            Err.Raise vbObjectError + 2111, "FindRowByExactMatch", _
                      "REFERシートの検索列にエラー値が含まれています（行: " & CStr(r) & "）。"
        End If

        cellText = CStr(cellValue)
        If StrComp(cellText, searchValue, vbBinaryCompare) = 0 Then
            matchCount = matchCount + 1
            If FindRowByExactMatch = 0 Then
                FindRowByExactMatch = r
            End If
        End If
    Next r
End Function

Private Function ColumnLetterToIndex( _
    ByVal columnLetter As String, _
    Optional ByVal labelForError As String = vbNullString) As Long

    ' "A" -> 1, "F" -> 6, "AA" -> 27 のように列記号を列番号へ変換
    ' Cells(row, "F") のような文字列渡しを避けるための関数
    Dim normalized As String
    Dim i As Long
    Dim ch As String
    Dim chCode As Long
    Dim prefix As String

    normalized = UCase$(Trim$(columnLetter))
    If Len(normalized) = 0 Then
        prefix = BuildErrorLabelPrefix(labelForError)
        Err.Raise vbObjectError + 2121, "ColumnLetterToIndex", prefix & "列指定が空です。"
    End If

    For i = 1 To Len(normalized)
        ch = Mid$(normalized, i, 1)
        chCode = Asc(ch)

        If chCode < 65 Or chCode > 90 Then
            prefix = BuildErrorLabelPrefix(labelForError)
            Err.Raise vbObjectError + 2122, "ColumnLetterToIndex", _
                      prefix & "列指定が不正です: " & columnLetter
        End If

        ColumnLetterToIndex = (ColumnLetterToIndex * 26) + (chCode - 64)
    Next i

    If ColumnLetterToIndex < 1 Or ColumnLetterToIndex > 16384 Then
        prefix = BuildErrorLabelPrefix(labelForError)
        Err.Raise vbObjectError + 2123, "ColumnLetterToIndex", _
                  prefix & "列番号がExcelの範囲外です: " & columnLetter
    End If
End Function

Private Function GetTrimmedCellStringOrRaise( _
    ByVal cellValue As Variant, _
    ByVal callerName As String, _
    ByVal valueLabel As String) As String

    If IsError(cellValue) Then
        Err.Raise vbObjectError + 2124, callerName, valueLabel & " にエラー値が入っています。"
    End If

    GetTrimmedCellStringOrRaise = Trim$(CStr(cellValue))
    If Len(GetTrimmedCellStringOrRaise) = 0 Then
        Err.Raise vbObjectError + 2125, callerName, valueLabel & " が空です。"
    End If
End Function

Private Function ToTwoDigitStringStrict( _
    ByVal valueD As Variant, _
    ByVal valueLabel As String) As String

    Dim numericValue As Double
    Dim normalizedText As String

    ' 仕様: REFERのD列(β)がNULL/空の場合は "01" を採用する
    If IsEmpty(valueD) Or IsNull(valueD) Then
        ToTwoDigitStringStrict = "01"
        Exit Function
    End If

    If VarType(valueD) = vbString Then
        normalizedText = Trim$(CStr(valueD))

        If Len(normalizedText) = 0 Then
            ToTwoDigitStringStrict = "01"
            Exit Function
        End If

        If StrComp(normalizedText, "NULL", vbTextCompare) = 0 Then
            ToTwoDigitStringStrict = "01"
            Exit Function
        End If

        valueD = normalizedText
    End If

    If Not IsNumeric(valueD) Then
        Err.Raise vbObjectError + 2127, "ToTwoDigitStringStrict", _
                  valueLabel & " は数値である必要があります（例: 1, 2, 10）。"
    End If

    numericValue = CDbl(valueD)
    If numericValue <> Fix(numericValue) Then
        Err.Raise vbObjectError + 2128, "ToTwoDigitStringStrict", _
                  valueLabel & " は整数である必要があります。"
    End If

    ToTwoDigitStringStrict = Format$(CLng(numericValue), "00")
End Function

Private Function BuildErrorLabelPrefix(ByVal labelText As String) As String
    If Len(Trim$(labelText)) = 0 Then
        BuildErrorLabelPrefix = vbNullString
    Else
        BuildErrorLabelPrefix = Trim$(labelText) & " "
    End If
End Function

' ============================================================
' 参照元シート探索
' ============================================================

Private Function FindWorksheetExact(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    ' シート名完全一致で取得します。見つからない場合は Nothing を返す
    On Error Resume Next
    Set FindWorksheetExact = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function FindIndividualSourceSheet(ByVal targetWb As Workbook, ByVal referValue As String) As Worksheet
    ' 個別モードの参照元: 【個別】referValue（完全一致）


    Dim candidateName As String

    candidateName = "【個別】" & referValue
    Set FindIndividualSourceSheet = FindWorksheetExact(targetWb, candidateName)



End Function

' ============================================================
' 参照元シート -> エビデンスシート生成
' ============================================================

Private Function ProcessReferenceSheet( _
    ByVal sourceWs As Worksheet, _
    ByVal targetWb As Workbook, _
    ByVal templateBodyWs As Worksheet, _
    ByVal templateHeaderWs As Worksheet, _
    ByVal baseName As String, _
    ByVal applyHeaderOverlay As Boolean, _
    ByVal modeLabel As String, _
    ByVal outputSheetFilter As Object, _
    ByRef createdSheetCountOut As Long) As String

    ' 参照元シート（共通または個別）を走査し、A/E/Hのルールに従って
    ' エビデンスシートを作成・更新する
    Dim r As Long
    Dim emptyStreak As Long

    Dim currentEvidenceWs As Worksheet
    Dim currentEvidenceSheetName As String

    Dim slotIndex As Long
    Dim hasPendingB As Boolean
    Dim pendingB As Variant

    Dim rawA As Variant
    Dim rawB As Variant
    Dim rawC As Variant

    Dim hasA As Boolean
    Dim hasB As Boolean
    Dim hasC As Boolean
    Dim aSheetName As String
    Dim templateWsForCreate As Worksheet
    Dim useHeaderTemplateForFirstSheet As Boolean

    Dim createdSheetCount As Long
    Dim slotWriteCount As Long
    Dim ignoredDataBeforeSheetCount As Long
    Dim skippedByFilterCount As Long
    Dim skippedByPatternRuleCount As Long

    Dim maxRowA As Long
    Dim maxRowB As Long
    Dim maxRowC As Long
    Dim scanEndRow As Long
    Dim sourceValuesA As Variant
    Dim sourceValuesB As Variant
    Dim sourceValuesC As Variant
    Dim rowOffset As Long

    maxRowA = GetLastUsedRowInColumn(sourceWs, SOURCE_COL_A)
    maxRowB = GetLastUsedRowInColumn(sourceWs, SOURCE_COL_B)
    maxRowC = GetLastUsedRowInColumn(sourceWs, SOURCE_COL_C)

    scanEndRow = maxRowA
    If maxRowB > scanEndRow Then scanEndRow = maxRowB
    If maxRowC > scanEndRow Then scanEndRow = maxRowC
    If scanEndRow < SOURCE_START_ROW Then scanEndRow = SOURCE_START_ROW

    ' ループ中のセル参照を減らすため、必要列を配列へ読み込む
    scanEndRow = scanEndRow + EMPTY_STREAK_STOP_COUNT
    sourceValuesA = ReadColumnValuesFromRow(sourceWs, SOURCE_COL_A, SOURCE_START_ROW, scanEndRow)
    sourceValuesB = ReadColumnValuesFromRow(sourceWs, SOURCE_COL_B, SOURCE_START_ROW, scanEndRow)
    sourceValuesC = ReadColumnValuesFromRow(sourceWs, SOURCE_COL_C, SOURCE_START_ROW, scanEndRow)

    emptyStreak = 0
    Set currentEvidenceWs = Nothing
    currentEvidenceSheetName = vbNullString
    slotIndex = 0
    hasPendingB = False
    useHeaderTemplateForFirstSheet = applyHeaderOverlay
    createdSheetCountOut = 0

    For rowOffset = 1 To UBound(sourceValuesA, 1)
        r = SOURCE_START_ROW + rowOffset - 1

        rawA = sourceValuesA(rowOffset, 1)
        rawB = sourceValuesB(rowOffset, 1)
        rawC = sourceValuesC(rowOffset, 1)
        If ShouldSkipSourceCellByFillColor(sourceWs, r, SOURCE_COL_A) Then rawA = vbNullString
        If ShouldSkipSourceCellByFillColor(sourceWs, r, SOURCE_COL_B) Then rawB = vbNullString
        If ShouldSkipSourceCellByFillColor(sourceWs, r, SOURCE_COL_C) Then rawC = vbNullString

        ' エラー値が紛れていると原因が分かりにくくなるため、行番号付きで即時中断する
        EnsureNotErrorValue rawA, sourceWs.Name, r, "A"
        EnsureNotErrorValue rawB, sourceWs.Name, r, "E"
        EnsureNotErrorValue rawC, sourceWs.Name, r, "H"

        hasA = HasValueForSourceCell(rawA)
        hasB = HasValueForSourceCell(rawB)
        hasC = HasValueForSourceCell(rawC)

        If (Not hasA) And (Not hasB) And (Not hasC) Then
            emptyStreak = emptyStreak + 1
        Else
            emptyStreak = 0
        End If

        ' A列に値が来たら、現在シートを切り替える
        ' その前に pendingB が残っていれば、前シートに B単体として確定させる
        If hasA Then
            If Not currentEvidenceWs Is Nothing Then
                FlushPendingBIfNeeded currentEvidenceWs, slotIndex, hasPendingB, pendingB, slotWriteCount
            End If

            aSheetName = NormalizeEvidenceSheetName(rawA, sourceWs.Name, r)
            ' 共通モードでは A1-1-1 は共通ヘッダ専用のため、参照元A列からは作成しない
            If StrComp(modeLabel, "共通", vbBinaryCompare) = 0 And _
               StrComp(aSheetName, TEMPLATE_HEADER_SHEET_NAME, vbBinaryCompare) = 0 Then
                skippedByPatternRuleCount = skippedByPatternRuleCount + 1
                Set currentEvidenceWs = Nothing
                currentEvidenceSheetName = vbNullString
                slotIndex = 0
                hasPendingB = False
            ElseIf IsExcludedByOutputSheetPattern(aSheetName) Then
                skippedByPatternRuleCount = skippedByPatternRuleCount + 1
                Set currentEvidenceWs = Nothing
                currentEvidenceSheetName = vbNullString
                slotIndex = 0
                hasPendingB = False
            ElseIf Not IsSheetAllowedByFilter(aSheetName, outputSheetFilter) Then
                skippedByFilterCount = skippedByFilterCount + 1
                Set currentEvidenceWs = Nothing
                currentEvidenceSheetName = vbNullString
                slotIndex = 0
                hasPendingB = False
            Else
                ' 共通モードの先頭1シートのみ A1-1-1 を使い、
                ' それ以外は A1 を使う
                If useHeaderTemplateForFirstSheet Then
                    Set templateWsForCreate = templateHeaderWs
                Else
                    Set templateWsForCreate = templateBodyWs
                End If

                Set currentEvidenceWs = RecreateEvidenceSheetFromTemplate( _
                    targetWb:=targetWb, _
                    templateSourceWs:=templateWsForCreate, _
                    newSheetName:=aSheetName, _
                    currentSourceSheetName:=sourceWs.Name)

                createdSheetCount = createdSheetCount + 1
                currentEvidenceSheetName = aSheetName

                ' シートが変わったら、スロットと pendingB を新しいシート用に初期化する
                slotIndex = 0
                hasPendingB = False

                ' 共通モードの先頭1シートのみ、〇〇〇 を baseName に置換する
                If useHeaderTemplateForFirstSheet Then
                    ReplaceHeaderPlaceholderInSheet currentEvidenceWs, baseName
                    useHeaderTemplateForFirstSheet = False
                End If
            End If
        End If

        ' E/H は「現在のエビデンスシート」が決まっている場合にのみ処理を行う
        ' Aがまだ一度も出ていない場合は、仕様に必要な書き込み先が未確定なのでスキップする
        If hasB Or hasC Then
            If currentEvidenceWs Is Nothing Then
                ignoredDataBeforeSheetCount = ignoredDataBeforeSheetCount + 1
            Else
                ' 先にBを pending として保持（同一行に C がある場合、直後の C でペア確定させるため）
                If hasB Then
                    ' Bが連続で来た場合は、前のpendingBを単体として確定してから新しいBを保持する
                    If hasPendingB Then
                        FlushPendingBIfNeeded currentEvidenceWs, slotIndex, hasPendingB, pendingB, slotWriteCount
                    End If

                    pendingB = rawB
                    hasPendingB = True
                End If

                If hasC Then
                    If hasPendingB Then
                        WritePairSlot currentEvidenceWs, slotIndex, pendingB, rawC
                        slotWriteCount = slotWriteCount + 1
                        slotIndex = slotIndex + 1
                        hasPendingB = False
                    Else
                        WriteCOnlySlot currentEvidenceWs, slotIndex, rawC
                        slotWriteCount = slotWriteCount + 1
                        slotIndex = slotIndex + 1
                    End If
                End If
            End If
        End If

        If emptyStreak >= EMPTY_STREAK_STOP_COUNT Then
            Exit For
        End If
    Next rowOffset

    ' 走査終了時にも pendingB が残っていれば、最後の1件を取りこぼさないよう確定させる
    If Not currentEvidenceWs Is Nothing Then
        FlushPendingBIfNeeded currentEvidenceWs, slotIndex, hasPendingB, pendingB, slotWriteCount
    End If

    createdSheetCountOut = createdSheetCount

    ProcessReferenceSheet = modeLabel & "モード: 完了（参照元=" & sourceWs.Name & _
                           ", 作成シート数=" & CStr(createdSheetCount) & _
                           ", スロット書込数=" & CStr(slotWriteCount) & _
                           IIf(ignoredDataBeforeSheetCount > 0, _
                               ", 先行E/Hスキップ行=" & CStr(ignoredDataBeforeSheetCount), _
                               vbNullString) & _
                           IIf(skippedByPatternRuleCount > 0, _
                               ", パターン除外シート=" & CStr(skippedByPatternRuleCount), _
                               vbNullString) & _
                           IIf(skippedByFilterCount > 0, _
                               ", フィルタ除外シート=" & CStr(skippedByFilterCount), _
                               vbNullString) & ")"
End Function

' ============================================================
' エビデンスシート作成・テンプレ適用
' ============================================================

Private Function RecreateEvidenceSheetFromTemplate( _
    ByVal targetWb As Workbook, _
    ByVal templateSourceWs As Worksheet, _
    ByVal newSheetName As String, _
    ByVal currentSourceSheetName As String) As Worksheet

    ' 同名シートが既にある場合は削除して作り直す
    ' ただし、現在走査中の参照元シートは削除してはいけないので保護する
    ValidateWorksheetName newSheetName

    If StrComp(newSheetName, currentSourceSheetName, vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 2201, "RecreateEvidenceSheetFromTemplate", _
                  "参照元シート名と同じ名前のエビデンスシートは作成できません: " & newSheetName
    End If

    DeleteWorksheetIfExists targetWb, newSheetName, currentSourceSheetName

    ' 指定された雛形シート（A1 または A1-1-1）をコピーして、
    ' 新しいエビデンスシートを作る
    templateSourceWs.Copy After:=targetWb.Worksheets(targetWb.Worksheets.Count)
    Set RecreateEvidenceSheetFromTemplate = targetWb.Worksheets(targetWb.Worksheets.Count)

    On Error GoTo RenameError
    RecreateEvidenceSheetFromTemplate.Name = newSheetName
    On Error GoTo 0
    Exit Function

RenameError:
    Err.Raise vbObjectError + 2202, "RecreateEvidenceSheetFromTemplate", _
              "エビデンスシート名を設定できませんでした: " & newSheetName & vbCrLf & _
              "（シート名の文字数・使用禁止文字・重複を確認してください）"
End Function

Private Sub DeleteWorksheetIfExists( _
    ByVal wb As Workbook, _
    ByVal targetSheetName As String, _
    Optional ByVal protectedSheetName As String = vbNullString)

    ' 既存シート削除用ヘルパー。
    ' DisplayAlerts は上位で OFF にしている前提だが、ここでは警告表示の制御は行わない
    Dim ws As Worksheet

    Set ws = FindWorksheetExact(wb, targetSheetName)
    If ws Is Nothing Then Exit Sub

    If Len(protectedSheetName) > 0 Then
        If StrComp(ws.Name, protectedSheetName, vbBinaryCompare) = 0 Then
            Err.Raise vbObjectError + 2211, "DeleteWorksheetIfExists", _
                      "保護対象のシートを削除しようとしました: " & ws.Name
        End If
    End If

    ' マクロブックの雛形シートは絶対に削除しないよう、念のためガードします。
    If wb Is ThisWorkbook Then
        If StrComp(ws.Name, TEMPLATE_BODY_SHEET_NAME, vbBinaryCompare) = 0 Or _
           StrComp(ws.Name, TEMPLATE_HEADER_SHEET_NAME, vbBinaryCompare) = 0 Then
            Err.Raise vbObjectError + 2212, "DeleteWorksheetIfExists", _
                      "マクロブックの雛形シートは削除できません: " & ws.Name
        End If
    End If

    ws.Delete
End Sub

Private Sub ReplaceHeaderPlaceholderInSheet( _
    ByVal evidenceWs As Worksheet, _
    ByVal baseName As String)

    ' 共通モード先頭シートの A3/B3 のみを置換対象にする
    Dim cellA3 As Range
    Dim cellB3 As Range

    Set cellA3 = evidenceWs.Range("A3")
    Set cellB3 = evidenceWs.Range("B3")

    If Not IsError(cellA3.Value) Then
        cellA3.Value = Replace(CStr(cellA3.Value), HEADER_PLACEHOLDER, baseName, 1, -1, vbTextCompare)
    End If

    If Not IsError(cellB3.Value) Then
        cellB3.Value = Replace(CStr(cellB3.Value), HEADER_PLACEHOLDER, baseName, 1, -1, vbTextCompare)
    End If
End Sub

' ============================================================
' 参照元 A/E/H の読み取り補助
' ============================================================

Private Sub EnsureNotErrorValue( _
    ByVal cellValue As Variant, _
    ByVal sheetName As String, _
    ByVal rowNumber As Long, _
    ByVal colLetter As String)

    If IsError(cellValue) Then
        Err.Raise vbObjectError + 2301, "EnsureNotErrorValue", _
                  "参照元シートにエラー値が含まれています。" & vbCrLf & _
                  "シート: " & sheetName & " / セル: " & colLetter & CStr(rowNumber)
    End If
End Sub

Private Function HasValueForSourceCell(ByVal cellValue As Variant) As Boolean
    ' A/E/H列の「値あり判定」。
    ' 文字列は Trim 後に空なら空扱い、数値は 0 でも値あり扱いにする
    If IsEmpty(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function

    If VarType(cellValue) = vbString Then
        HasValueForSourceCell = (Len(Trim$(CStr(cellValue))) > 0)
    Else
        HasValueForSourceCell = (Len(CStr(cellValue)) > 0)
    End If
End Function

Private Function GetLastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnIndex As Long) As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
    If lastRow < 1 Then
        lastRow = 1
    End If

    GetLastUsedRowInColumn = lastRow
End Function

Private Function ReadColumnValuesFromRow( _
    ByVal ws As Worksheet, _
    ByVal columnIndex As Long, _
    ByVal startRow As Long, _
    ByVal endRow As Long) As Variant

    Dim rawValues As Variant
    Dim singleCell(1 To 1, 1 To 1) As Variant

    If endRow < startRow Then
        endRow = startRow
    End If

    rawValues = ws.Range(ws.Cells(startRow, columnIndex), ws.Cells(endRow, columnIndex)).Value

    If startRow = endRow Then
        singleCell(1, 1) = rawValues
        ReadColumnValuesFromRow = singleCell
    Else
        ReadColumnValuesFromRow = rawValues
    End If
End Function

Private Function NormalizeEvidenceSheetName( _
    ByVal rawValue As Variant, _
    ByVal sourceSheetName As String, _
    ByVal rowNumber As Long) As String

    ' A列の値をシート名として使うため、文字列化＋前後空白除去を行う
    ' 空になってしまう場合は呼び出し元のロジックと矛盾するため、明示的にエラーにする
    NormalizeEvidenceSheetName = Trim$(CStr(rawValue))

    If Len(NormalizeEvidenceSheetName) = 0 Then
        Err.Raise vbObjectError + 2311, "NormalizeEvidenceSheetName", _
                  "A列のシート名が空です（シート: " & sourceSheetName & ", 行: " & CStr(rowNumber) & "）。"
    End If
End Function

Private Sub ValidateWorksheetName(ByVal sheetNameText As String)
    ' Excelシート名として明らかに不正な値は、コピー/リネーム前に弾いて原因を明確にする
    Dim invalidChars As Variant
    Dim i As Long

    If Len(sheetNameText) = 0 Then
        Err.Raise vbObjectError + 2321, "ValidateWorksheetName", "シート名が空です。"
    End If

    If Len(sheetNameText) > 31 Then
        Err.Raise vbObjectError + 2322, "ValidateWorksheetName", _
                  "シート名は31文字以内である必要があります: " & sheetNameText
    End If

    invalidChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(invalidChars) To UBound(invalidChars)
        If InStr(1, sheetNameText, CStr(invalidChars(i)), vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 2323, "ValidateWorksheetName", _
                      "シート名に使用できない文字が含まれています: " & CStr(invalidChars(i))
        End If
    Next i
End Sub

' ============================================================
' スロット書き込み（E/H -> エビデンスシート）
' ============================================================

Private Sub FlushPendingBIfNeeded( _
    ByVal destWs As Worksheet, _
    ByRef slotIndex As Long, _
    ByRef hasPendingB As Boolean, _
    ByRef pendingB As Variant, _
    ByRef slotWriteCount As Long)

    ' pendingB が残っている場合、仕様どおり B単体 として1スロット書き込む
    If Not hasPendingB Then Exit Sub

    WriteBOnlySlot destWs, slotIndex, pendingB
    slotWriteCount = slotWriteCount + 1
    slotIndex = slotIndex + 1
    hasPendingB = False
End Sub

Private Sub WritePairSlot( _
    ByVal destWs As Worksheet, _
    ByVal slotIndex As Long, _
    ByVal pendingB As Variant, _
    ByVal cValue As Variant)

    ' ペア書き込みは、どのスロットでも A/B 列に固定する
    Dim destRow As Long

    destRow = GetDestRowForSlot(slotIndex)

    destWs.Cells(destRow, DEST_COL_A).Value = pendingB
    destWs.Cells(destRow, DEST_COL_B).Value = cValue
    ApplyTopBorderToConfirmedRow destWs, destRow
    ApplyRightBorderToConfiguredColumn destWs, destRow
End Sub

Private Sub WriteCOnlySlot( _
    ByVal destWs As Worksheet, _
    ByVal slotIndex As Long, _
    ByVal cValue As Variant)

    ' C単体は、どのスロットでも B列に書き込む
    Dim destRow As Long

    destRow = GetDestRowForSlot(slotIndex)
    destWs.Cells(destRow, DEST_COL_B).Value = cValue
    ApplyTopBorderToConfirmedRow destWs, destRow
    ApplyRightBorderToConfiguredColumn destWs, destRow
End Sub

Private Sub WriteBOnlySlot( _
    ByVal destWs As Worksheet, _
    ByVal slotIndex As Long, _
    ByVal bValue As Variant)

    ' B単体は、どのスロットでも A列に書き込む
    Dim destRow As Long

    destRow = GetDestRowForSlot(slotIndex)
    destWs.Cells(destRow, DEST_COL_A).Value = bValue
    ApplyTopBorderToConfirmedRow destWs, destRow
    ApplyRightBorderToConfiguredColumn destWs, destRow
End Sub

Private Sub ApplyRightBorderToConfiguredColumn( _
    ByVal destWs As Worksheet, _
    ByVal lastWrittenRow As Long)

    Dim endRow As Long
    Dim targetCol As Long
    Dim borderRange As Range

    If Not IsRightBorderEnabled() Then Exit Sub
    If lastWrittenRow < FIRST_DEST_ROW Then Exit Sub

    targetCol = GetRightBorderTargetCol()
    If targetCol < 1 Or targetCol > 16384 Then
        targetCol = 26
    End If

    endRow = lastWrittenRow + RIGHT_BORDER_EXTRA_ROWS
    Set borderRange = destWs.Range( _
        destWs.Cells(FIRST_DEST_ROW, targetCol), _
        destWs.Cells(endRow, targetCol))

    With borderRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Private Sub ApplyTopBorderToConfirmedRow( _
    ByVal destWs As Worksheet, _
    ByVal targetRow As Long)

    Dim aValue As Variant
    Dim bValue As Variant
    Dim hasAValue As Boolean
    Dim hasBValue As Boolean
    Dim borderRange As Range

    If targetRow = FIRST_DEST_ROW Then Exit Sub
    If Not IsTopBorderEnabled() Then Exit Sub ' OFF時は上罫線処理をスキップ

    aValue = destWs.Cells(targetRow, DEST_COL_A).Value
    bValue = destWs.Cells(targetRow, DEST_COL_B).Value

    If IsError(aValue) Then Exit Sub
    If IsError(bValue) Then Exit Sub

    hasAValue = (Len(Trim$(CStr(aValue))) > 0)
    hasBValue = (Len(Trim$(CStr(bValue))) > 0)

    If Not hasAValue And Not hasBValue Then Exit Sub

    Set borderRange = destWs.Range( _
        destWs.Cells(targetRow, DEST_COL_A), _
        destWs.Cells(targetRow, BORDER_END_COL))

    With borderRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Private Function GetDestRowForSlot(ByVal slotIndex As Long) As Long
    Dim slotHeightForWrite As Long

    If slotIndex < 0 Then
        Err.Raise vbObjectError + 2401, "GetDestRowForSlot", "slotIndex が負数です。"
    End If

    slotHeightForWrite = mSlotHeight
    If slotHeightForWrite <= 0 Then
        slotHeightForWrite = SLOT_HEIGHT
    End If

    GetDestRowForSlot = FIRST_DEST_ROW + (slotIndex * slotHeightForWrite)
End Function

' ============================================================
' 共通ユーティリティ
' ============================================================

Private Function GetWorksheetOrRaise( _
    ByVal wb As Workbook, _
    ByVal sheetName As String, _
    ByVal labelForMessage As String) As Worksheet

    Set GetWorksheetOrRaise = FindWorksheetExact(wb, sheetName)
    If GetWorksheetOrRaise Is Nothing Then
        Err.Raise vbObjectError + 2501, "GetWorksheetOrRaise", _
                  labelForMessage & " が見つかりません: " & sheetName
    End If
End Function

Private Function RemoveExtension(ByVal fileNameText As String) As String
    ' "foo.php" -> "foo"
    ' "foo.bar.php" -> "foo.bar"
    ' "foo" -> "foo"
    ' パスが混ざっていても最後の区切り以降だけを対象にする
    Dim lastDotPos As Long
    Dim lastSlashPos As Long
    Dim lastBackslashPos As Long
    Dim lastSeparatorPos As Long

    fileNameText = Trim$(fileNameText)
    If Len(fileNameText) = 0 Then Exit Function

    lastSlashPos = InStrRev(fileNameText, "/")
    lastBackslashPos = InStrRev(fileNameText, "\")
    If lastSlashPos > lastBackslashPos Then
        lastSeparatorPos = lastSlashPos
    Else
        lastSeparatorPos = lastBackslashPos
    End If

    lastDotPos = InStrRev(fileNameText, ".")
    If lastDotPos > (lastSeparatorPos + 1) Then
        RemoveExtension = Left$(fileNameText, lastDotPos - 1)
    Else
        RemoveExtension = fileNameText
    End If
End Function






