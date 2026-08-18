Attribute VB_Name = "BetaEvidenceGenerator"
Option Explicit

' ============================================================
' 単体テストエビデンス シート生成マクロ
' ------------------------------------------------------------
' このモジュールは、マクロブック（ThisWorkbook）にある雛形シートを使って、
' ユーザーが選択した参照元ブック（xlsx）から出力先ブックを決定し、そこへエビデンスシートを生成する
'
' 主な流れ
' 1. 参照元xlsxファイルを選択する
' 2. 入力ファイル名（例: foo.php）を入力する
' 3. マクロブックのREFERシートを参照して referValue を取得する
' 4. 参照元ブックの【共通】/【個別】参照元シートを走査する
' 5. 共通/個別それぞれの想定出力xlsxを決定し、既存ブックを再利用できる場合はそこへ追加する
'    同名シート衝突がある場合は連番付きの新規ブックへ丸ごと出力する
' 6. 出力対象シート名を任意入力し、空欄なら全シートを出力する
' 7. 共通モードでは A1-1-1（指定時のみ）を独立シートとして出力する
' 8. エビデンスシートは A1 テンプレを複製して作成する
' 9. A1-1-1 の A3/B3 の ○○○ を baseName に置換し、E/H列を規則で書き込む
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
Private Const EVIDENCE_NEW_SIDE_FIRST_COL As Long = 3 ' 新側の先頭列（C列）
Private Const EVIDENCE_OLD_SIDE_COL_COUNT As Long = 15 ' 旧側の比較列数
Private Const DEFAULT_BORDER_END_COL As Long = 32 ' 既定構成では AF列（A:B + 新15 + 旧15）
Public Const OPTION_RIGHT_BORDER_ENABLED As Boolean = True ' True: 右罫線を適用 / False: 右罫線を適用しない
Private Const DEFAULT_RIGHT_BORDER_TARGET_COL As Long = 17 ' 既定構成では Q列が新旧の境界
Private Const RIGHT_BORDER_EXTRA_ROWS As Long = SLOT_HEIGHT ' 行オフセット未指定時の既定延長行数
Private Const EVIDENCE_HEADER_ROW As Long = 2 ' 雛形ヘッダ行
Private Const EVIDENCE_OLD_HEADER_COL As Long = 18 ' R2 = 旧
Private Const EVIDENCE_COMPARE_BASE_COL_COUNT As Long = 15 ' 雛形既定の比較列数
Private Const EVIDENCE_COMPARE_MIN_COL_COUNT As Long = 3 ' 新側を増減する際の最小列数
Private Const EVIDENCE_NEW_SIDE_ADJUST_COL As Long = 4 ' 新側は D列をテンプレ/増減起点に使う
Private Const EVIDENCE_COLUMN_LAYOUT_SCOPE_NEWONLY As String = "NewOnly"
Private Const EVIDENCE_COLUMN_LAYOUT_SCOPE_BOTH As String = "Both"
Public Const OPTION_CLEAR_OLD_HEADER_TEXT_ENABLED As Boolean = False ' True: A1複製シートの R2「旧」を消す / False: 残す
Private Const OPTION_EVIDENCE_NEW_SIDE_COL_COUNT As Long = EVIDENCE_COMPARE_BASE_COL_COUNT ' 列数オプション。Both 指定時は旧側にも同じ列数を適用する
Private Const OPTION_EVIDENCE_COLUMN_LAYOUT_SCOPE As String = EVIDENCE_COLUMN_LAYOUT_SCOPE_NEWONLY ' NewOnly: 新側だけ増減 / Both: 新旧両側を同数で増減
Private Const EXCEL_MAX_COLUMN_COUNT As Long = 16384

' ===== 共通モード先頭シートのヘッダ置換 =====
Private Const HEADER_PLACEHOLDER As String = "○○○"

' ===== Office定数を数値で扱う（参照設定に依存しにくくするため） =====
Private Const FILE_DIALOG_PICKER As Long = 3 ' msoFileDialogFilePicker
Public Const OPTION_TOP_BORDER_ENABLED As Boolean = True ' True: A列から比較領域終端まで上罫線を適用 / False: 上罫線を適用しない
Public Const OPTION_SLOT_HEIGHT_PROMPT_ENABLED As Boolean = False ' True: 行オフセット入力を表示 / False: SLOT_HEIGHTを使用
Public Const OPTION_OUTPUT_SHEET_SELECTION_PROMPT_ENABLED As Boolean = False ' True: 作成シート選択入力を表示 / False: 全シート出力
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
    sourceWorkbookPath As String
    inputFileName As String

    useSlotHeight As Boolean
    slotHeight As Long

    useOutputSheetFilter As Boolean
    outputSheetFilterText As String

    OverrideTopBorderEnabled As Boolean
    topBorderEnabled As Boolean

    OverrideSlotHeightPromptEnabled As Boolean
    SlotHeightPromptEnabled As Boolean

    OverrideOutputSheetSelectionPromptEnabled As Boolean
    OutputSheetSelectionPromptEnabled As Boolean

    OverrideExcludeOutputSheetByPatternEnabled As Boolean
    excludeOutputSheetByPatternEnabled As Boolean

    UseExcludedOutputSheetNamePatterns As Boolean
    excludedOutputSheetNamePatterns As String

    OverrideSkipGrayFilledSourceCellEnabled As Boolean
    skipGrayFilledSourceCellEnabled As Boolean

    UseSourceSkipFillColorHexCodes As Boolean
    sourceSkipFillColorHexCodes As String

    OverrideRightBorderEnabled As Boolean
    rightBorderEnabled As Boolean


    OverrideClearOldHeaderTextEnabled As Boolean
    clearOldHeaderTextEnabled As Boolean

    UseEvidenceNewSideColCount As Boolean
    evidenceNewSideColCount As Long

    UseEvidenceColumnLayoutScope As Boolean
    evidenceColumnLayoutScope As String
End Type

Private mUiOptions As BetaEvidenceUiOptions
Private mSlotHeight As Long ' スロット行オフセット（未指定時は既定値を使用）
Private mSkipSourceFillColorMap As Object ' 参照元塗りつぶしスキップ色マップ

' 目的: フォームまたはCONFIGから渡された設定を一時適用し、通常の実行経路でエビデンス生成を行う。
Public Sub RunMainWithUiOptions(ByRef options As BetaEvidenceUiOptions)
    ClearUiOptions
    mUiOptions = options
    mUiOptions.Enabled = True

    RunMain

    ClearUiOptions
End Sub

' 目的: フォームとCONFIGで共有するエビデンス生成の既定設定を作成する。
Public Function CreateBetaEvidenceUiOptionsForForm() As BetaEvidenceUiOptions
    Dim options As BetaEvidenceUiOptions

    InitializeBetaEvidenceUiOptionsForForm options
    CreateBetaEvidenceUiOptionsForForm = options
End Function

' 目的: 呼び出し側が用意した設定へ、ソース定数に基づく既定値を設定する。
Public Sub InitializeBetaEvidenceUiOptionsForForm(ByRef options As BetaEvidenceUiOptions)
    options.Enabled = True
    options.sourceWorkbookPath = vbNullString
    options.inputFileName = vbNullString

    options.useSlotHeight = True
    options.slotHeight = SLOT_HEIGHT

    options.useOutputSheetFilter = True
    options.outputSheetFilterText = vbNullString

    options.OverrideTopBorderEnabled = True
    options.topBorderEnabled = OPTION_TOP_BORDER_ENABLED

    options.OverrideSlotHeightPromptEnabled = True
    options.SlotHeightPromptEnabled = False

    options.OverrideOutputSheetSelectionPromptEnabled = True
    options.OutputSheetSelectionPromptEnabled = False

    options.OverrideExcludeOutputSheetByPatternEnabled = True
    options.excludeOutputSheetByPatternEnabled = OPTION_EXCLUDE_OUTPUT_SHEET_BY_PATTERN_ENABLED

    options.UseExcludedOutputSheetNamePatterns = True
    options.excludedOutputSheetNamePatterns = EXCLUDED_OUTPUT_SHEET_NAME_PATTERNS

    options.OverrideSkipGrayFilledSourceCellEnabled = True
    options.skipGrayFilledSourceCellEnabled = OPTION_SKIP_GRAY_FILLED_SOURCE_CELL_ENABLED

    options.UseSourceSkipFillColorHexCodes = True
    options.sourceSkipFillColorHexCodes = SOURCE_SKIP_FILL_COLOR_HEX_CODES

    options.OverrideRightBorderEnabled = True
    options.rightBorderEnabled = OPTION_RIGHT_BORDER_ENABLED


    options.OverrideClearOldHeaderTextEnabled = True
    options.clearOldHeaderTextEnabled = OPTION_CLEAR_OLD_HEADER_TEXT_ENABLED

    options.UseEvidenceNewSideColCount = True
    options.evidenceNewSideColCount = OPTION_EVIDENCE_NEW_SIDE_COL_COUNT

    options.UseEvidenceColumnLayoutScope = True
    options.evidenceColumnLayoutScope = OPTION_EVIDENCE_COLUMN_LAYOUT_SCOPE
End Sub

' 目的: 前回のUI設定が単体実行へ漏れないよう、モジュール保持値を初期化する。
Private Sub ClearUiOptions()
    mUiOptions.Enabled = False
    mUiOptions.sourceWorkbookPath = vbNullString
    mUiOptions.inputFileName = vbNullString

    mUiOptions.useSlotHeight = False
    mUiOptions.slotHeight = 0

    mUiOptions.useOutputSheetFilter = False
    mUiOptions.outputSheetFilterText = vbNullString

    mUiOptions.OverrideTopBorderEnabled = False
    mUiOptions.topBorderEnabled = False

    mUiOptions.OverrideSlotHeightPromptEnabled = False
    mUiOptions.SlotHeightPromptEnabled = False

    mUiOptions.OverrideOutputSheetSelectionPromptEnabled = False
    mUiOptions.OutputSheetSelectionPromptEnabled = False

    mUiOptions.OverrideExcludeOutputSheetByPatternEnabled = False
    mUiOptions.excludeOutputSheetByPatternEnabled = False

    mUiOptions.UseExcludedOutputSheetNamePatterns = False
    mUiOptions.excludedOutputSheetNamePatterns = vbNullString

    mUiOptions.OverrideSkipGrayFilledSourceCellEnabled = False
    mUiOptions.skipGrayFilledSourceCellEnabled = False

    mUiOptions.UseSourceSkipFillColorHexCodes = False
    mUiOptions.sourceSkipFillColorHexCodes = vbNullString

    mUiOptions.OverrideRightBorderEnabled = False
    mUiOptions.rightBorderEnabled = False


    mUiOptions.OverrideClearOldHeaderTextEnabled = False
    mUiOptions.clearOldHeaderTextEnabled = False

    mUiOptions.UseEvidenceNewSideColCount = False
    mUiOptions.evidenceNewSideColCount = 0

    mUiOptions.UseEvidenceColumnLayoutScope = False
    mUiOptions.evidenceColumnLayoutScope = vbNullString
End Sub

' ============================================================
' エントリポイント
' ============================================================

' 目的: 入力取得、REFER照合、共通・個別ブック生成、保存までの処理全体を統括する。
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
    Dim outputSheetFilterRaw As String
    Dim commonCreatedSheetCount As Long
    Dim individualCreatedSheetCount As Long
    Dim commonHeaderCreated As Boolean
    Dim commonPlannedSheetNameMap As Object
    Dim individualPlannedSheetNameMap As Object
    Dim commonReusedExistingWorkbook As Boolean
    Dim individualReusedExistingWorkbook As Boolean
    Dim commonOutputWorkbookWasAlreadyOpen As Boolean
    Dim individualOutputWorkbookWasAlreadyOpen As Boolean
    Dim commonFallbackToNewWorkbook As Boolean
    Dim individualFallbackToNewWorkbook As Boolean
    Dim commonConflictSheetName As String
    Dim individualConflictSheetName As String
    Dim currentTargetWorkbookWasAlreadyOpen As Boolean
    Dim currentTargetWorkbookCreatedNewFile As Boolean
    Dim currentTargetOutputPath As String

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
        If mUiOptions.useSlotHeight And mUiOptions.slotHeight > 0 Then
            mSlotHeight = mUiOptions.slotHeight
        Else
            mSlotHeight = SLOT_HEIGHT
        End If
    ElseIf IsSlotHeightPromptEnabled() Then
        mSlotHeight = PromptSlotHeightOrDefault(SLOT_HEIGHT)
    Else
        mSlotHeight = SLOT_HEIGHT ' 入力ダイアログOFF時は既定オフセットをそのまま使う
    End If

    If mUiOptions.Enabled Then
        If mUiOptions.useOutputSheetFilter Then
            outputSheetFilterRaw = mUiOptions.outputSheetFilterText
        Else
            outputSheetFilterRaw = vbNullString
        End If
    ElseIf IsOutputSheetSelectionPromptEnabled() Then
        outputSheetFilterRaw = PromptOutputSheetFilter()
    Else
        outputSheetFilterRaw = vbNullString ' 入力ダイアログOFF時は全シートを出力対象にする
    End If

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

    Set outputSheetFilter = ParseOutputSheetFilter(outputSheetFilterRaw, commonSourceWs, individualSourceWs)
    outputSheetFilterLabel = BuildOutputSheetFilterLabel(outputSheetFilter)

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
        Set commonPlannedSheetNameMap = BuildPlannedEvidenceSheetNameMap( _
            sourceWs:=commonSourceWs, _
            modeLabel:="共通", _
            outputSheetFilter:=outputSheetFilter, _
            includeCommonHeaderSheet:=IsSheetAllowedByFilter(TEMPLATE_HEADER_SHEET_NAME, outputSheetFilter))

        Set targetWb = PrepareOutputWorkbookForEvidenceMode( _
            desiredOutputPath:=BuildOutputWorkbookPath(targetPath, expectedCommonWorkbookName), _
            plannedSheetNameMap:=commonPlannedSheetNameMap, _
            seedSheetNameOut:=seedSheetName, _
            actualOutputPathOut:=commonOutputPath, _
            reusedExistingWorkbookOut:=commonReusedExistingWorkbook, _
            targetWorkbookWasAlreadyOpenOut:=commonOutputWorkbookWasAlreadyOpen, _
            fallbackToNewWorkbookOut:=commonFallbackToNewWorkbook, _
            conflictSheetNameOut:=commonConflictSheetName)

        currentTargetWorkbookWasAlreadyOpen = commonOutputWorkbookWasAlreadyOpen
        currentTargetWorkbookCreatedNewFile = Not commonReusedExistingWorkbook
        currentTargetOutputPath = commonOutputPath

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
            targetWb.Save
            If Not commonOutputWorkbookWasAlreadyOpen Then
                ActivateFirstWorksheetForOpenState targetWb
                targetWb.Close SaveChanges:=True
            End If
            Set targetWb = Nothing
            currentTargetWorkbookWasAlreadyOpen = False
            currentTargetWorkbookCreatedNewFile = False
            currentTargetOutputPath = vbNullString

            commonSummary = commonSummary & " / 出力: " & commonOutputPath & _
                            BuildOutputWorkbookUsageLabel( _
                                commonReusedExistingWorkbook, _
                                commonOutputWorkbookWasAlreadyOpen, _
                                commonFallbackToNewWorkbook, _
                                commonConflictSheetName)
            processedAnyMode = True
        Else
            ReleasePreparedOutputWorkbookWithoutSave _
                wb:=targetWb, _
                outputPath:=commonOutputPath, _
                reusedExistingWorkbook:=commonReusedExistingWorkbook, _
                workbookWasAlreadyOpen:=commonOutputWorkbookWasAlreadyOpen
            currentTargetWorkbookWasAlreadyOpen = False
            currentTargetWorkbookCreatedNewFile = False
            currentTargetOutputPath = vbNullString
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
        Set individualPlannedSheetNameMap = BuildPlannedEvidenceSheetNameMap( _
            sourceWs:=individualSourceWs, _
            modeLabel:="個別", _
            outputSheetFilter:=outputSheetFilter)

        Set targetWb = PrepareOutputWorkbookForEvidenceMode( _
            desiredOutputPath:=BuildOutputWorkbookPath(targetPath, expectedIndividualWorkbookName), _
            plannedSheetNameMap:=individualPlannedSheetNameMap, _
            seedSheetNameOut:=seedSheetName, _
            actualOutputPathOut:=individualOutputPath, _
            reusedExistingWorkbookOut:=individualReusedExistingWorkbook, _
            targetWorkbookWasAlreadyOpenOut:=individualOutputWorkbookWasAlreadyOpen, _
            fallbackToNewWorkbookOut:=individualFallbackToNewWorkbook, _
            conflictSheetNameOut:=individualConflictSheetName)

        currentTargetWorkbookWasAlreadyOpen = individualOutputWorkbookWasAlreadyOpen
        currentTargetWorkbookCreatedNewFile = Not individualReusedExistingWorkbook
        currentTargetOutputPath = individualOutputPath
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
            targetWb.Save
            If Not individualOutputWorkbookWasAlreadyOpen Then
                ActivateFirstWorksheetForOpenState targetWb
                targetWb.Close SaveChanges:=True
            End If
            Set targetWb = Nothing
            currentTargetWorkbookWasAlreadyOpen = False
            currentTargetWorkbookCreatedNewFile = False
            currentTargetOutputPath = vbNullString

            individualSummary = individualSummary & " / 出力: " & individualOutputPath & _
                                BuildOutputWorkbookUsageLabel( _
                                    individualReusedExistingWorkbook, _
                                    individualOutputWorkbookWasAlreadyOpen, _
                                    individualFallbackToNewWorkbook, _
                                    individualConflictSheetName)
            processedAnyMode = True
        Else
            ReleasePreparedOutputWorkbookWithoutSave _
                wb:=targetWb, _
                outputPath:=individualOutputPath, _
                reusedExistingWorkbook:=individualReusedExistingWorkbook, _
                workbookWasAlreadyOpen:=individualOutputWorkbookWasAlreadyOpen
            currentTargetWorkbookWasAlreadyOpen = False
            currentTargetWorkbookCreatedNewFile = False
            currentTargetOutputPath = vbNullString
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
        If currentTargetWorkbookCreatedNewFile Then
            DiscardOutputWorkbookAndFile targetWb, currentTargetOutputPath
        ElseIf currentTargetWorkbookWasAlreadyOpen Then
            Set targetWb = Nothing
        Else
            targetWb.Close SaveChanges:=False
            Set targetWb = Nothing
        End If
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

' 目的: 参照元として読み込むExcelブックをファイル選択ダイアログから取得する。
Private Function SelectTargetWorkbookPath() As String
    ' FileDialog を使って、参照元の xlsx をユーザーに選ばせる
    ' 参照設定依存を避けるため、FileDialog型ではなく Object で扱う
    Dim fd As Object

    If mUiOptions.Enabled Then
        SelectTargetWorkbookPath = Trim$(mUiOptions.sourceWorkbookPath)
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

' 目的: REFER照合と出力名に使用する対象ファイル名をユーザーから取得する。
Private Function PromptInputFileName() As String
    ' REFER検索キーになる入力ファイル名を受け取る
    ' 前後の空白は誤入力になりやすいため Trim する
    Dim s As String

    If mUiOptions.Enabled Then
        PromptInputFileName = Trim$(mUiOptions.inputFileName)
        Exit Function
    End If

    s = InputBox("入力ファイル名を入力してください（例: menu/mainmenu.php）", "入力ファイル名")
    PromptInputFileName = Trim$(s)
End Function
' 目的: オプションに応じて行オフセットを入力させ、未指定時はソース既定値を採用する。
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

' 目的: 個別名と範囲指定に対応した出力対象シート文字列をユーザーから取得する。
Private Function PromptOutputSheetFilter() As String
    Dim inputText As String

    inputText = InputBox( _
        "出力するシート名をカンマ区切りで入力してください（任意）。" & vbCrLf & _
        "範囲指定は : が使えます。" & vbCrLf & _
        "例: A1:A3 / :A2 / A3:" & vbCrLf & _
        "共通と個別をまたぐ指定（例: A1:B3）はできません。" & vbCrLf & _
        "空欄の場合は全シートを出力します。", _
        "出力シート名（任意）")

    PromptOutputSheetFilter = inputText
End Function

' 目的: カンマ区切り・コロン範囲の指定を、照合用シート名マップへ展開する。
Private Function ParseOutputSheetFilter( _
    ByVal rawInput As String, _
    Optional ByVal commonSourceWs As Worksheet = Nothing, _
    Optional ByVal individualSourceWs As Worksheet = Nothing) As Object

    Dim normalizedText As String
    Dim names As Variant
    Dim tokenText As String
    Dim i As Long
    Dim dict As Object
    Dim rangeMaxMap As Object

    normalizedText = Replace(rawInput, "，", ",")
    normalizedText = Replace(normalizedText, "：", ":")
    normalizedText = Trim$(normalizedText)
    If Len(normalizedText) = 0 Then Exit Function

    names = Split(normalizedText, ",")

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare
    Set rangeMaxMap = BuildOutputSheetRangeMaxMap(commonSourceWs, individualSourceWs)

    For i = LBound(names) To UBound(names)
        tokenText = Trim$(CStr(names(i)))
        If Len(tokenText) > 0 Then
            ExpandOutputSheetFilterToken dict, tokenText, rangeMaxMap
        End If
    Next i

    If dict.Count = 0 Then
        Set ParseOutputSheetFilter = Nothing
    Else
        Set ParseOutputSheetFilter = dict
    End If
End Function

' 目的: 開放範囲の終端を決めるため、参照元に存在する系列別の最大番号を収集する。
Private Function BuildOutputSheetRangeMaxMap( _
    Optional ByVal commonSourceWs As Worksheet = Nothing, _
    Optional ByVal individualSourceWs As Worksheet = Nothing) As Object

    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    RegisterOutputSheetRangeMaxFromSource commonSourceWs, dict
    RegisterOutputSheetRangeMaxFromSource individualSourceWs, dict

    Set BuildOutputSheetRangeMaxMap = dict
End Function

' 目的: 参照元A列のシート候補を走査し、系列別の最大番号を登録する。
Private Sub RegisterOutputSheetRangeMaxFromSource( _
    ByVal sourceWs As Worksheet, _
    ByVal rangeMaxMap As Object)

    Dim lastRowA As Long
    Dim scanEndRow As Long
    Dim sourceValuesA As Variant
    Dim rowOffset As Long
    Dim rawA As Variant
    Dim sheetName As String
    Dim prefixText As String
    Dim numericIndex As Long

    If sourceWs Is Nothing Then Exit Sub

    lastRowA = GetLastUsedRowInColumn(sourceWs, SOURCE_COL_A)
    scanEndRow = lastRowA + EMPTY_STREAK_STOP_COUNT
    sourceValuesA = ReadColumnValuesFromRow(sourceWs, SOURCE_COL_A, SOURCE_START_ROW, scanEndRow)

    For rowOffset = 1 To UBound(sourceValuesA, 1)
        rawA = sourceValuesA(rowOffset, 1)
        If HasValueForSourceCell(rawA) Then
            sheetName = Trim$(CStr(rawA))
            If TryParseSimpleSheetSeriesToken(sheetName, prefixText, numericIndex) Then
                RegisterOutputSheetRangeMax rangeMaxMap, prefixText, numericIndex
            End If
        End If
    Next rowOffset
End Sub

' 目的: 単純なシート系列名だけを対象に、既知の最大番号を更新する。
Private Sub RegisterOutputSheetRangeMax( _
    ByVal rangeMaxMap As Object, _
    ByVal prefixText As String, _
    ByVal numericIndex As Long)

    If rangeMaxMap Is Nothing Then Exit Sub
    If Len(prefixText) = 0 Then Exit Sub
    If numericIndex < 1 Then Exit Sub

    If Not rangeMaxMap.Exists(prefixText) Then
        rangeMaxMap.Add prefixText, numericIndex
    ElseIf CLng(rangeMaxMap(prefixText)) < numericIndex Then
        rangeMaxMap(prefixText) = numericIndex
    End If
End Sub

' 目的: 単一名または範囲トークンを判別し、出力対象マップへ追加する。
Private Sub ExpandOutputSheetFilterToken( _
    ByVal outputSheetFilter As Object, _
    ByVal tokenText As String, _
    ByVal rangeMaxMap As Object)

    If outputSheetFilter Is Nothing Then Exit Sub

    If InStr(1, tokenText, ":", vbBinaryCompare) > 0 Then
        ExpandOutputSheetFilterRange outputSheetFilter, tokenText, rangeMaxMap
    Else
        AddOutputSheetFilterName outputSheetFilter, tokenText
    End If
End Sub

' 目的: コロン範囲を同一系列の連続したシート名へ展開する。
Private Sub ExpandOutputSheetFilterRange( _
    ByVal outputSheetFilter As Object, _
    ByVal tokenText As String, _
    ByVal rangeMaxMap As Object)

    Dim colonPos As Long
    Dim startToken As String
    Dim endToken As String
    Dim startPrefix As String
    Dim endPrefix As String
    Dim prefixText As String
    Dim startIndex As Long
    Dim endIndex As Long
    Dim currentIndex As Long

    colonPos = InStr(1, tokenText, ":", vbBinaryCompare)
    If colonPos <= 0 Then
        AddOutputSheetFilterName outputSheetFilter, tokenText
        Exit Sub
    End If

    If InStr(colonPos + 1, tokenText, ":", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 2411, "ParseOutputSheetFilter", _
                  "出力シート範囲指定に ':' を複数含めることはできません: " & tokenText
    End If

    startToken = Trim$(Left$(tokenText, colonPos - 1))
    endToken = Trim$(Mid$(tokenText, colonPos + 1))

    If Len(startToken) = 0 And Len(endToken) = 0 Then
        Err.Raise vbObjectError + 2412, "ParseOutputSheetFilter", _
                  "出力シート範囲指定が空です: " & tokenText
    End If

    If Len(startToken) > 0 Then
        If Not TryParseSimpleSheetSeriesToken(startToken, startPrefix, startIndex) Then
            Err.Raise vbObjectError + 2413, "ParseOutputSheetFilter", _
                      "範囲指定の開始値が不正です。A1 のような形式で指定してください: " & tokenText
        End If
        prefixText = startPrefix
    End If

    If Len(endToken) > 0 Then
        If Not TryParseSimpleSheetSeriesToken(endToken, endPrefix, endIndex) Then
            Err.Raise vbObjectError + 2414, "ParseOutputSheetFilter", _
                      "範囲指定の終了値が不正です。A3 のような形式で指定してください: " & tokenText
        End If
        If Len(prefixText) = 0 Then
            prefixText = endPrefix
        End If
    End If

    If Len(startToken) > 0 And Len(endToken) > 0 Then
        If StrComp(startPrefix, endPrefix, vbBinaryCompare) <> 0 Then
            Err.Raise vbObjectError + 2415, "ParseOutputSheetFilter", _
                      "共通と個別をまたぐ範囲指定はできません。開始と終了は同じ接頭辞で指定してください: " & tokenText
        End If
    End If

    If Len(startToken) = 0 Then
        startIndex = 1
    End If

    If Len(endToken) = 0 Then
        endIndex = ResolveOutputSheetRangeLastIndex(prefixText, rangeMaxMap, tokenText)
    End If

    If endIndex < startIndex Then
        Err.Raise vbObjectError + 2416, "ParseOutputSheetFilter", _
                  "出力シート範囲の開始値が終了値を超えています: " & tokenText
    End If

    For currentIndex = startIndex To endIndex
        AddOutputSheetFilterName outputSheetFilter, prefixText & CStr(currentIndex)
    Next currentIndex
End Sub

' 目的: 範囲の右端省略時に、参照元で確認できる系列の最終番号を解決する。
Private Function ResolveOutputSheetRangeLastIndex( _
    ByVal prefixText As String, _
    ByVal rangeMaxMap As Object, _
    ByVal tokenText As String) As Long

    If rangeMaxMap Is Nothing Then
        Err.Raise vbObjectError + 2417, "ParseOutputSheetFilter", _
                  "終端省略の範囲指定の末尾を判断できませんでした: " & tokenText
    End If

    If Not rangeMaxMap.Exists(prefixText) Then
        Err.Raise vbObjectError + 2418, "ParseOutputSheetFilter", _
                  "終端省略の範囲指定に対応するシートが参照元に見つかりませんでした: " & tokenText
    End If

    ResolveOutputSheetRangeLastIndex = CLng(rangeMaxMap(prefixText))
End Function

' 目的: 空文字と重複を除外して、シート名を大文字小文字非依存のマップへ登録する。
Private Sub AddOutputSheetFilterName( _
    ByVal outputSheetFilter As Object, _
    ByVal sheetName As String)

    Dim normalizedName As String

    If outputSheetFilter Is Nothing Then Exit Sub

    normalizedName = Trim$(sheetName)
    If Len(normalizedName) = 0 Then Exit Sub

    If Not outputSheetFilter.Exists(normalizedName) Then
        outputSheetFilter.Add normalizedName, True
    End If
End Sub

' 目的: A1のような単純系列名を接頭辞と数値へ分解できるか試行する。
Private Function TryParseSimpleSheetSeriesToken( _
    ByVal tokenText As String, _
    ByRef prefixTextOut As String, _
    ByRef numericIndexOut As Long) As Boolean

    Dim normalized As String
    Dim i As Long
    Dim ch As String
    Dim numberText As String

    normalized = Trim$(tokenText)
    prefixTextOut = vbNullString
    numericIndexOut = 0

    If Len(normalized) = 0 Then Exit Function

    For i = 1 To Len(normalized)
        ch = Mid$(normalized, i, 1)
        If ch >= "0" And ch <= "9" Then Exit For
        If (ch < "A" Or ch > "Z") And (ch < "a" Or ch > "z") Then Exit Function
    Next i

    If i <= 1 Or i > Len(normalized) Then Exit Function

    prefixTextOut = UCase$(Left$(normalized, i - 1))
    numberText = Mid$(normalized, i)
    If Len(numberText) = 0 Then Exit Function

    For i = 1 To Len(numberText)
        ch = Mid$(numberText, i, 1)
        If ch < "0" Or ch > "9" Then
            prefixTextOut = vbNullString
            Exit Function
        End If
    Next i

    On Error GoTo ParseError
    numericIndexOut = CLng(numberText)
    If numericIndexOut < 1 Then GoTo ParseError

    TryParseSimpleSheetSeriesToken = True
    Exit Function

ParseError:
    prefixTextOut = vbNullString
    numericIndexOut = 0
End Function

' 目的: 処理結果メッセージに表示する出力対象の説明文を組み立てる。
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

' 目的: UI上書きを優先し、確定行の横罫線を適用するか決定する。
Private Function IsTopBorderEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideTopBorderEnabled Then
        IsTopBorderEnabled = mUiOptions.topBorderEnabled
    Else
        IsTopBorderEnabled = OPTION_TOP_BORDER_ENABLED
    End If
End Function

' 目的: UI上書きを優先し、単体実行時に行オフセット入力を表示するか決定する。
Private Function IsSlotHeightPromptEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideSlotHeightPromptEnabled Then
        IsSlotHeightPromptEnabled = mUiOptions.SlotHeightPromptEnabled
    Else
        IsSlotHeightPromptEnabled = OPTION_SLOT_HEIGHT_PROMPT_ENABLED
    End If
End Function

' 目的: UI上書きを優先し、単体実行時にシート選択入力を表示するか決定する。
Private Function IsOutputSheetSelectionPromptEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideOutputSheetSelectionPromptEnabled Then
        IsOutputSheetSelectionPromptEnabled = mUiOptions.OutputSheetSelectionPromptEnabled
    Else
        IsOutputSheetSelectionPromptEnabled = OPTION_OUTPUT_SHEET_SELECTION_PROMPT_ENABLED
    End If
End Function

' 目的: UI上書きを優先し、除外パターン判定を有効にするか決定する。
Private Function IsExcludeOutputSheetByPatternEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideExcludeOutputSheetByPatternEnabled Then
        IsExcludeOutputSheetByPatternEnabled = mUiOptions.excludeOutputSheetByPatternEnabled
    Else
        IsExcludeOutputSheetByPatternEnabled = OPTION_EXCLUDE_OUTPUT_SHEET_BY_PATTERN_ENABLED
    End If
End Function

' 目的: UI指定があればそれを使い、なければソース定数の除外パターンを返す。
Private Function GetExcludedOutputSheetNamePatterns() As String
    If mUiOptions.Enabled And mUiOptions.UseExcludedOutputSheetNamePatterns Then
        GetExcludedOutputSheetNamePatterns = CStr(mUiOptions.excludedOutputSheetNamePatterns)
    Else
        GetExcludedOutputSheetNamePatterns = EXCLUDED_OUTPUT_SHEET_NAME_PATTERNS
    End If
End Function

' 目的: UI上書きを優先し、参照元セルの塗りつぶし色による読取除外を決定する。
Private Function IsSkipGrayFilledSourceCellEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideSkipGrayFilledSourceCellEnabled Then
        IsSkipGrayFilledSourceCellEnabled = mUiOptions.skipGrayFilledSourceCellEnabled
    Else
        IsSkipGrayFilledSourceCellEnabled = OPTION_SKIP_GRAY_FILLED_SOURCE_CELL_ENABLED
    End If
End Function

' 目的: UI指定があればそれを使い、なければソース定数の読み飛ばし色一覧を返す。
Private Function GetSourceSkipFillColorHexCodes() As String
    If mUiOptions.Enabled And mUiOptions.UseSourceSkipFillColorHexCodes Then
        GetSourceSkipFillColorHexCodes = CStr(mUiOptions.sourceSkipFillColorHexCodes)
    Else
        GetSourceSkipFillColorHexCodes = SOURCE_SKIP_FILL_COLOR_HEX_CODES
    End If
End Function

' 目的: UI上書きを優先し、新旧境界の縦罫線を適用するか決定する。
Private Function IsRightBorderEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideRightBorderEnabled Then
        IsRightBorderEnabled = mUiOptions.rightBorderEnabled
    Else
        IsRightBorderEnabled = OPTION_RIGHT_BORDER_ENABLED
    End If
End Function

' 目的: UI上書きを優先し、現行側ヘッダ文字を削除するか決定する。
Private Function IsClearOldHeaderTextEnabled() As Boolean
    If mUiOptions.Enabled And mUiOptions.OverrideClearOldHeaderTextEnabled Then
        IsClearOldHeaderTextEnabled = mUiOptions.clearOldHeaderTextEnabled
    Else
        IsClearOldHeaderTextEnabled = OPTION_CLEAR_OLD_HEADER_TEXT_ENABLED
    End If
End Function

' 目的: UI指定と下限・上限を考慮して、新側に確保する比較列数を決定する。
Private Function GetEvidenceNewSideColCount() As Long
    Dim desiredCount As Long
    Dim maxCount As Long

    If mUiOptions.Enabled And mUiOptions.UseEvidenceNewSideColCount Then
        desiredCount = mUiOptions.evidenceNewSideColCount
    Else
        desiredCount = OPTION_EVIDENCE_NEW_SIDE_COL_COUNT
    End If

    If desiredCount < EVIDENCE_COMPARE_MIN_COL_COUNT Then
        desiredCount = EVIDENCE_COMPARE_MIN_COL_COUNT
    End If

    maxCount = GetEvidenceMaxNewSideColCount()
    If desiredCount > maxCount Then
        desiredCount = maxCount
    End If

    GetEvidenceNewSideColCount = desiredCount
End Function

' 目的: Excelの列上限と現行側領域を考慮し、新側列数の安全な上限を算出する。
Private Function GetEvidenceMaxNewSideColCount() As Long
    If StrComp(GetEvidenceColumnLayoutScope(), EVIDENCE_COLUMN_LAYOUT_SCOPE_BOTH, vbTextCompare) = 0 Then
        GetEvidenceMaxNewSideColCount = (EXCEL_MAX_COLUMN_COUNT - DEST_COL_B) \ 2
    Else
        GetEvidenceMaxNewSideColCount = EXCEL_MAX_COLUMN_COUNT - DEST_COL_B - EVIDENCE_OLD_SIDE_COL_COUNT
    End If

    If GetEvidenceMaxNewSideColCount < EVIDENCE_COMPARE_MIN_COL_COUNT Then
        GetEvidenceMaxNewSideColCount = EVIDENCE_COMPARE_MIN_COL_COUNT
    End If
End Function

' 目的: 設定値を正規化し、新側だけか新旧両側を変更するか決定する。
Private Function GetEvidenceColumnLayoutScope() As String
    Dim normalizedScope As String

    If mUiOptions.Enabled And mUiOptions.UseEvidenceColumnLayoutScope Then
        normalizedScope = Trim$(CStr(mUiOptions.evidenceColumnLayoutScope))
    Else
        normalizedScope = OPTION_EVIDENCE_COLUMN_LAYOUT_SCOPE
    End If

    Select Case UCase$(normalizedScope)
        Case UCase$(EVIDENCE_COLUMN_LAYOUT_SCOPE_BOTH)
            GetEvidenceColumnLayoutScope = EVIDENCE_COLUMN_LAYOUT_SCOPE_BOTH
        Case Else
            GetEvidenceColumnLayoutScope = EVIDENCE_COLUMN_LAYOUT_SCOPE_NEWONLY
    End Select
End Function

' 目的: 列構成オプションに基づき、現行側に確保する比較列数を返す。
Private Function GetEvidenceOldSideColCount() As Long
    If StrComp(GetEvidenceColumnLayoutScope(), EVIDENCE_COLUMN_LAYOUT_SCOPE_BOTH, vbTextCompare) = 0 Then
        GetEvidenceOldSideColCount = GetEvidenceNewSideColCount()
    Else
        GetEvidenceOldSideColCount = EVIDENCE_OLD_SIDE_COL_COUNT
    End If
End Function

' 目的: 変更後の新側列数から、現行側領域の先頭列を算出する。
Private Function GetEvidenceOldSideFirstCol() As Long
    GetEvidenceOldSideFirstCol = GetRightBorderTargetCol() + 1
End Function

' 目的: 現行側の列増減でテンプレートまたは削除起点に使う列を算出する。
Private Function GetEvidenceOldSideAdjustCol() As Long
    GetEvidenceOldSideAdjustCol = GetEvidenceOldSideFirstCol() + 1
End Function

' 目的: 変更後の列構成に追従する新旧境界列を算出する。
Private Function GetRightBorderTargetCol() As Long
    GetRightBorderTargetCol = ResolveEvidenceNewSideRightEdgeCol()
End Function

' 目的: 変更後の新旧列数に追従する横罫線の終端列を算出する。
Private Function GetTopBorderEndCol() As Long
    GetTopBorderEndCol = GetRightBorderTargetCol() + GetEvidenceOldSideColCount()
End Function

' 目的: 指定された新側列数から、新側領域の右端列を算出する。
Private Function ResolveEvidenceNewSideRightEdgeCol() As Long
    ResolveEvidenceNewSideRightEdgeCol = EVIDENCE_NEW_SIDE_FIRST_COL + GetEvidenceNewSideColCount() - 1

    If ResolveEvidenceNewSideRightEdgeCol < EVIDENCE_NEW_SIDE_FIRST_COL Then
        ResolveEvidenceNewSideRightEdgeCol = DEFAULT_RIGHT_BORDER_TARGET_COL
    End If
End Function

' 目的: フィルター未指定を全件扱いとし、指定時だけシート名の完全一致を判定する。
Private Function IsSheetAllowedByFilter( _
    ByVal sheetName As String, _
    ByVal outputSheetFilter As Object) As Boolean

    If outputSheetFilter Is Nothing Then
        IsSheetAllowedByFilter = True
    Else
        IsSheetAllowedByFilter = outputSheetFilter.Exists(sheetName)
    End If
End Function

' 目的: 有効なLikeパターンのいずれかにシート名が一致するか判定する。
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

' 目的: カラーコード一覧をExcel色値の辞書へ変換し、セル走査中の判定を高速化する。
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

' 目的: 色指定を検証し、比較用の小文字#RRGGBB形式へ正規化する。
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

' 目的: #RRGGBB文字列をExcelが使用するBGR形式の色値へ変換する。
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

' 目的: 有効時だけセルの実塗りつぶし色を辞書と照合し、未入力扱いにするか判定する。
Private Function ShouldSkipSourceCellByFillColor( _
    ByVal sourceWs As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal columnNumber As Long) As Boolean

    Dim colorValue As Long

    If mSkipSourceFillColorMap Is Nothing Then Exit Function

    colorValue = sourceWs.Cells(rowNumber, columnNumber).Interior.Color
    ShouldSkipSourceCellByFillColor = mSkipSourceFillColorMap.Exists(CStr(colorValue))
End Function

' 目的: 選択された参照元ブックを開き、後続処理で利用できるWorkbookを返す。
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

' 目的: REFER情報と共通・個別種別から、既定の出力ブックパスを組み立てる。
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

' 目的: 既存ファイルと開いているブックを考慮し、再利用または新規作成する出力先を決定する。
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

' 目的: フルパスを比較し、対象ブックが現在のExcelで既に開かれているか判定する。
Private Function IsWorkbookAlreadyOpen(ByVal workbookPath As String) As Boolean
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, workbookPath, vbTextCompare) = 0 Then
            IsWorkbookAlreadyOpen = True
            Exit Function
        End If
    Next wb
End Function

' 目的: 後から必要シートだけを追加できる、種シート付きの空出力ブックを作成する。
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

' 目的: 予定シートとの衝突を検査し、既存ブック再利用か新規ブック作成かを確定する。
Private Function PrepareOutputWorkbookForEvidenceMode( _
    ByVal desiredOutputPath As String, _
    ByVal plannedSheetNameMap As Object, _
    ByRef seedSheetNameOut As String, _
    ByRef actualOutputPathOut As String, _
    ByRef reusedExistingWorkbookOut As Boolean, _
    ByRef targetWorkbookWasAlreadyOpenOut As Boolean, _
    ByRef fallbackToNewWorkbookOut As Boolean, _
    ByRef conflictSheetNameOut As String) As Workbook

    Dim wb As Workbook

    If Len(Trim$(desiredOutputPath)) = 0 Then
        Err.Raise vbObjectError + 2013, "PrepareOutputWorkbookForEvidenceMode", "出力先パスが空です。"
    End If

    seedSheetNameOut = vbNullString
    actualOutputPathOut = desiredOutputPath
    reusedExistingWorkbookOut = False
    targetWorkbookWasAlreadyOpenOut = False
    fallbackToNewWorkbookOut = False
    conflictSheetNameOut = vbNullString

    If Len(Dir$(desiredOutputPath)) = 0 And Not IsWorkbookAlreadyOpen(desiredOutputPath) Then
        Set PrepareOutputWorkbookForEvidenceMode = CreateEmptyOutputWorkbook(desiredOutputPath, seedSheetNameOut)
        Exit Function
    End If

    targetWorkbookWasAlreadyOpenOut = IsWorkbookAlreadyOpen(desiredOutputPath)
    Set wb = OpenTargetWorkbook(desiredOutputPath, False)
    If wb Is Nothing Then
        Err.Raise vbObjectError + 2014, "PrepareOutputWorkbookForEvidenceMode", _
                  "既存の出力ブックを開けませんでした。" & vbCrLf & desiredOutputPath
    End If

    conflictSheetNameOut = FindFirstConflictingSheetName(wb, plannedSheetNameMap)
    If Len(conflictSheetNameOut) > 0 Then
        If Not targetWorkbookWasAlreadyOpenOut Then
            wb.Close SaveChanges:=False
        End If
        Set wb = Nothing

        actualOutputPathOut = ResolveOutputWorkbookPath(desiredOutputPath)
        Set PrepareOutputWorkbookForEvidenceMode = CreateEmptyOutputWorkbook(actualOutputPathOut, seedSheetNameOut)
        fallbackToNewWorkbookOut = True
        targetWorkbookWasAlreadyOpenOut = False
        Exit Function
    End If

    reusedExistingWorkbookOut = True
    Set PrepareOutputWorkbookForEvidenceMode = wb
End Function

' 目的: 作成予定名と既存シート名を照合し、最初の競合名を返す。
Private Function FindFirstConflictingSheetName( _
    ByVal wb As Workbook, _
    ByVal plannedSheetNameMap As Object) As String

    Dim key As Variant

    If wb Is Nothing Then Exit Function
    If plannedSheetNameMap Is Nothing Then Exit Function

    For Each key In plannedSheetNameMap.Keys
        If Not FindWorksheetExact(wb, CStr(key)) Is Nothing Then
            FindFirstConflictingSheetName = CStr(key)
            Exit Function
        End If
    Next key
End Function

' 目的: 完了メッセージ向けに、既存追記か新規作成かを表す文言を組み立てる。
Private Function BuildOutputWorkbookUsageLabel( _
    ByVal reusedExistingWorkbook As Boolean, _
    ByVal workbookWasAlreadyOpen As Boolean, _
    ByVal fallbackToNewWorkbook As Boolean, _
    ByVal conflictSheetName As String) As String

    If fallbackToNewWorkbook Then
        BuildOutputWorkbookUsageLabel = "（既存ブックに同名シートがあるため新規ブックへ出力: " & conflictSheetName & "）"
    ElseIf reusedExistingWorkbook Then
        If workbookWasAlreadyOpen Then
            BuildOutputWorkbookUsageLabel = "（既存ブックへ追加: 開いているブックを再利用）"
        Else
            BuildOutputWorkbookUsageLabel = "（既存ブックへ追加）"
        End If
    Else
        BuildOutputWorkbookUsageLabel = vbNullString
    End If
End Function

' 目的: 出力を行わなかった新規ブックだけを、保存せず安全に閉じる。
Private Sub ReleasePreparedOutputWorkbookWithoutSave( _
    ByRef wb As Workbook, _
    ByVal outputPath As String, _
    ByVal reusedExistingWorkbook As Boolean, _
    ByVal workbookWasAlreadyOpen As Boolean)

    If wb Is Nothing Then Exit Sub

    If reusedExistingWorkbook Then
        If Not workbookWasAlreadyOpen Then
            wb.Close SaveChanges:=False
        End If
        Set wb = Nothing
    Else
        DiscardOutputWorkbookAndFile wb, outputPath
    End If
End Sub

' 目的: 有効シートを残すための種シートを、実シート作成後に削除する。
Private Sub RemoveSeedSheetIfNeeded(ByVal wb As Workbook, ByVal seedSheetName As String)
    If wb Is Nothing Then Exit Sub
    If Len(seedSheetName) = 0 Then Exit Sub
    If wb.Worksheets.Count <= 1 Then Exit Sub

    DeleteWorksheetIfExists wb, seedSheetName
End Sub

' 目的: 保存後に開いたとき先頭シートが選択されるよう、ウィンドウ状態を整える。
Private Sub ActivateFirstWorksheetForOpenState(ByVal wb As Workbook)
    Dim firstWs As Worksheet

    If wb Is Nothing Then Exit Sub
    If wb.Worksheets.Count = 0 Then Exit Sub

    Set firstWs = wb.Worksheets(1)
    firstWs.Activate
End Sub


' 目的: 失敗時に新規作成したブックを閉じ、途中生成ファイルも残さないよう破棄する。
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

' 目的: 共通モード用A1-1-1を複製し、対象名置換済みの先頭シートとして作成する。
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

' 目的: 対象ファイル名をREFERで完全一致検索し、出力判定に使う値を取得する。
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
    referValueRaw = referWs.Cells(matchedRow, valueColIndex).value

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

' 目的: REFERのα・β情報から、共通用と個別用の出力ブック名を組み立てる。
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
        referWs.Cells(matchedRow, alphaColIndex).value, _
        "BuildEvidenceWorkbookNamesFromRefer", _
        "REFERシートの" & REFER_ALPHA_COL_LETTER & "列（α）"))
    If Len(alphaText) = 0 Then
        Err.Raise vbObjectError + 2133, "BuildEvidenceWorkbookNamesFromRefer", _
                  "REFERシートの" & REFER_ALPHA_COL_LETTER & "列（α）から拡張子なし文字列を取得できませんでした。"
    End If

    betaRaw = referWs.Cells(matchedRow, betaColIndex).value
    If IsError(betaRaw) Then
        Err.Raise vbObjectError + 2134, "BuildEvidenceWorkbookNamesFromRefer", _
                  "REFERシートの" & REFER_BETA_COL_LETTER & "列（β）にエラー値が入っています。"
    End If
    betaText = ToTwoDigitStringStrict(betaRaw, "REFERシートの" & REFER_BETA_COL_LETTER & "列（β）")

    gammaText = GetTrimmedCellStringOrRaise( _
        referWs.Cells(matchedRow, gammaColIndex).value, _
        "BuildEvidenceWorkbookNamesFromRefer", _
        "REFERシートの" & REFER_VALUE_COL_LETTER & "列（γ）")

    commonWorkbookName = alphaText & "_【共通】" & betaText & gammaText & "_単体テストエビデンス_初期開発.xlsx"
    individualWorkbookName = alphaText & "_【個別】" & betaText & gammaText & "_単体テストエビデンス_初期開発.xlsx"
End Sub

' 目的: 生成予定ブック名と既存有無を、実行前確認用のメッセージへ整形する。
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

' 目的: 指定列を配列走査し、文字列が完全一致する最初の行番号を返す。
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
        cellValue = ws.Cells(r, targetColIndex).value

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

' 目的: 列記号をCellsで安全に使える1始まりの列番号へ変換する。
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

' 目的: 必須セルを文字列として読み取り、空欄やエラー値なら位置付きで中断する。
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

' 目的: REFERのβ値を検証し、ブック名に使う2桁数字へ整形する。
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

' 目的: シート名とセル番地を含む、入力不正エラー共通の位置情報を組み立てる。
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

' 目的: 大文字小文字を区別せず、指定名と完全一致するワークシートを検索する。
Private Function FindWorksheetExact(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    ' シート名完全一致で取得します。見つからない場合は Nothing を返す
    On Error Resume Next
    Set FindWorksheetExact = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

' 目的: 対象ファイル名に対応する個別参照元シートを、候補規則に従って検索する。
Private Function FindIndividualSourceSheet(ByVal targetWb As Workbook, ByVal referValue As String) As Worksheet
    ' 個別モードの参照元: 【個別】referValue（完全一致）


    Dim candidateName As String

    candidateName = "【個別】" & referValue
    Set FindIndividualSourceSheet = FindWorksheetExact(targetWb, candidateName)



End Function

' ============================================================
' 参照元シート -> エビデンスシート生成
' ============================================================

' 目的: 参照元と各種除外条件から、事前衝突検査に使う作成予定シート名を収集する。
Private Function BuildPlannedEvidenceSheetNameMap( _
    ByVal sourceWs As Worksheet, _
    ByVal modeLabel As String, _
    ByVal outputSheetFilter As Object, _
    Optional ByVal includeCommonHeaderSheet As Boolean = False) As Object

    Dim plannedSheetNameMap As Object
    Dim maxRowA As Long
    Dim maxRowB As Long
    Dim maxRowC As Long
    Dim scanEndRow As Long
    Dim sourceValuesA As Variant
    Dim sourceValuesB As Variant
    Dim sourceValuesC As Variant
    Dim rowOffset As Long
    Dim r As Long
    Dim rawA As Variant
    Dim rawB As Variant
    Dim rawC As Variant
    Dim hasA As Boolean
    Dim hasB As Boolean
    Dim hasC As Boolean
    Dim emptyStreak As Long
    Dim aSheetName As String

    Set plannedSheetNameMap = CreateObject("Scripting.Dictionary")
    plannedSheetNameMap.CompareMode = vbTextCompare

    If includeCommonHeaderSheet Then
        plannedSheetNameMap(TEMPLATE_HEADER_SHEET_NAME) = True
    End If

    maxRowA = GetLastUsedRowInColumn(sourceWs, SOURCE_COL_A)
    maxRowB = GetLastUsedRowInColumn(sourceWs, SOURCE_COL_B)
    maxRowC = GetLastUsedRowInColumn(sourceWs, SOURCE_COL_C)

    scanEndRow = maxRowA
    If maxRowB > scanEndRow Then scanEndRow = maxRowB
    If maxRowC > scanEndRow Then scanEndRow = maxRowC
    If scanEndRow < SOURCE_START_ROW Then scanEndRow = SOURCE_START_ROW

    scanEndRow = scanEndRow + EMPTY_STREAK_STOP_COUNT
    sourceValuesA = ReadColumnValuesFromRow(sourceWs, SOURCE_COL_A, SOURCE_START_ROW, scanEndRow)
    sourceValuesB = ReadColumnValuesFromRow(sourceWs, SOURCE_COL_B, SOURCE_START_ROW, scanEndRow)
    sourceValuesC = ReadColumnValuesFromRow(sourceWs, SOURCE_COL_C, SOURCE_START_ROW, scanEndRow)

    emptyStreak = 0

    For rowOffset = 1 To UBound(sourceValuesA, 1)
        r = SOURCE_START_ROW + rowOffset - 1

        rawA = sourceValuesA(rowOffset, 1)
        rawB = sourceValuesB(rowOffset, 1)
        rawC = sourceValuesC(rowOffset, 1)
        If ShouldSkipSourceCellByFillColor(sourceWs, r, SOURCE_COL_A) Then rawA = vbNullString
        If ShouldSkipSourceCellByFillColor(sourceWs, r, SOURCE_COL_B) Then rawB = vbNullString
        If ShouldSkipSourceCellByFillColor(sourceWs, r, SOURCE_COL_C) Then rawC = vbNullString

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

        If hasA Then
            aSheetName = NormalizeEvidenceSheetName(rawA, sourceWs.Name, r)

            If StrComp(modeLabel, "共通", vbBinaryCompare) = 0 And _
               StrComp(aSheetName, TEMPLATE_HEADER_SHEET_NAME, vbBinaryCompare) = 0 Then
            ElseIf IsExcludedByOutputSheetPattern(aSheetName) Then
            ElseIf Not IsSheetAllowedByFilter(aSheetName, outputSheetFilter) Then
            ElseIf Not plannedSheetNameMap.Exists(aSheetName) Then
                plannedSheetNameMap.Add aSheetName, True
            End If
        End If

        If emptyStreak >= EMPTY_STREAK_STOP_COUNT Then Exit For
    Next rowOffset

    Set BuildPlannedEvidenceSheetNameMap = plannedSheetNameMap
End Function

' 目的: 参照元A・E・H列を一度走査し、シート切替とスロット書き込みを状態管理する。
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
                FinalizeEvidenceSheetBorders currentEvidenceWs
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

                ApplyRightBorderToConfiguredColumn currentEvidenceWs, FIRST_DEST_ROW

                ' シートが変わったら、スロットと pendingB を新しいシート用に初期化する
                slotIndex = 0
                hasPendingB = False

                ' 共通モードの先頭1シートのみ、○○○ を baseName に置換する
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
        FinalizeEvidenceSheetBorders currentEvidenceWs
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

' 目的: 同名シートを置き換え、指定テンプレートからエビデンスシートを再作成する。
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

    If StrComp(templateSourceWs.Name, TEMPLATE_BODY_SHEET_NAME, vbBinaryCompare) = 0 Then
        ConfigureEvidenceBodySheetLayout RecreateEvidenceSheetFromTemplate
    End If
    Exit Function

RenameError:
    Err.Raise vbObjectError + 2202, "RecreateEvidenceSheetFromTemplate", _
              "エビデンスシート名を設定できませんでした: " & newSheetName & vbCrLf & _
              "（シート名の文字数・使用禁止文字・重複を確認してください）"
End Function

' 目的: 再生成前の同名ワークシートが存在する場合だけ、確認なしで削除する。
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

' 目的: A1-1-1内の置換対象セルだけを走査し、○○○を対象名へ置換する。
Private Sub ReplaceHeaderPlaceholderInSheet( _
    ByVal evidenceWs As Worksheet, _
    ByVal baseName As String)

    ' 共通モード先頭シートの A3/B3 のみを置換対象にする
    Dim cellA3 As Range
    Dim cellB3 As Range

    Set cellA3 = evidenceWs.Range("A3")
    Set cellB3 = evidenceWs.Range("B3")

    If Not IsError(cellA3.value) Then
        cellA3.value = Replace(CStr(cellA3.value), HEADER_PLACEHOLDER, baseName, 1, -1, vbTextCompare)
    End If

    If Not IsError(cellB3.value) Then
        cellB3.value = Replace(CStr(cellB3.value), HEADER_PLACEHOLDER, baseName, 1, -1, vbTextCompare)
    End If
End Sub

' 目的: A1複製直後に列構成、現行ラベル、罫線の初期状態をまとめて整える。
Private Sub ConfigureEvidenceBodySheetLayout(ByVal evidenceWs As Worksheet)
    If evidenceWs Is Nothing Then Exit Sub

    If IsClearOldHeaderTextEnabled() Then
        evidenceWs.Cells(EVIDENCE_HEADER_ROW, EVIDENCE_OLD_HEADER_COL).value = vbNullString
    End If

    AdjustEvidenceBodyColumnLayout evidenceWs
End Sub

' 目的: 設定された適用範囲に従い、新側と必要なら現行側の列数を変更する。
Private Sub AdjustEvidenceBodyColumnLayout(ByVal evidenceWs As Worksheet)
    Dim desiredNewSideColCount As Long
    Dim layoutScope As String

    desiredNewSideColCount = GetEvidenceNewSideColCount()
    layoutScope = GetEvidenceColumnLayoutScope()

    ResizeEvidenceNewSideColumns evidenceWs, desiredNewSideColCount

    If StrComp(layoutScope, EVIDENCE_COLUMN_LAYOUT_SCOPE_BOTH, vbTextCompare) = 0 Then
        ResizeEvidenceOldSideColumns evidenceWs, desiredNewSideColCount
    End If
End Sub

' 目的: 雛形15列との差分に応じて、新側領域を拡張または縮小する。
Private Sub ResizeEvidenceNewSideColumns( _
    ByVal evidenceWs As Worksheet, _
    ByVal desiredNewSideColCount As Long)

    If desiredNewSideColCount < EVIDENCE_COMPARE_BASE_COL_COUNT Then
        ShrinkEvidenceNewSideColumns evidenceWs, desiredNewSideColCount
    ElseIf desiredNewSideColCount > EVIDENCE_COMPARE_BASE_COL_COUNT Then
        ExpandEvidenceNewSideColumns evidenceWs, desiredNewSideColCount
    End If
End Sub

' 目的: 新旧境界を崩さない順序で、新側の不要列を削除する。
Private Sub ShrinkEvidenceNewSideColumns( _
    ByVal evidenceWs As Worksheet, _
    ByVal desiredNewSideColCount As Long)

    Dim deleteCount As Long
    Dim i As Long

    If evidenceWs Is Nothing Then Exit Sub

    deleteCount = EVIDENCE_COMPARE_BASE_COL_COUNT - desiredNewSideColCount
    If deleteCount <= 0 Then Exit Sub

    For i = 1 To deleteCount
        evidenceWs.Columns(EVIDENCE_NEW_SIDE_ADJUST_COL).Delete
    Next i
End Sub

' 目的: D列の書式・幅・数式をテンプレートとして、新側へ必要列を挿入する。
Private Sub ExpandEvidenceNewSideColumns( _
    ByVal evidenceWs As Worksheet, _
    ByVal desiredNewSideColCount As Long)

    Dim insertCount As Long
    Dim i As Long

    If evidenceWs Is Nothing Then Exit Sub

    insertCount = desiredNewSideColCount - EVIDENCE_COMPARE_BASE_COL_COUNT
    If insertCount <= 0 Then Exit Sub

    For i = 1 To insertCount
        evidenceWs.Columns(EVIDENCE_NEW_SIDE_ADJUST_COL).Copy
        evidenceWs.Columns(EVIDENCE_NEW_SIDE_ADJUST_COL).Insert Shift:=xlToRight
    Next i

    Application.CutCopyMode = False
End Sub

' 目的: 雛形15列との差分に応じて、現行側領域を拡張または縮小する。
Private Sub ResizeEvidenceOldSideColumns( _
    ByVal evidenceWs As Worksheet, _
    ByVal desiredOldSideColCount As Long)

    If desiredOldSideColCount < EVIDENCE_OLD_SIDE_COL_COUNT Then
        ShrinkEvidenceOldSideColumns evidenceWs, desiredOldSideColCount
    ElseIf desiredOldSideColCount > EVIDENCE_OLD_SIDE_COL_COUNT Then
        ExpandEvidenceOldSideColumns evidenceWs, desiredOldSideColCount
    End If
End Sub

' 目的: 右側レイアウトへの影響を抑えながら、現行側の不要列を削除する。
Private Sub ShrinkEvidenceOldSideColumns( _
    ByVal evidenceWs As Worksheet, _
    ByVal desiredOldSideColCount As Long)

    Dim deleteCount As Long
    Dim i As Long
    Dim adjustCol As Long

    If evidenceWs Is Nothing Then Exit Sub

    deleteCount = EVIDENCE_OLD_SIDE_COL_COUNT - desiredOldSideColCount
    If deleteCount <= 0 Then Exit Sub

    adjustCol = GetEvidenceOldSideAdjustCol()
    For i = 1 To deleteCount
        evidenceWs.Columns(adjustCol).Delete
    Next i
End Sub

' 目的: 現行側先頭列をテンプレートとして、現行側へ必要列を挿入する。
Private Sub ExpandEvidenceOldSideColumns( _
    ByVal evidenceWs As Worksheet, _
    ByVal desiredOldSideColCount As Long)

    Dim insertCount As Long
    Dim i As Long
    Dim adjustCol As Long

    If evidenceWs Is Nothing Then Exit Sub

    insertCount = desiredOldSideColCount - EVIDENCE_OLD_SIDE_COL_COUNT
    If insertCount <= 0 Then Exit Sub

    adjustCol = GetEvidenceOldSideAdjustCol()
    For i = 1 To insertCount
        evidenceWs.Columns(adjustCol).Copy
        evidenceWs.Columns(adjustCol).Insert Shift:=xlToRight
    Next i

    Application.CutCopyMode = False
End Sub

' ============================================================
' 参照元 A/E/H の読み取り補助
' ============================================================

' 目的: 参照元セルのエラー値を見逃さず、シート・行・列を示して処理を中断する。
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

' 目的: Empty・空文字・空白だけを未入力とし、参照元セルに実値があるか判定する。
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

' 目的: 書式だけのセルに影響されず、指定列で値または数式がある最終行を取得する。
Private Function GetLastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnIndex As Long) As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
    If lastRow < 1 Then
        lastRow = 1
    End If

    GetLastUsedRowInColumn = lastRow
End Function

' 目的: セル単位アクセスを避けるため、指定列の走査範囲を二次元配列で読み込む。
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

    rawValues = ws.Range(ws.Cells(startRow, columnIndex), ws.Cells(endRow, columnIndex)).value

    If startRow = endRow Then
        singleCell(1, 1) = rawValues
        ReadColumnValuesFromRow = singleCell
    Else
        ReadColumnValuesFromRow = rawValues
    End If
End Function

' 目的: 参照元A列の値をシート名として整形し、不正な空欄を位置付きで拒否する。
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

' 目的: Excelの禁止文字・長さ・重複条件に照らしてシート名を検証する。
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

' 目的: 全スロット確定後に既存罫線を整理し、横罫線・縦罫線・下端の蓋を再構成する。
Private Sub FinalizeEvidenceSheetBorders(ByVal destWs As Worksheet)
    Dim targetRow As Long
    Dim lastConfirmedRow As Long
    Dim slotHeightForWrite As Long

    If destWs Is Nothing Then Exit Sub

    slotHeightForWrite = GetSlotHeightForWrite()
    lastConfirmedRow = GetLastConfirmedDestRow(destWs)

    If slotHeightForWrite <= 0 Then Exit Sub

    ResetEvidenceTopBorders destWs, lastConfirmedRow, slotHeightForWrite
    ApplyRightBorderToConfiguredColumn destWs, lastConfirmedRow

    For targetRow = FIRST_DEST_ROW + slotHeightForWrite To lastConfirmedRow Step slotHeightForWrite
        ApplyTopBorderToConfirmedRow destWs, targetRow
    Next targetRow
End Sub

' 目的: 雛形由来の余分な横罫線を消し、今回のスロット境界だけを引き直せる状態にする。
Private Sub ResetEvidenceTopBorders( _
    ByVal destWs As Worksheet, _
    ByVal lastConfirmedRow As Long, _
    ByVal slotHeightForWrite As Long)

    Dim targetRow As Long
    Dim borderRange As Range

    If destWs Is Nothing Then Exit Sub
    If Not IsTopBorderEnabled() Then Exit Sub
    If slotHeightForWrite <= 0 Then Exit Sub
    If lastConfirmedRow < FIRST_DEST_ROW + slotHeightForWrite Then Exit Sub

    For targetRow = FIRST_DEST_ROW + slotHeightForWrite To lastConfirmedRow Step slotHeightForWrite
        Set borderRange = destWs.Range( _
            destWs.Cells(targetRow, DEST_COL_A), _
            destWs.Cells(targetRow, GetTopBorderEndCol()))

        With borderRange.Borders(xlEdgeTop)
            .LineStyle = xlNone
        End With
    Next targetRow
End Sub

' 目的: 入力値が有効な場合だけそれを使い、書き込み行間隔を安全な値で返す。
Private Function GetSlotHeightForWrite() As Long
    GetSlotHeightForWrite = mSlotHeight
    If GetSlotHeightForWrite <= 0 Then
        GetSlotHeightForWrite = SLOT_HEIGHT
    End If
End Function
' 目的: 未確定のE列値をシート切替前にB単独スロットとして確定する。
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

' 目的: E列とH列がそろったケースを同じスロットのA・B列へ書き込む。
Private Sub WritePairSlot( _
    ByVal destWs As Worksheet, _
    ByVal slotIndex As Long, _
    ByVal pendingB As Variant, _
    ByVal cValue As Variant)

    ' ペア書き込みは、どのスロットでも A/B 列に固定する
    Dim destRow As Long

    destRow = GetDestRowForSlot(slotIndex)

    destWs.Cells(destRow, DEST_COL_A).value = pendingB
    destWs.Cells(destRow, DEST_COL_B).value = cValue
    ApplyTopBorderToConfirmedRow destWs, destRow
    ApplyRightBorderToConfiguredColumn destWs, destRow
End Sub

' 目的: H列だけで確定したケースをA列へ書き込み、対応する境界罫線を適用する。
Private Sub WriteCOnlySlot( _
    ByVal destWs As Worksheet, _
    ByVal slotIndex As Long, _
    ByVal cValue As Variant)

    ' C単体は、どのスロットでも B列に書き込む
    Dim destRow As Long

    destRow = GetDestRowForSlot(slotIndex)
    destWs.Cells(destRow, DEST_COL_B).value = cValue
    ApplyTopBorderToConfirmedRow destWs, destRow
    ApplyRightBorderToConfiguredColumn destWs, destRow
End Sub

' 目的: 保留されていたE列だけのケースをA列へ書き込み、対応する境界罫線を適用する。
Private Sub WriteBOnlySlot( _
    ByVal destWs As Worksheet, _
    ByVal slotIndex As Long, _
    ByVal bValue As Variant)

    ' B単体は、どのスロットでも A列に書き込む
    Dim destRow As Long

    destRow = GetDestRowForSlot(slotIndex)
    destWs.Cells(destRow, DEST_COL_A).value = bValue
    ApplyTopBorderToConfirmedRow destWs, destRow
    ApplyRightBorderToConfiguredColumn destWs, destRow
End Sub

' 目的: 雛形の残存罫線を消してから、変更後境界列へ必要長の縦罫線を引き直す。
Private Sub ApplyRightBorderToConfiguredColumn( _
    ByVal destWs As Worksheet, _
    ByVal lastWrittenRow As Long)

    Dim endRow As Long
    Dim targetCol As Long
    Dim cleanupEndRow As Long
    Dim borderRange As Range

    If destWs Is Nothing Then Exit Sub

    targetCol = GetRightBorderTargetCol()
    endRow = ResolveRightBorderEndRow(destWs, lastWrittenRow)
    cleanupEndRow = ResolveRightBorderCleanupEndRow(destWs, endRow)

    Set borderRange = destWs.Range( _
        destWs.Cells(FIRST_DEST_ROW, targetCol), _
        destWs.Cells(cleanupEndRow, targetCol))

    With borderRange.Borders(xlEdgeRight)
        .LineStyle = xlNone
    End With

    Set borderRange = destWs.Range( _
        destWs.Cells(FIRST_DEST_ROW, targetCol), _
        destWs.Cells(endRow, targetCol))

    With borderRange.Borders(xlEdgeRight)
        If IsRightBorderEnabled() Then
            .LineStyle = xlContinuous
            .Weight = xlThin
        Else
            .LineStyle = xlNone
        End If
    End With
    ApplyBottomBorderClosure destWs, lastWrittenRow, endRow
End Sub

' 目的: 最終確定行と行オフセットから、縦罫線および下端罫線の終端行を決定する。
Private Function ResolveRightBorderEndRow( _
    ByVal destWs As Worksheet, _
    ByVal lastWrittenRow As Long) As Long

    Dim endRow As Long
    Dim confirmedLastRow As Long

    endRow = lastWrittenRow
    If endRow < FIRST_DEST_ROW Then
        endRow = FIRST_DEST_ROW
    End If

    confirmedLastRow = GetLastConfirmedDestRow(destWs)
    If confirmedLastRow > endRow Then
        endRow = confirmedLastRow
    End If

    ResolveRightBorderEndRow = endRow + GetConfiguredBorderExtensionRows() - 1
    If ResolveRightBorderEndRow < endRow Then ResolveRightBorderEndRow = endRow
End Function

' 目的: 過去の長い縦罫線も消せるよう、今回終端と使用済み最終行の大きい方を返す。
Private Function ResolveRightBorderCleanupEndRow( _
    ByVal destWs As Worksheet, _
    ByVal resolvedEndRow As Long) As Long

    Dim cleanupEndRow As Long
    Dim usedRangeLastRow As Long

    cleanupEndRow = resolvedEndRow + GetSlotHeightForWrite()
    If cleanupEndRow < resolvedEndRow Then cleanupEndRow = resolvedEndRow

    usedRangeLastRow = GetWorksheetUsedLastRow(destWs)
    If usedRangeLastRow > cleanupEndRow Then
        cleanupEndRow = usedRangeLastRow
    End If

    ResolveRightBorderCleanupEndRow = cleanupEndRow
End Function

' 目的: 値・数式・書式を含む使用範囲から、罫線清掃に必要な最終行を取得する。
Private Function GetWorksheetUsedLastRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    On Error GoTo 0

    If lastRow < FIRST_DEST_ROW Then
        lastRow = FIRST_DEST_ROW
    End If

    GetWorksheetUsedLastRow = lastRow
End Function

' 目的: ユーザー指定の行オフセットを、最終確定行より下へ延長する行数として返す。
Private Function GetConfiguredBorderExtensionRows() As Long
    If mSlotHeight > 0 Then
        GetConfiguredBorderExtensionRows = mSlotHeight
    Else
        GetConfiguredBorderExtensionRows = RIGHT_BORDER_EXTRA_ROWS
    End If
End Function

' 目的: A列とB列の値を比較し、実際に確定した最終書き込み行を取得する。
Private Function GetLastConfirmedDestRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    Dim lastRowA As Long
    Dim lastRowB As Long

    lastRow = FIRST_DEST_ROW
    lastRowA = GetLastNonEmptyRowInColumn(ws, DEST_COL_A)
    lastRowB = GetLastNonEmptyRowInColumn(ws, DEST_COL_B)

    If lastRowA > lastRow Then
        lastRow = lastRowA
    End If
    If lastRowB > lastRow Then
        lastRow = lastRowB
    End If

    GetLastConfirmedDestRow = lastRow
End Function

' 目的: 指定列で最後に値または数式がある行を検索する。
Private Function GetLastNonEmptyRowInColumn( _
    ByVal ws As Worksheet, _
    ByVal targetCol As Long) As Long

    Dim foundCell As Range

    If ws Is Nothing Then
        GetLastNonEmptyRowInColumn = FIRST_DEST_ROW
        Exit Function
    End If

    On Error Resume Next
    Set foundCell = ws.Columns(targetCol).Find( _
        What:="*", _
        After:=ws.Cells(1, targetCol), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)
    On Error GoTo 0

    If foundCell Is Nothing Then
        GetLastNonEmptyRowInColumn = FIRST_DEST_ROW
    ElseIf foundCell.Row < FIRST_DEST_ROW Then
        GetLastNonEmptyRowInColumn = FIRST_DEST_ROW
    Else
        GetLastNonEmptyRowInColumn = foundCell.Row
    End If
End Function

' 目的: 縦罫線の終端行に比較領域全体の下罫線を引き、表の下端を閉じる。
Private Sub ApplyBottomBorderClosure( _
    ByVal destWs As Worksheet, _
    ByVal lastWrittenRow As Long, _
    ByVal endRow As Long)

    Dim clearRow As Long
    Dim endCol As Long
    Dim borderRange As Range

    If destWs Is Nothing Then Exit Sub
    If Not IsTopBorderEnabled() Then Exit Sub

    clearRow = lastWrittenRow - 1
    endCol = GetTopBorderEndCol()

    If clearRow >= FIRST_DEST_ROW Then
        Set borderRange = destWs.Range( _
            destWs.Cells(clearRow, DEST_COL_A), _
            destWs.Cells(clearRow, endCol))

        With borderRange.Borders(xlEdgeBottom)
            .LineStyle = xlNone
        End With
    End If


    Set borderRange = destWs.Range( _
        destWs.Cells(endRow, DEST_COL_A), _
        destWs.Cells(endRow, endCol))

    With borderRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

' 目的: 先頭行を除く確定スロット行へ、変更後比較領域全体の上罫線を引く。
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

    aValue = destWs.Cells(targetRow, DEST_COL_A).value
    bValue = destWs.Cells(targetRow, DEST_COL_B).value

    If IsError(aValue) Then Exit Sub
    If IsError(bValue) Then Exit Sub

    hasAValue = (Len(Trim$(CStr(aValue))) > 0)
    hasBValue = (Len(Trim$(CStr(bValue))) > 0)

    If Not hasAValue And Not hasBValue Then Exit Sub

    Set borderRange = destWs.Range( _
        destWs.Cells(targetRow, DEST_COL_A), _
        destWs.Cells(targetRow, GetTopBorderEndCol()))

    With borderRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

' 目的: スロット番号と行オフセットから、A列・B列へ書き込む行番号を算出する。
Private Function GetDestRowForSlot(ByVal slotIndex As Long) As Long
    If slotIndex < 0 Then
        Err.Raise vbObjectError + 2401, "GetDestRowForSlot", "slotIndex が負数です。"
    End If

    GetDestRowForSlot = FIRST_DEST_ROW + (slotIndex * GetSlotHeightForWrite())
End Function

' ============================================================
' 共通ユーティリティ
' ============================================================

' 目的: 必須ワークシートを完全一致で取得し、存在しなければ用途を示して中断する。
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

' 目的: ファイル名の最後の拡張子だけを除き、REFER照合や置換に使う基底名を返す。
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






