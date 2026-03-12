Attribute VB_Name = "BetaTestCaseGenerator"
Option Explicit

' ============================================================
' Beta テストケースブック生成マクロ
' ------------------------------------------------------------
' 概要:
' 1) ユーザーが機能連番を入力
' 2) ThisWorkbook の REFER シートを走査して α/β/γ を収集
' 3) テンプレシートを複製して新規 .xlsx を作成
' 4) ルールに従ってシート名/セル値を設定して保存
' ============================================================

' ===== REFER 定義 =====
Private Const REFER_SHEET_NAME As String = "REFER"
Private Const REFER_ALPHA_COL_LETTER As String = "J" ' α: J列
Private Const REFER_BETA_COL_LETTER As String = "F"  ' β: F列
Private Const REFER_GAMMA_COL_LETTER As String = "E" ' γ: E列

' ===== テンプレシート名（ThisWorkbook 側） =====
Private Const TEMPLATE_COMMON_SHEET_NAME As String = "【共通】機能名"
Private Const TEMPLATE_INDIVIDUAL_SHEET_NAME As String = "【個別】機能名"
Private Const TEMPLATE_REFERENCE_SHEET_NAME As String = "⇒参考"
Private Const TEMPLATE_SOURCE_SHEET_NAME As String = "現行ソース（PHP）"
Private Const TEMPLATE_SCREEN_SHEET_NAME As String = "現行画面"

' ===== 出力シート名プレフィックス =====
Private Const OUTPUT_COMMON_PREFIX As String = "【共通】"
Private Const OUTPUT_INDIVIDUAL_PREFIX As String = "【個別】"
Private Const OUTPUT_SOURCE_PREFIX As String = "現行ソース（PHP）"

' ===== セル書き込み先 =====
Private Const TARGET_ALPHA_CELL As String = "BD1"
Private Const TARGET_FEATURE_ID_CELL As String = "BD3"
Private Const TARGET_GAMMA_CELL As String = "C4"

' ===== 出力ファイル名 =====
Private Const OUTPUT_FILE_SUFFIX As String = "_単体テストケース_初期開発"
Private Const OUTPUT_FILE_EXT As String = ".xlsx"

' ===== 参照レコード配列のインデックス =====
Private Const MATCH_IDX_ALPHA As Long = 1
Private Const MATCH_IDX_BETA As Long = 2
Private Const MATCH_IDX_GAMMA As Long = 3
Private Const MATCH_IDX_ROW As Long = 4

' ===== SaveAs ダイアログ =====
Private Const SAVE_AS_FILTER As String = "Excel ブック (*.xlsx),*.xlsx"
Public Type BetaTestCaseUiOptions
    Enabled As Boolean
    featureId As String
    useOutputPath As Boolean
    outputPath As String
End Type

Private mUiOptions As BetaTestCaseUiOptions

Public Sub RunMainWithUiOptions(ByRef options As BetaTestCaseUiOptions)
    ClearUiOptions
    mUiOptions = options
    mUiOptions.Enabled = True

    RunMain

    ClearUiOptions
End Sub

Public Function CreateBetaTestCaseUiOptionsForForm() As BetaTestCaseUiOptions
    Dim defaults As BetaTestCaseUiOptions

    defaults.Enabled = True
    defaults.featureId = vbNullString
    defaults.useOutputPath = False
    defaults.outputPath = vbNullString

    CreateBetaTestCaseUiOptionsForForm = defaults
End Function

Private Sub ClearUiOptions()
    mUiOptions.Enabled = False
    mUiOptions.featureId = vbNullString
    mUiOptions.useOutputPath = False
    mUiOptions.outputPath = vbNullString
End Sub

' ============================================================
' 実行入口
' ============================================================

Public Sub RunMain()
    On Error GoTo ErrorHandler

    Dim macroWb As Workbook
    Dim referWs As Worksheet
    Dim outputWb As Workbook

    Dim featureId As String
    Dim alpha As String
    Dim alphaWarning As String
    Dim outputPath As String
    Dim finalMessage As String

    Dim matches As Collection
    Dim createdSheetCount As Long
    Dim completedSuccessfully As Boolean

    ' 例外時に必ず戻すため、アプリ状態を退避してから変更する
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim appStateCaptured As Boolean

    Set macroWb = ThisWorkbook
    Set referWs = GetWorksheetOrRaise(macroWb, REFER_SHEET_NAME, "REFER シート")

    featureId = PromptFeatureId()
    If Len(featureId) = 0 Then Exit Sub

    Set matches = FindReferMatches(referWs, featureId)
    If matches.Count = 0 Then
        MsgBox "REFER シートの " & REFER_ALPHA_COL_LETTER & " 列に、" & _
               "機能連番を含む行が見つかりませんでした。" & vbCrLf & _
               "入力値: " & featureId, vbInformation
        Exit Sub
    End If

    alpha = ResolvePrimaryAlpha(matches, alphaWarning)
    If Len(alphaWarning) > 0 Then
        MsgBox alphaWarning, vbExclamation
    End If

    outputPath = DecideOutputPath(macroWb, alpha)
    If Len(outputPath) = 0 Then
        MsgBox "保存先が未選択のため処理を中断しました。", vbInformation
        Exit Sub
    End If

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    appStateCaptured = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set outputWb = BuildOutputWorkbook( _
        macroWb:=macroWb, _
        matches:=matches, _
        alpha:=alpha, _
        featureId:=featureId, _
        createdSheetCount:=createdSheetCount)

    ActivateFirstWorksheetForOpenState outputWb
    outputWb.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    outputWb.Close SaveChanges:=False
    Set outputWb = Nothing

    completedSuccessfully = True

    finalMessage = "テストケースブックを作成しました。" & vbCrLf & _
                   "保存先: " & outputPath & vbCrLf & _
                   "機能連番: " & featureId & vbCrLf & _
                   "ヒット件数: " & CStr(matches.Count) & vbCrLf & _
                   "作成シート数: " & CStr(createdSheetCount)

    GoTo SafeExit

ErrorHandler:
    finalMessage = "エラーが発生しました。" & vbCrLf & _
                   Err.Number & " : " & Err.Description

SafeExit:
    On Error Resume Next

    ' 作成途中で失敗した場合は、作りかけブック/ファイルを片付ける
    If Not completedSuccessfully Then
        If Not outputWb Is Nothing Then
            outputWb.Close SaveChanges:=False
        End If
        If Len(outputPath) > 0 Then
            If Len(Dir$(outputPath)) > 0 Then
                Kill outputPath
            End If
        End If
    End If

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
' 入力・出力パス決定
' ============================================================

Private Function PromptFeatureId() As String
    ' 機能連番を InputBox で入力させる。
    ' キャンセルまたは空文字は空で返し、呼び出し元で中断判断する。
    Dim s As String

    If mUiOptions.Enabled Then
        PromptFeatureId = Trim$(mUiOptions.featureId)
        Exit Function
    End If

    s = InputBox("機能連番を入力してください（例: S99-999-99）", "機能連番入力")
    PromptFeatureId = Trim$(s)
End Function

Private Function DecideOutputPath(ByVal macroWb As Workbook, ByVal alpha As String) As String
    ' 1) ThisWorkbook が保存済みなら同フォルダに出力
    ' 2) 未保存なら SaveAs ダイアログで保存先を選ばせる
    ' 3) 同名ファイルがある場合は _001, _002 ... を付与する
    Dim defaultFileName As String
    Dim desiredPath As String

    defaultFileName = alpha & OUTPUT_FILE_SUFFIX & OUTPUT_FILE_EXT

    If mUiOptions.Enabled Then
        If mUiOptions.useOutputPath Then
            DecideOutputPath = BuildUniquePath(mUiOptions.outputPath)
            Exit Function
        End If

        If Len(Trim$(macroWb.Path)) > 0 Then
            desiredPath = macroWb.Path & "\" & defaultFileName
        Else
            DecideOutputPath = vbNullString
            Exit Function
        End If

        DecideOutputPath = BuildUniquePath(desiredPath)
        Exit Function
    End If

    If Len(Trim$(macroWb.Path)) > 0 Then
        desiredPath = macroWb.Path & "\" & defaultFileName
    Else
        desiredPath = PromptOutputPathByDialog(defaultFileName)
        If Len(desiredPath) = 0 Then Exit Function
    End If

    DecideOutputPath = BuildUniquePath(desiredPath)
End Function

Private Function PromptOutputPathByDialog(ByVal defaultFileName As String) As String
    ' ThisWorkbook 未保存時の保存先選択ダイアログ。
    Dim pickedPath As Variant

    pickedPath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultFileName, _
        FileFilter:=SAVE_AS_FILTER, _
        Title:="出力ファイルの保存先を選択してください")

    If VarType(pickedPath) = vbBoolean Then
        PromptOutputPathByDialog = vbNullString
        Exit Function
    End If

    PromptOutputPathByDialog = EnsureXlsxExtension(CStr(pickedPath))
End Function

Private Function BuildUniquePath(ByVal desiredPath As String) As String
    ' 同名ファイルがある場合は _001 形式で連番を付与する。
    ' 既に開いているブック名とも衝突しないようにする。
    Dim basePart As String
    Dim extPart As String
    Dim candidate As String
    Dim seqNo As Long
    Dim lastSepPos As Long
    Dim lastDotPos As Long

    desiredPath = EnsureXlsxExtension(Trim$(desiredPath))
    If Len(desiredPath) = 0 Then Exit Function

    If Len(Dir$(desiredPath)) = 0 And Not IsWorkbookAlreadyOpen(desiredPath) Then
        BuildUniquePath = desiredPath
        Exit Function
    End If

    lastSepPos = InStrRev(desiredPath, "\")
    lastDotPos = InStrRev(desiredPath, ".")

    If lastDotPos > (lastSepPos + 1) Then
        basePart = Left$(desiredPath, lastDotPos - 1)
        extPart = Mid$(desiredPath, lastDotPos)
    Else
        basePart = desiredPath
        extPart = OUTPUT_FILE_EXT
    End If

    For seqNo = 1 To 9999
        candidate = basePart & "_" & Format$(seqNo, "000") & extPart
        If Len(Dir$(candidate)) = 0 And Not IsWorkbookAlreadyOpen(candidate) Then
            BuildUniquePath = candidate
            Exit Function
        End If
    Next seqNo

    Err.Raise vbObjectError + 6101, "BuildUniquePath", _
              "連番付きファイル名を決定できませんでした（上限: 9999）。"
End Function

Private Function EnsureXlsxExtension(ByVal filePath As String) As String
    ' 指定パスの拡張子を .xlsx に正規化する。
    Dim lastSepPos As Long
    Dim lastDotPos As Long
    Dim extPart As String

    filePath = Trim$(filePath)
    If Len(filePath) = 0 Then Exit Function

    lastSepPos = InStrRev(filePath, "\")
    lastDotPos = InStrRev(filePath, ".")

    If lastDotPos > (lastSepPos + 1) Then
        extPart = Mid$(filePath, lastDotPos)
        If StrComp(extPart, OUTPUT_FILE_EXT, vbTextCompare) = 0 Then
            EnsureXlsxExtension = filePath
        Else
            EnsureXlsxExtension = Left$(filePath, lastDotPos - 1) & OUTPUT_FILE_EXT
        End If
    Else
        EnsureXlsxExtension = filePath & OUTPUT_FILE_EXT
    End If
End Function

Private Function IsWorkbookAlreadyOpen(ByVal workbookPath As String) As Boolean
    ' フルパス一致で、既に開いているブックがあるかを判定。
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, workbookPath, vbTextCompare) = 0 Then
            IsWorkbookAlreadyOpen = True
            Exit Function
        End If
    Next wb
End Function

' ============================================================
' REFER 探索（操作X）
' ============================================================

Private Function FindReferMatches( _
    ByVal referWs As Worksheet, _
    ByVal featureId As String) As Collection

    ' REFER の J 列を上から走査し、機能連番を含む行を収集する。
    ' 戻り値の Collection は 1要素=Variant配列(1..4):
    '   1=alpha, 2=beta, 3=gamma, 4=rowNumber
    Dim results As Collection
    Dim alphaColIndex As Long
    Dim betaColIndex As Long
    Dim gammaColIndex As Long

    Dim lastRow As Long
    Dim r As Long

    Dim alphaRaw As Variant
    Dim betaRaw As Variant
    Dim gammaRaw As Variant

    Dim alphaText As String
    Dim betaText As String
    Dim gammaText As String

    Dim alphaValues As Variant
    Dim betaValues As Variant
    Dim gammaValues As Variant

    Dim record(1 To 4) As Variant

    Set results = New Collection

    alphaColIndex = ColumnLetterToIndex(REFER_ALPHA_COL_LETTER, "α列")
    betaColIndex = ColumnLetterToIndex(REFER_BETA_COL_LETTER, "β列")
    gammaColIndex = ColumnLetterToIndex(REFER_GAMMA_COL_LETTER, "γ列")

    lastRow = referWs.Cells(referWs.Rows.Count, alphaColIndex).End(xlUp).Row
    If lastRow < 1 Then
        Set FindReferMatches = results
        Exit Function
    End If

    ' COMアクセス回数を減らすため、必要列を配列へ読み込む
    alphaValues = ReadColumnValues(referWs, alphaColIndex, 1, lastRow)
    betaValues = ReadColumnValues(referWs, betaColIndex, 1, lastRow)
    gammaValues = ReadColumnValues(referWs, gammaColIndex, 1, lastRow)

    For r = 1 To lastRow
        alphaRaw = alphaValues(r, 1)
        If IsError(alphaRaw) Then
            Err.Raise vbObjectError + 6201, "FindReferMatches", _
                      "REFER の " & REFER_ALPHA_COL_LETTER & " 列にエラー値があります（行: " & CStr(r) & "）。"
        End If

        alphaText = Trim$(CStr(alphaRaw))
        If Len(alphaText) = 0 Then GoTo ContinueRow

        ' 仕様: 機能連番を含む行をヒット（部分一致）
        If InStr(1, alphaText, featureId, vbTextCompare) > 0 Then
            betaRaw = betaValues(r, 1)
            gammaRaw = gammaValues(r, 1)

            If IsError(betaRaw) Then
                Err.Raise vbObjectError + 6202, "FindReferMatches", _
                          "REFER の " & REFER_BETA_COL_LETTER & " 列にエラー値があります（行: " & CStr(r) & "）。"
            End If
            If IsError(gammaRaw) Then
                Err.Raise vbObjectError + 6203, "FindReferMatches", _
                          "REFER の " & REFER_GAMMA_COL_LETTER & " 列にエラー値があります（行: " & CStr(r) & "）。"
            End If

            alphaText = RemoveExtension(alphaText)
            betaText = Trim$(CStr(betaRaw))
            gammaText = Trim$(CStr(gammaRaw))

            If Len(alphaText) = 0 Then
                Err.Raise vbObjectError + 6204, "FindReferMatches", _
                          "α（J列）から拡張子除去後の文字列が空です（行: " & CStr(r) & "）。"
            End If
            If Len(betaText) = 0 Then
                Err.Raise vbObjectError + 6205, "FindReferMatches", _
                          "β（F列）が空です（行: " & CStr(r) & "）。"
            End If
            If Len(gammaText) = 0 Then
                Err.Raise vbObjectError + 6206, "FindReferMatches", _
                          "γ（E列）が空です（行: " & CStr(r) & "）。"
            End If

            record(MATCH_IDX_ALPHA) = alphaText
            record(MATCH_IDX_BETA) = betaText
            record(MATCH_IDX_GAMMA) = gammaText
            record(MATCH_IDX_ROW) = r
            results.Add record
        End If

ContinueRow:
    Next r

    Set FindReferMatches = results
End Function

Private Function ReadColumnValues( _
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
        ReadColumnValues = singleCell
    Else
        ReadColumnValues = rawValues
    End If
End Function
Private Function ResolvePrimaryAlpha( _
    ByVal matches As Collection, _
    ByRef warningMessage As String) As String

    ' ヒット複数時、α差異があれば「最初のαを採用」して警告を返す。
    Dim i As Long
    Dim firstAlpha As String
    Dim currentAlpha As String
    Dim mismatchCount As Long
    Dim details As String
    Dim record As Variant

    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    record = matches(1)
    firstAlpha = CStr(record(MATCH_IDX_ALPHA))

    For i = 2 To matches.Count
        record = matches(i)
        currentAlpha = CStr(record(MATCH_IDX_ALPHA))

        If StrComp(firstAlpha, currentAlpha, vbBinaryCompare) <> 0 Then
            mismatchCount = mismatchCount + 1
            If Len(details) < 300 Then
                details = details & "行" & CStr(record(MATCH_IDX_ROW)) & ": " & currentAlpha & vbCrLf
            End If
        End If
    Next i

    If mismatchCount > 0 Then
        warningMessage = "REFER のヒット行で α が一致しません。最初の α を採用して処理を継続します。" & vbCrLf & _
                         "採用 α: " & firstAlpha & vbCrLf & _
                         "差異件数: " & CStr(mismatchCount) & vbCrLf & details
    End If

    ResolvePrimaryAlpha = firstAlpha
End Function

' ============================================================
' 出力ブック構築
' ============================================================

Private Function BuildOutputWorkbook( _
    ByVal macroWb As Workbook, _
    ByVal matches As Collection, _
    ByVal alpha As String, _
    ByVal featureId As String, _
    ByRef createdSheetCount As Long) As Workbook

    ' 新規ブックを作り、テンプレシートをルール順に複製して値を書き込む。
    Dim outputWb As Workbook
    Dim seedSheetNames As Collection

    Dim templateCommonWs As Worksheet
    Dim templateIndividualWs As Worksheet
    Dim templateReferenceWs As Worksheet
    Dim templateSourceWs As Worksheet
    Dim templateScreenWs As Worksheet

    Dim i As Long
    Dim record As Variant
    Dim betaText As String
    Dim gammaText As String

    Dim ws As Worksheet

    Set outputWb = Application.Workbooks.Add(xlWBATWorksheet)
    Set seedSheetNames = CaptureInitialSheetNames(outputWb)

    Set templateCommonWs = GetWorksheetOrRaise(macroWb, TEMPLATE_COMMON_SHEET_NAME, "共通テンプレ")
    Set templateIndividualWs = GetWorksheetOrRaise(macroWb, TEMPLATE_INDIVIDUAL_SHEET_NAME, "個別テンプレ")
    Set templateReferenceWs = GetWorksheetOrRaise(macroWb, TEMPLATE_REFERENCE_SHEET_NAME, "参考テンプレ")
    Set templateSourceWs = GetWorksheetOrRaise(macroWb, TEMPLATE_SOURCE_SHEET_NAME, "現行ソーステンプレ")
    Set templateScreenWs = GetWorksheetOrRaise(macroWb, TEMPLATE_SCREEN_SHEET_NAME, "現行画面テンプレ")

    ' 1) βごとに 【共通】/【個別】 を作成
    For i = 1 To matches.Count
        record = matches(i)
        betaText = CStr(record(MATCH_IDX_BETA))

        Set ws = CopyTemplateSheet(templateCommonWs, outputWb, OUTPUT_COMMON_PREFIX & betaText)
        FillCaseSheet ws, betaText, featureId
        createdSheetCount = createdSheetCount + 1

        Set ws = CopyTemplateSheet(templateIndividualWs, outputWb, OUTPUT_INDIVIDUAL_PREFIX & betaText)
        FillCaseSheet ws, betaText, featureId
        createdSheetCount = createdSheetCount + 1
    Next i

    ' 2) ⇒参考 は1枚だけ
    Set ws = CopyTemplateSheet(templateReferenceWs, outputWb, TEMPLATE_REFERENCE_SHEET_NAME)
    createdSheetCount = createdSheetCount + 1

    ' 3) βごとに 現行ソース（PHP） を作成し C4 に γ を設定
    For i = 1 To matches.Count
        record = matches(i)
        betaText = CStr(record(MATCH_IDX_BETA))
        gammaText = CStr(record(MATCH_IDX_GAMMA))

        Set ws = CopyTemplateSheet(templateSourceWs, outputWb, OUTPUT_SOURCE_PREFIX & betaText)
        FillSourceSheet ws, gammaText
        createdSheetCount = createdSheetCount + 1
    Next i

    ' 4) 現行画面 は1枚だけ
    Set ws = CopyTemplateSheet(templateScreenWs, outputWb, TEMPLATE_SCREEN_SHEET_NAME)
    createdSheetCount = createdSheetCount + 1

    RemoveSeedSheetsIfNeeded outputWb, seedSheetNames

    Set BuildOutputWorkbook = outputWb
End Function

Private Sub FillCaseSheet( _
    ByVal targetWs As Worksheet, _
    ByVal beta As String, _
    ByVal featureId As String)

    ' 【共通】/【個別】シートの共通セル埋め。
    ' 仕様: BD1 は β、BD3 は入力した機能連番。
    targetWs.Range(TARGET_ALPHA_CELL).value = beta
    targetWs.Range(TARGET_FEATURE_ID_CELL).value = featureId
End Sub

Private Sub FillSourceSheet( _
    ByVal sourceWs As Worksheet, _
    ByVal gamma As String)

    ' 現行ソース（PHP）シートのセル埋め。
    sourceWs.Range(TARGET_GAMMA_CELL).value = gamma
End Sub

Private Function CopyTemplateSheet( _
    ByVal templateWs As Worksheet, _
    ByVal targetWb As Workbook, _
    ByVal desiredName As String) As Worksheet

    ' テンプレートをターゲットブックへコピーし、安全なシート名へ変更する。
    ' 先にユニーク名を決めることで、「⇒参考」「現行画面」に不要な連番が付くのを防ぐ。
    Dim safeName As String

    safeName = MakeUniqueSheetName(targetWb, desiredName)

    templateWs.Copy After:=targetWb.Worksheets(targetWb.Worksheets.Count)
    Set CopyTemplateSheet = targetWb.Worksheets(targetWb.Worksheets.Count)

    If StrComp(CopyTemplateSheet.Name, safeName, vbBinaryCompare) = 0 Then
        Exit Function
    End If

    On Error GoTo RenameError
    CopyTemplateSheet.Name = safeName
    On Error GoTo 0

    Exit Function

RenameError:
    Err.Raise vbObjectError + 6301, "CopyTemplateSheet", _
              "シート名を設定できませんでした。" & vbCrLf & _
              "指定名: " & desiredName & vbCrLf & _
              "安全化後: " & safeName
End Function

Private Function CaptureInitialSheetNames(ByVal wb As Workbook) As Collection
    ' 新規ブック作成直後に存在したシート名を保持する。
    ' Excel設定で初期シートが複数でも、最後に確実に除去できるようにする。
    Dim names As Collection
    Dim ws As Worksheet

    Set names = New Collection
    For Each ws In wb.Worksheets
        names.Add ws.Name
    Next ws

    Set CaptureInitialSheetNames = names
End Function

Private Sub RemoveSeedSheetsIfNeeded(ByVal wb As Workbook, ByVal seedSheetNames As Collection)
    ' 新規ブック作成時の初期シート（Sheet1 など）を削除する。
    ' 初期シートが複数ある設定でも全て除去する。
    Dim i As Long
    Dim seedName As String
    Dim seedWs As Worksheet

    If wb Is Nothing Then Exit Sub
    If seedSheetNames Is Nothing Then Exit Sub

    For i = seedSheetNames.Count To 1 Step -1
        If wb.Worksheets.Count <= 1 Then Exit For

        seedName = CStr(seedSheetNames(i))
        Set seedWs = FindWorksheetExact(wb, seedName)
        If Not seedWs Is Nothing Then
            seedWs.Delete
        End If
    Next i
End Sub

Private Sub ActivateFirstWorksheetForOpenState(ByVal wb As Workbook)
    Dim firstWs As Worksheet

    If wb Is Nothing Then Exit Sub
    If wb.Worksheets.Count = 0 Then Exit Sub

    Set firstWs = wb.Worksheets(1)
    firstWs.Activate
End Sub


' ============================================================
' シート名安全化
' ============================================================

Private Function MakeUniqueSheetName( _
    ByVal wb As Workbook, _
    ByVal desiredName As String) As String

    ' Excel シート名制約(31文字/禁止文字)を満たしつつ、
    ' 同一ブック内で重複しない名前を返す。
    Dim baseName As String
    Dim candidate As String
    Dim suffix As String
    Dim seqNo As Long

    baseName = NormalizeSheetName(desiredName)
    candidate = baseName

    If FindWorksheetExact(wb, candidate) Is Nothing Then
        MakeUniqueSheetName = candidate
        Exit Function
    End If

    For seqNo = 1 To 9999
        suffix = "_" & Format$(seqNo, "000")
        candidate = Left$(baseName, 31 - Len(suffix)) & suffix

        If FindWorksheetExact(wb, candidate) Is Nothing Then
            MakeUniqueSheetName = candidate
            Exit Function
        End If
    Next seqNo

    Err.Raise vbObjectError + 6401, "MakeUniqueSheetName", _
              "一意なシート名を決定できませんでした: " & desiredName
End Function

Private Function NormalizeSheetName(ByVal rawName As String) As String
    ' 禁止文字置換 + 長さ制限 + 空文字回避
    Dim s As String

    s = Trim$(rawName)
    If Len(s) = 0 Then s = "Sheet"

    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, "[", "_")
    s = Replace(s, "]", "_")

    If Left$(s, 1) = "'" Then s = Mid$(s, 2)
    If Right$(s, 1) = "'" Then s = Left$(s, Len(s) - 1)

    If Len(s) = 0 Then s = "Sheet"
    If Len(s) > 31 Then s = Left$(s, 31)

    NormalizeSheetName = s
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
        Err.Raise vbObjectError + 6501, "GetWorksheetOrRaise", _
                  labelForMessage & " が見つかりません: " & sheetName
    End If
End Function

Private Function FindWorksheetExact(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    ' シート名完全一致。見つからない場合は Nothing。
    On Error Resume Next
    Set FindWorksheetExact = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function ColumnLetterToIndex( _
    ByVal columnLetter As String, _
    Optional ByVal labelForError As String = vbNullString) As Long

    ' "A" -> 1, "F" -> 6, "AA" -> 27
    Dim normalized As String
    Dim i As Long
    Dim ch As String
    Dim chCode As Long
    Dim prefix As String

    normalized = UCase$(Trim$(columnLetter))
    If Len(normalized) = 0 Then
        prefix = BuildErrorLabelPrefix(labelForError)
        Err.Raise vbObjectError + 6502, "ColumnLetterToIndex", prefix & "列指定が空です。"
    End If

    For i = 1 To Len(normalized)
        ch = Mid$(normalized, i, 1)
        chCode = Asc(ch)

        If chCode < 65 Or chCode > 90 Then
            prefix = BuildErrorLabelPrefix(labelForError)
            Err.Raise vbObjectError + 6503, "ColumnLetterToIndex", _
                      prefix & "列指定が不正です: " & columnLetter
        End If

        ColumnLetterToIndex = (ColumnLetterToIndex * 26) + (chCode - 64)
    Next i

    If ColumnLetterToIndex < 1 Or ColumnLetterToIndex > 16384 Then
        prefix = BuildErrorLabelPrefix(labelForError)
        Err.Raise vbObjectError + 6504, "ColumnLetterToIndex", _
                  prefix & "列番号がExcelの範囲外です: " & columnLetter
    End If
End Function

Private Function BuildErrorLabelPrefix(ByVal labelText As String) As String
    If Len(Trim$(labelText)) = 0 Then
        BuildErrorLabelPrefix = vbNullString
    Else
        BuildErrorLabelPrefix = Trim$(labelText) & " "
    End If
End Function

Private Function RemoveExtension(ByVal fileNameText As String) As String
    ' "foo.php" -> "foo"
    ' "foo.bar.php" -> "foo.bar"
    ' "foo" -> "foo"
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



