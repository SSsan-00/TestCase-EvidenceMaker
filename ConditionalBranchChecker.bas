Attribute VB_Name = "ConditionalBranchChecker"
Option Explicit

' ==========================================
' Excel用: 現行ソースシートのマーキング + 個別シート出力
' 文字列ベースで構文を判定する
' ==========================================

Private Const SOURCE_TEXT_COL As Long = 3   ' C列
Private Const MARK_COL As Long = 2          ' B列
Private Const SECTION_HEADER_START_ROW As Long = 9
Private Const BLOCK_STEP_NORMAL As Long = 5
Private Const INDIVIDUAL_TEMPLATE_INSERT_COUNT As Long = 50
Private Const INDIVIDUAL_TEMPLATE_PREINSERT_MARGIN As Long = 50
Private Const INDIVIDUAL_TEMPLATE_SOURCE_START_ROW As Long = 15
Private Const INDIVIDUAL_TEMPLATE_SOURCE_ROW_COUNT As Long = 10
Private Const SYMBOL_FILLED As String = "■"
Private Const SYMBOL_EMPTY As String = "□"
Private Const SHEET_KEY_CURRENT_SOURCE As String = "現行ソース"
Private Const SHEET_KEY_INDIVIDUAL_PREFIX As String = "【個別】"
Private Const TEMPLATE_SNAPSHOT_SHEET_PREFIX As String = "__TMPROW15_"
Private Const LEADING_FUNCTION_STARTS_FROM_B1 As Boolean = True  ' True: 最初の判定対象がfunctionならB1開始にする

Private mTemplateSnapshotSheet As Worksheet
Private mPreAllocatedWritableLastRow As Long

Public Sub RunMain()
    On Error GoTo ErrorHandler

    Dim featureName As String
    Dim workbookPath As String
    Dim targetWorkbook As Workbook
    Dim currentSourceSheet As Worksheet
    Dim individualSheet As Worksheet
    Dim syntaxEvents As Collection
    Dim markedCount As Long
    Dim resultMessage As String
    Dim sourceTextValues As Variant
    Dim sourceLastRow As Long
    Dim useLeadingFunctionB1 As Boolean

    ' 1) 機能名を入力
    featureName = PromptFeatureName()
    If Len(featureName) = 0 Then
        MsgBox "処理をキャンセルしました（機能名が未入力です）。", vbInformation
        Exit Sub
    End If

    ' 2) 対象ファイルを選択
    workbookPath = SelectTargetWorkbookPath()
    If Len(workbookPath) = 0 Then
        MsgBox "処理をキャンセルしました（対象ファイルが未選択です）。", vbInformation
        Exit Sub
    End If

    ' 3) 対象ブックを開く
    Set targetWorkbook = OpenTargetWorkbook(workbookPath)
    If targetWorkbook Is Nothing Then
        MsgBox "対象ブックを開けませんでした。", vbExclamation
        Exit Sub
    End If

    ' 4) 対象シートを判定
    Set currentSourceSheet = FindCurrentSourceSheet(targetWorkbook, featureName)
    If currentSourceSheet Is Nothing Then
        MsgBox "現行ソースシートが見つかりませんでした。" & vbCrLf & _
               "条件: シート名に「" & SHEET_KEY_CURRENT_SOURCE & "」を含む", vbExclamation
        Exit Sub
    End If

    Set individualSheet = FindIndividualSheet(targetWorkbook, featureName)

    sourceLastRow = GetLastRow(currentSourceSheet, SOURCE_TEXT_COL)
    sourceTextValues = ReadColumnValues(currentSourceSheet, SOURCE_TEXT_COL, 1, sourceLastRow)
    useLeadingFunctionB1 = ShouldStartFunctionSectionFromB1(sourceTextValues, sourceLastRow)

    ' 5) 現行ソースシートに対してマーキング
    MarkCurrentSourceSheet currentSourceSheet, sourceTextValues, sourceLastRow, markedCount, useLeadingFunctionB1

    ' 6) 個別シートがある場合は解析結果を書き込む
    If Not individualSheet Is Nothing Then
        Set syntaxEvents = CollectSyntaxEvents(sourceTextValues, sourceLastRow)
        WriteIndividualSheet individualSheet, syntaxEvents, useLeadingFunctionB1
        resultMessage = "個別シート出力: 実施（" & individualSheet.Name & "）"
    Else
        resultMessage = "個別シート出力: スキップ（対象シートなし）"
    End If

    ' 対象ブックへ書き込みした内容を保存
    ActivateSheetForNextOpen targetWorkbook, currentSourceSheet
    targetWorkbook.Save

    ' 7) 完了メッセージ
    MsgBox "処理が完了しました。" & vbCrLf & _
           "対象ブック: " & targetWorkbook.Name & vbCrLf & _
           "現行ソース: " & currentSourceSheet.Name & vbCrLf & _
           "マーキング件数: " & CStr(markedCount) & vbCrLf & _
           resultMessage, vbInformation

    Exit Sub

ErrorHandler:
    ' 8) エラー時はMsgBox表示のみ（ログ不要）
    MsgBox "エラーが発生しました。" & vbCrLf & _
           Err.Number & " : " & Err.Description, vbExclamation
End Sub

Private Function PromptFeatureName() As String
    ' 機能名の入力を受け取り、前後空白を除去して返す
    Dim inputValue As String
    inputValue = InputBox("機能名を入力してください。", "機能名入力")
    PromptFeatureName = Trim$(inputValue)
End Function

Private Function SelectTargetWorkbookPath() As String
    ' .xlsx を選択させる簡易ダイアログ
    Dim selectedPath As Variant

    selectedPath = Application.GetOpenFilename( _
        FileFilter:="Excel ブック (*.xlsx),*.xlsx", _
        Title:="対象のExcelファイル（.xlsx）を選択してください")

    If VarType(selectedPath) = vbBoolean Then
        SelectTargetWorkbookPath = vbNullString
    Else
        SelectTargetWorkbookPath = CStr(selectedPath)
    End If
End Function

Private Function OpenTargetWorkbook(ByVal workbookPath As String) As Workbook
    ' 既に開いている場合は既存のWorkbookを返し、未オープンなら開く
    Dim wb As Workbook

    If Len(Trim$(workbookPath)) = 0 Then
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, workbookPath, vbTextCompare) = 0 Then
            Set OpenTargetWorkbook = wb
            Exit Function
        End If
    Next wb

    Set OpenTargetWorkbook = Application.Workbooks.Open( _
        Filename:=workbookPath, _
        UpdateLinks:=0, _
        ReadOnly:=False)
End Function

Private Function FindCurrentSourceSheet(ByVal targetWorkbook As Workbook, ByVal featureName As String) As Worksheet
    ' 現行ソースシートの判定ルール:
    ' - 名前に「現行ソース」を含むシートを候補
    ' - 候補が1枚ならそれを採用
    ' - 候補が複数なら、機能名が完全一致するものを採用
    Dim ws As Worksheet
    Dim candidates As Collection
    Dim item As Variant

    Set candidates = New Collection

    For Each ws In targetWorkbook.Worksheets
        If ContainsText(ws.Name, SHEET_KEY_CURRENT_SOURCE) Then
            candidates.Add ws
        End If
    Next ws

    If candidates.Count = 0 Then
        Exit Function
    End If

    If candidates.Count = 1 Then
        Set FindCurrentSourceSheet = candidates.item(1)
        Exit Function
    End If

    If Len(featureName) > 0 Then
        For Each item In candidates
            Set ws = item
            If IsCurrentSourceFeatureExactMatch(ws.Name, featureName) Then
                Set FindCurrentSourceSheet = ws
                Exit Function
            End If
        Next item
    End If

    ' 複数候補があり、機能名一致がない場合は未確定としてNothingを返す
End Function

Private Function IsCurrentSourceFeatureExactMatch(ByVal sheetName As String, ByVal featureName As String) As Boolean
    ' 現行ソースシート名から機能名相当部分を取り出し、入力機能名と完全一致で比較する
    ' 例: 現行ソース（PHP）＜機能名＞
    Dim normalizedFeatureName As String
    Dim extractedFeatureName As String
    Dim normalizedSheetName As String

    normalizedFeatureName = Trim$(featureName)
    If Len(normalizedFeatureName) = 0 Then Exit Function

    extractedFeatureName = ExtractFeatureNameFromCurrentSourceSheetName(sheetName)
    If Len(extractedFeatureName) > 0 Then
        If StrComp(extractedFeatureName, normalizedFeatureName, vbTextCompare) = 0 Then
            IsCurrentSourceFeatureExactMatch = True
            Exit Function
        End If
    End If

    normalizedSheetName = Trim$(sheetName)

    ' 現行ソース（PHP）＜機能名＞ / 現行ソース(PHP)<機能名> の形式にも対応
    If StrComp(normalizedSheetName, "現行ソース（PHP）＜" & normalizedFeatureName & "＞", vbTextCompare) = 0 Or _
       StrComp(normalizedSheetName, "現行ソース(PHP)＜" & normalizedFeatureName & "＞", vbTextCompare) = 0 Or _
       StrComp(normalizedSheetName, "現行ソース（PHP）<" & normalizedFeatureName & ">", vbTextCompare) = 0 Or _
       StrComp(normalizedSheetName, "現行ソース(PHP)<" & normalizedFeatureName & ">", vbTextCompare) = 0 Then
        IsCurrentSourceFeatureExactMatch = True
        Exit Function
    End If

    ' 角括弧がない連結形式にも対応
    If StrComp(normalizedSheetName, "現行ソース（PHP）" & normalizedFeatureName, vbTextCompare) = 0 Or _
       StrComp(normalizedSheetName, "現行ソース(PHP)" & normalizedFeatureName, vbTextCompare) = 0 Then
        IsCurrentSourceFeatureExactMatch = True
    End If
End Function

Private Function ExtractFeatureNameFromCurrentSourceSheetName(ByVal sheetName As String) As String
    ' 例:
    ' - 現行ソース（PHP）＜機能A＞ -> 機能A
    ' - 現行ソース(PHP)<機能A>   -> 機能A
    ' - 現行ソース（PHP）機能A    -> 機能A
    Dim token As String

    token = Trim$(sheetName)

    ' 最優先: ＜...＞ / <...> で囲まれた部分を機能名として採用
    ExtractFeatureNameFromCurrentSourceSheetName = ExtractWrappedValue(token, "＜", "＞")
    If Len(ExtractFeatureNameFromCurrentSourceSheetName) = 0 Then
        ExtractFeatureNameFromCurrentSourceSheetName = ExtractWrappedValue(token, "<", ">")
    End If
    If Len(ExtractFeatureNameFromCurrentSourceSheetName) > 0 Then
        Exit Function
    End If

    token = Replace(token, SHEET_KEY_CURRENT_SOURCE, vbNullString, 1, -1, vbTextCompare)
    token = Replace(token, "（PHP）", vbNullString, 1, -1, vbTextCompare)
    token = Replace(token, "(PHP)", vbNullString, 1, -1, vbTextCompare)

    token = Trim$(token)
    token = TrimLeadingSeparators(token)
    token = TrimTrailingSeparators(token)
    token = UnwrapBracketPair(token)

    ExtractFeatureNameFromCurrentSourceSheetName = Trim$(token)
End Function

Private Function ExtractWrappedValue(ByVal sourceText As String, ByVal openToken As String, ByVal closeToken As String) As String
    Dim openPos As Long
    Dim closePos As Long

    If Len(openToken) = 0 Or Len(closeToken) = 0 Then Exit Function

    openPos = InStr(1, sourceText, openToken, vbTextCompare)
    If openPos = 0 Then Exit Function

    closePos = InStr(openPos + Len(openToken), sourceText, closeToken, vbTextCompare)
    If closePos = 0 Then Exit Function

    ExtractWrappedValue = Trim$(Mid$(sourceText, openPos + Len(openToken), closePos - openPos - Len(openToken)))
End Function

Private Function TrimLeadingSeparators(ByVal valueText As String) As String
    Dim textBuffer As String
    textBuffer = valueText

    Do While Len(textBuffer) > 0
        Select Case Left$(textBuffer, 1)
            Case " ", vbTab, "_", "-", ":", "：", "・", "/", "／", "|", "｜"
                textBuffer = Mid$(textBuffer, 2)
            Case Else
                Exit Do
        End Select
    Loop

    TrimLeadingSeparators = textBuffer
End Function

Private Function TrimTrailingSeparators(ByVal valueText As String) As String
    Dim textBuffer As String
    textBuffer = valueText

    Do While Len(textBuffer) > 0
        Select Case Right$(textBuffer, 1)
            Case " ", vbTab, "_", "-", ":", "：", "・", "/", "／", "|", "｜"
                textBuffer = Left$(textBuffer, Len(textBuffer) - 1)
            Case Else
                Exit Do
        End Select
    Loop

    TrimTrailingSeparators = textBuffer
End Function

Private Function UnwrapBracketPair(ByVal valueText As String) As String
    Dim textBuffer As String
    textBuffer = valueText

    If Len(textBuffer) < 2 Then
        UnwrapBracketPair = textBuffer
        Exit Function
    End If

    If (Left$(textBuffer, 1) = "（" And Right$(textBuffer, 1) = "）") Or _
       (Left$(textBuffer, 1) = "(" And Right$(textBuffer, 1) = ")") Or _
       (Left$(textBuffer, 1) = "【" And Right$(textBuffer, 1) = "】") Or _
       (Left$(textBuffer, 1) = "[" And Right$(textBuffer, 1) = "]") Then
        textBuffer = Mid$(textBuffer, 2, Len(textBuffer) - 2)
    End If

    UnwrapBracketPair = textBuffer
End Function

Private Function FindIndividualSheet(ByVal targetWorkbook As Workbook, ByVal featureName As String) As Worksheet
    ' 個別シート名: 「【個別】」 + 機能名（完全一致）
    Dim targetSheetName As String
    Dim ws As Worksheet

    targetSheetName = SHEET_KEY_INDIVIDUAL_PREFIX & featureName

    For Each ws In targetWorkbook.Worksheets
        If StrComp(ws.Name, targetSheetName, vbTextCompare) = 0 Then
            Set FindIndividualSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Sub MarkCurrentSourceSheet( _
    ByVal sourceSheet As Worksheet, _
    ByRef sourceTextValues As Variant, _
    ByVal lastRow As Long, _
    ByRef markedCount As Long, _
    ByVal leadingFunctionStartsAtB1 As Boolean)

    ' 現行ソースシートのC列を走査し、対象構文に応じてB列へセクション番号を設定する
    ' - function 行      : B(次セクション番号)   例: B2（先頭function開始時は設定でB1）
    ' - その他の対象構文 : B(現セクション番号)- 例: B1-
    Dim rowIndex As Long
    Dim lineText As String
    Dim currentSectionIndex As Long

    markedCount = 0
    If leadingFunctionStartsAtB1 Then
        currentSectionIndex = 0
    Else
        currentSectionIndex = 1
    End If

    For rowIndex = 1 To lastRow
        lineText = GetCellTextFromValue(sourceTextValues(rowIndex, 1))

        ' コメント行（# / // 先頭）は書き込み対象外
        If IsCommentLine(lineText) Then
            GoTo ContinueMarkLoop
        End If

        ' HTMLの開始タグを検出したら、それ以降の行は検索しない
        If IsSourceSearchStopLine(lineText) Then
            Exit For
        End If

        If IsFunctionLine(lineText) Then
            currentSectionIndex = currentSectionIndex + 1
            sourceSheet.Cells(rowIndex, MARK_COL).Value = "B" & CStr(currentSectionIndex)
            markedCount = markedCount + 1
        ElseIf IsMarkTargetLine(lineText) Then
            sourceSheet.Cells(rowIndex, MARK_COL).Value = "B" & CStr(currentSectionIndex) & "-"
            markedCount = markedCount + 1
        End If

ContinueMarkLoop:
    Next rowIndex
End Sub

Private Function ShouldStartFunctionSectionFromB1( _
    ByRef sourceTextValues As Variant, _
    ByVal lastRow As Long) As Boolean

    ' 最初の判定対象構文がfunctionなら、先頭セクションをB1から開始する
    Dim rowIndex As Long
    Dim lineText As String

    If Not LEADING_FUNCTION_STARTS_FROM_B1 Then Exit Function

    For rowIndex = 1 To lastRow
        lineText = GetCellTextFromValue(sourceTextValues(rowIndex, 1))

        If IsCommentLine(lineText) Then GoTo ContinueLeadingCheck
        If IsSourceSearchStopLine(lineText) Then Exit For
        If Len(Trim$(lineText)) = 0 Then GoTo ContinueLeadingCheck

        If IsFunctionLine(lineText) Then
            ShouldStartFunctionSectionFromB1 = True
            Exit Function
        End If

        If IsMarkTargetLine(lineText) Then
            Exit Function
        End If

ContinueLeadingCheck:
    Next rowIndex
End Function

Private Function CollectSyntaxEvents( _
    ByRef sourceTextValues As Variant, _
    ByVal lastRow As Long) As Collection

    ' 個別シート出力用に、現行ソースシートの構文イベントを上から順に収集する
    ' 文字列ベース判定（部分一致）
    Dim events As Collection
    Dim rowIndex As Long
    Dim lineText As String
    Dim eventItem As Collection
    Dim switchEndRow As Long

    Set events = New Collection
    rowIndex = 1

    Do While rowIndex <= lastRow
        lineText = GetCellTextFromValue(sourceTextValues(rowIndex, 1))

        ' コメント行（# / // 先頭）は個別シート出力の解析対象外
        If IsCommentLine(lineText) Then
            rowIndex = rowIndex + 1
            GoTo ContinueLoop
        End If

        ' HTMLの開始タグを検出したら、それ以降の行は検索しない
        If IsSourceSearchStopLine(lineText) Then
            Exit Do
        End If

        If Len(Trim$(lineText)) = 0 Then
            rowIndex = rowIndex + 1
            GoTo ContinueLoop
        End If

        ' function を最優先で判定（新しい処理セクション開始のため）
        If IsFunctionLine(lineText) Then
            Set eventItem = CreateFunctionEvent(ParseFunctionName(lineText))
            events.Add eventItem

        ' switch は後続行の case/default を収集するので、まとめてイベント化する
        ElseIf IsSwitchLine(lineText) Then
            Set eventItem = CollectSwitchEvent(sourceTextValues, rowIndex, lastRow, switchEndRow)
            events.Add eventItem
            If switchEndRow > rowIndex Then
                rowIndex = switchEndRow
            End If

        ' else / else if / default / case の単体行は、個別シート出力をしない
        ElseIf IsElseIfLine(lineText) Then
            ' 個別シート出力なし
        ElseIf IsElseLine(lineText) Then
            ' 個別シート出力なし
        ElseIf IsCaseLine(lineText) Then
            ' switch収集中に扱う想定
        ElseIf IsDefaultLine(lineText) Then
            ' switch収集中に扱う想定

        ' 通常ブロックの判定（順序に注意: foreach を for より先に判定）
        ElseIf IsForeachLine(lineText) Then
            Set eventItem = CreateNormalSyntaxEvent("FOREACH")
            events.Add eventItem
        ElseIf IsForLine(lineText) Then
            Set eventItem = CreateNormalSyntaxEvent("FOR")
            events.Add eventItem
        ElseIf IsWhileLine(lineText) Then
            Set eventItem = CreateNormalSyntaxEvent("WHILE")
            events.Add eventItem
        ElseIf IsTernaryLine(lineText) Then
            Set eventItem = CreateNormalSyntaxEvent("TERNARY")
            events.Add eventItem
        ElseIf IsIfLine(lineText) Then
            Set eventItem = CreateNormalSyntaxEvent("IF")
            events.Add eventItem
        End If

        rowIndex = rowIndex + 1

ContinueLoop:
    Loop

    Set CollectSyntaxEvents = events
End Function

Private Sub WriteIndividualSheet(ByVal individualSheet As Worksheet, ByVal syntaxEvents As Collection, ByVal leadingFunctionStartsAtB1 As Boolean)
    ' 個別シートへ、仕様の書式で処理セクション/確認ブロックを書き込む
    On Error GoTo ErrorHandler

    Dim sectionIndex As Long
    Dim nextBlockStartRow As Long
    Dim i As Long
    Dim eventItem As Collection
    Dim eventKind As String
    Dim functionName As String
    Dim startFromFunctionAtB1 As Boolean
    Dim plannedLastWriteRow As Long
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String

    ' 追記処理の開始前に、個別シート15行目をテンプレートとして退避しておく
    PrepareIndividualSheetTemplateSnapshot individualSheet

    startFromFunctionAtB1 = (leadingFunctionStartsAtB1 And IsFirstSyntaxEventFunction(syntaxEvents))
    plannedLastWriteRow = EstimateLastWriteRow(syntaxEvents, startFromFunctionAtB1)
    EnsureIndividualSheetWritableCapacity individualSheet, plannedLastWriteRow
    mPreAllocatedWritableLastRow = plannedLastWriteRow

    ' 初期値（固定）
    If startFromFunctionAtB1 Then
        sectionIndex = 0
        nextBlockStartRow = SECTION_HEADER_START_ROW
    Else
        sectionIndex = 1
        WriteSectionHeader individualSheet, SECTION_HEADER_START_ROW, sectionIndex, "MAIN"
        nextBlockStartRow = SECTION_HEADER_START_ROW + 1
    End If

    For i = 1 To syntaxEvents.Count
        Set eventItem = syntaxEvents.item(i)
        eventKind = UCase$(EventText(eventItem, "Kind"))

        Select Case eventKind
            Case "FUNCTION"
                ' functionを検出したら、新しい処理セクション見出しを開始
                sectionIndex = sectionIndex + 1
                functionName = EventText(eventItem, "FunctionName", "UNKNOWN")
                WriteSectionHeader individualSheet, nextBlockStartRow, sectionIndex, functionName
                nextBlockStartRow = nextBlockStartRow + 1

            Case "SWITCH"
                EnsureSectionHeaderStarted individualSheet, sectionIndex, nextBlockStartRow
                nextBlockStartRow = WriteSwitchBlock(individualSheet, nextBlockStartRow, eventItem)

            Case "IF", "TERNARY", "FOR", "FOREACH", "WHILE"
                EnsureSectionHeaderStarted individualSheet, sectionIndex, nextBlockStartRow
                WriteNormalBlock individualSheet, nextBlockStartRow, eventItem
                nextBlockStartRow = nextBlockStartRow + BLOCK_STEP_NORMAL
        End Select
    Next i

    ClearIndividualSheetTemplateSnapshot
    mPreAllocatedWritableLastRow = 0
    Exit Sub

ErrorHandler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.Source

    ClearIndividualSheetTemplateSnapshot
    mPreAllocatedWritableLastRow = 0

    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Function IsFirstSyntaxEventFunction(ByVal syntaxEvents As Collection) As Boolean
    Dim firstEvent As Collection

    If syntaxEvents Is Nothing Then Exit Function
    If syntaxEvents.Count = 0 Then Exit Function

    Set firstEvent = syntaxEvents.item(1)
    IsFirstSyntaxEventFunction = (UCase$(EventText(firstEvent, "Kind")) = "FUNCTION")
End Function

Private Sub EnsureSectionHeaderStarted( _
    ByVal ws As Worksheet, _
    ByRef sectionIndex As Long, _
    ByRef nextBlockStartRow As Long)

    If sectionIndex > 0 Then Exit Sub

    sectionIndex = 1
    WriteSectionHeader ws, nextBlockStartRow, sectionIndex, "MAIN"
    nextBlockStartRow = nextBlockStartRow + 1
End Sub

Private Function EstimateLastWriteRow( _
    ByVal syntaxEvents As Collection, _
    ByVal startFromFunctionAtB1 As Boolean) As Long

    Dim sectionIndex As Long
    Dim nextBlockStartRow As Long
    Dim i As Long
    Dim eventItem As Collection
    Dim eventKind As String
    Dim maxRow As Long
    Dim switchLastRow As Long

    If startFromFunctionAtB1 Then
        sectionIndex = 0
        nextBlockStartRow = SECTION_HEADER_START_ROW
    Else
        sectionIndex = 1
        nextBlockStartRow = SECTION_HEADER_START_ROW + 1
        maxRow = SECTION_HEADER_START_ROW
    End If

    For i = 1 To syntaxEvents.Count
        Set eventItem = syntaxEvents.item(i)
        eventKind = UCase$(EventText(eventItem, "Kind"))

        Select Case eventKind
            Case "FUNCTION"
                sectionIndex = sectionIndex + 1
                If nextBlockStartRow > maxRow Then
                    maxRow = nextBlockStartRow
                End If
                nextBlockStartRow = nextBlockStartRow + 1

            Case "SWITCH"
                If sectionIndex = 0 Then
                    sectionIndex = 1
                    If nextBlockStartRow > maxRow Then
                        maxRow = nextBlockStartRow
                    End If
                    nextBlockStartRow = nextBlockStartRow + 1
                End If

                switchLastRow = EstimateSwitchBlockLastRow(eventItem, nextBlockStartRow)
                If switchLastRow > maxRow Then
                    maxRow = switchLastRow
                End If
                nextBlockStartRow = switchLastRow + 2

            Case "IF", "TERNARY", "FOR", "FOREACH", "WHILE"
                If sectionIndex = 0 Then
                    sectionIndex = 1
                    If nextBlockStartRow > maxRow Then
                        maxRow = nextBlockStartRow
                    End If
                    nextBlockStartRow = nextBlockStartRow + 1
                End If

                If (nextBlockStartRow + 3) > maxRow Then
                    maxRow = nextBlockStartRow + 3
                End If
                nextBlockStartRow = nextBlockStartRow + BLOCK_STEP_NORMAL
        End Select
    Next i

    EstimateLastWriteRow = maxRow
End Function

Private Function EstimateSwitchBlockLastRow(ByVal eventItem As Collection, ByVal startRow As Long) As Long
    Dim caseValues As Collection
    Dim caseCount As Long

    Set caseValues = EventCollection(eventItem, "CaseValues")
    If Not caseValues Is Nothing Then
        caseCount = caseValues.Count
    End If

    ' switch本体のヘッダ(startRow) + 分岐行(startRow+1, +2刻み)
    EstimateSwitchBlockLastRow = startRow + 1 + (caseCount * 2)
End Function

Private Sub EnsureIndividualSheetWritableCapacity(ByVal ws As Worksheet, ByVal requiredLastRow As Long)
    ' 書き込み予定の最終行が分かっている場合、必要な行挿入を先にまとめて実施する
    Dim alphaRow As Long
    Dim triggerStartRow As Long
    Dim insertAtRow As Long

    If requiredLastRow < SECTION_HEADER_START_ROW Then Exit Sub

    Do
        alphaRow = FindAlphaRow(ws, requiredLastRow)
        If alphaRow = 0 Then Exit Sub

        triggerStartRow = alphaRow - INDIVIDUAL_TEMPLATE_PREINSERT_MARGIN
        If triggerStartRow < 1 Then
            triggerStartRow = 1
        End If

        If requiredLastRow < triggerStartRow Then
            Exit Do
        End If

        insertAtRow = alphaRow - 10
        If insertAtRow < 1 Then
            insertAtRow = 1
        End If

        InsertTemplateRowsChunk ws, insertAtRow, INDIVIDUAL_TEMPLATE_INSERT_COUNT
        Application.CutCopyMode = False
    Loop
End Sub

Private Sub WriteSectionHeader(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal sectionIndex As Long, ByVal sectionName As String)
    ' 処理セクション見出し行の出力
    EnsureIndividualSheetWritableRow ws, headerRow

    ws.Range("A" & CStr(headerRow)).Value = "B" & CStr(sectionIndex)
    ws.Range("K" & CStr(headerRow)).Value = SYMBOL_FILLED
    ws.Range("M" & CStr(headerRow)).Value = "処理（" & sectionName & "）"
End Sub

Private Sub WriteNormalBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal eventItem As Collection)
    ' if / 三項演算子 / for / foreach / while の共通形式
    EnsureIndividualSheetWritableRow ws, startRow
    ws.Range("E" & CStr(startRow)).Value = "***"
    ws.Range("L" & CStr(startRow)).Value = SYMBOL_EMPTY
    ws.Range("N" & CStr(startRow)).Value = EventText(eventItem, "Title")

    EnsureIndividualSheetWritableRow ws, startRow + 1
    ws.Range("H" & CStr(startRow + 1)).Value = 1
    ws.Range("M" & CStr(startRow + 1)).Value = SYMBOL_EMPTY
    ws.Range("O" & CStr(startRow + 1)).Value = EventText(eventItem, "Cond1")
    ws.Range("AX" & CStr(startRow + 1)).Value = SYMBOL_EMPTY
    ws.Range("AZ" & CStr(startRow + 1)).Value = EventText(eventItem, "Result1")
    ws.Range("CF" & CStr(startRow + 1)).Value = "1,4"

    EnsureIndividualSheetWritableRow ws, startRow + 3
    ws.Range("H" & CStr(startRow + 3)).Value = 2
    ws.Range("M" & CStr(startRow + 3)).Value = SYMBOL_EMPTY
    ws.Range("O" & CStr(startRow + 3)).Value = EventText(eventItem, "Cond2")
    ws.Range("AX" & CStr(startRow + 3)).Value = SYMBOL_EMPTY
    ws.Range("AZ" & CStr(startRow + 3)).Value = EventText(eventItem, "Result2")
    ws.Range("CF" & CStr(startRow + 3)).Value = "1,4"
End Sub

Private Function WriteSwitchBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal eventItem As Collection) As Long
    ' switchブロック:
    ' - ヘッダ1行
    ' - 分岐行を 2行おき（r+1, r+3, ...）に配置
    ' - 最後に「上記のいずれでもない場合」を追加
    ' 戻り値は次の確認ブロック開始行
    Dim switchArg As String
    Dim titleText As String
    Dim caseValues As Collection
    Dim hasDefault As Boolean
    Dim branchRow As Long
    Dim seqNo As Long
    Dim i As Long
    Dim caseValue As String
    Dim lastUsedBranchRow As Long

    switchArg = EventText(eventItem, "SwitchArg", "UNKNOWN")
    titleText = "条件分岐（SWITCH文（" & switchArg & "の値））の確認"
    hasDefault = EventFlag(eventItem, "HasDefault")
    Set caseValues = EventCollection(eventItem, "CaseValues")

    EnsureIndividualSheetWritableRow ws, startRow
    ws.Range("E" & CStr(startRow)).Value = "***"
    ws.Range("L" & CStr(startRow)).Value = SYMBOL_EMPTY
    ws.Range("N" & CStr(startRow)).Value = titleText

    branchRow = startRow + 1
    seqNo = 1
    lastUsedBranchRow = startRow

    If Not caseValues Is Nothing Then
        For i = 1 To caseValues.Count
            caseValue = CStr(caseValues.item(i))
            WriteSwitchBranchRow ws, branchRow, seqNo, caseValue & "の場合", "CASE内の処理が行われること"
            lastUsedBranchRow = branchRow
            seqNo = seqNo + 1
            branchRow = branchRow + 2
        Next i
    End If

    ' 追加ケース（いずれでもない場合）
    If hasDefault Then
        WriteSwitchBranchRow ws, branchRow, seqNo, "上記のいずれでもない場合", "DEFAULT内の処理が行われること"
    Else
        WriteSwitchBranchRow ws, branchRow, seqNo, "上記のいずれでもない場合", "CASE内の処理が行われないこと"
    End If
    lastUsedBranchRow = branchRow

    ' 次の開始行は、最後に使った分岐行の2行後（1行空け）
    WriteSwitchBlock = lastUsedBranchRow + 2
End Function

Private Sub WriteSwitchBranchRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal seqNo As Long, ByVal conditionText As String, ByVal expectedText As String)
    ' switchの分岐行（case/default相当）の共通出力
    EnsureIndividualSheetWritableRow ws, rowIndex

    ws.Range("H" & CStr(rowIndex)).Value = seqNo
    ws.Range("M" & CStr(rowIndex)).Value = SYMBOL_EMPTY
    ws.Range("O" & CStr(rowIndex)).Value = conditionText
    ws.Range("AX" & CStr(rowIndex)).Value = SYMBOL_EMPTY
    ws.Range("AZ" & CStr(rowIndex)).Value = expectedText
    ws.Range("CF" & CStr(rowIndex)).Value = "1,4"
End Sub

Private Sub EnsureIndividualSheetWritableRow(ByVal ws As Worksheet, ByVal rowIndex As Long)
    ' A:Dが未結合の最初の行をα行とし、
    ' 書き込み予定行が α-50 ～ α に入る場合は事前退避テンプレートを50行挿入する
    ' 事前確保済み行以内は判定を省略して高速化する
    Dim alphaRow As Long
    Dim triggerStartRow As Long
    Dim insertAtRow As Long

    If rowIndex < SECTION_HEADER_START_ROW Then Exit Sub

    If mPreAllocatedWritableLastRow > 0 Then
        If rowIndex <= mPreAllocatedWritableLastRow Then Exit Sub
    End If

    alphaRow = FindAlphaRow(ws, rowIndex)
    If alphaRow = 0 Then Exit Sub

    triggerStartRow = alphaRow - INDIVIDUAL_TEMPLATE_PREINSERT_MARGIN
    If triggerStartRow < 1 Then
        triggerStartRow = 1
    End If

    If rowIndex >= triggerStartRow And rowIndex <= alphaRow Then
        insertAtRow = alphaRow - 10
        If insertAtRow < 1 Then
            insertAtRow = 1
        End If

        InsertTemplateRowsChunk ws, insertAtRow, INDIVIDUAL_TEMPLATE_INSERT_COUNT
        Application.CutCopyMode = False
    End If
End Sub

Private Function IsIndividualSheetADMerged(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    ' A列セルの結合範囲が、その行の A:D と「ちょうど一致」するか判定
    Dim firstCell As Range

    Set firstCell = ws.Range("A" & CStr(rowIndex))

    If Not firstCell.MergeCells Then Exit Function

    With firstCell.MergeArea
        IsIndividualSheetADMerged = (.Row = rowIndex And .Column = 1 And .Rows.Count = 1 And .Columns.Count = 4)
    End With
End Function

Private Function FindAlphaRow(ByVal ws As Worksheet, ByVal writeTargetRow As Long) As Long
    ' α行 = A:D が未結合の最初の行
    Dim usedLastRow As Long
    Dim searchEndRow As Long
    Dim r As Long

    usedLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    searchEndRow = writeTargetRow + INDIVIDUAL_TEMPLATE_PREINSERT_MARGIN
    If searchEndRow < usedLastRow Then
        searchEndRow = usedLastRow
    End If
    If searchEndRow < SECTION_HEADER_START_ROW Then
        searchEndRow = SECTION_HEADER_START_ROW
    End If

    For r = SECTION_HEADER_START_ROW To searchEndRow
        If Not IsIndividualSheetADMerged(ws, r) Then
            FindAlphaRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub InsertTemplateRowsChunk(ByVal ws As Worksheet, ByVal insertAtRow As Long, ByVal insertCount As Long)
    ' 退避済みテンプレート（個別シート15行目から10行）を使って、
    ' insertAtRow へ insertCount 行分のテンプレートを挿入する
    Dim copyCount As Long
    Dim insertedCount As Long

    If insertAtRow < 1 Then Exit Sub
    If insertCount < 1 Then Exit Sub

    If mTemplateSnapshotSheet Is Nothing Then
        Err.Raise vbObjectError + 1001, "InsertTemplateRowsChunk", "テンプレート行の退避データがありません。"
    End If

    insertedCount = 0
    Do While insertedCount < insertCount
        copyCount = insertCount - insertedCount
        If copyCount > INDIVIDUAL_TEMPLATE_SOURCE_ROW_COUNT Then
            copyCount = INDIVIDUAL_TEMPLATE_SOURCE_ROW_COUNT
        End If

        mTemplateSnapshotSheet.Rows("1:" & CStr(copyCount)).Copy
        ws.Rows(CStr(insertAtRow + insertedCount) & ":" & CStr(insertAtRow + insertedCount + copyCount - 1)).Insert Shift:=xlDown
        insertedCount = insertedCount + copyCount
    Loop
End Sub

Private Sub PrepareIndividualSheetTemplateSnapshot(ByVal individualSheet As Worksheet)
    ' 個別シート15行目から10行を、後続の行挿入用テンプレートとして一時シートへ退避する
    Dim wb As Workbook
    Dim snapshotSheet As Worksheet
    Dim sourceStartRow As Long
    Dim sourceEndRow As Long

    ClearIndividualSheetTemplateSnapshot

    Set wb = individualSheet.Parent
    Set snapshotSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    snapshotSheet.Name = BuildTemplateSnapshotSheetName(wb)

    sourceStartRow = INDIVIDUAL_TEMPLATE_SOURCE_START_ROW
    sourceEndRow = sourceStartRow + INDIVIDUAL_TEMPLATE_SOURCE_ROW_COUNT - 1

    individualSheet.Rows(CStr(sourceStartRow) & ":" & CStr(sourceEndRow)).Copy
    snapshotSheet.Rows("1:" & CStr(INDIVIDUAL_TEMPLATE_SOURCE_ROW_COUNT)).PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    snapshotSheet.Visible = xlSheetVeryHidden
    Set mTemplateSnapshotSheet = snapshotSheet
End Sub

Private Sub ClearIndividualSheetTemplateSnapshot()
    ' 一時テンプレートシートを削除する
    Dim previousDisplayAlerts As Boolean

    If mTemplateSnapshotSheet Is Nothing Then Exit Sub

    previousDisplayAlerts = Application.DisplayAlerts

    On Error Resume Next
    Application.DisplayAlerts = False
    mTemplateSnapshotSheet.Visible = xlSheetVisible
    mTemplateSnapshotSheet.Delete
    Set mTemplateSnapshotSheet = Nothing
    Application.DisplayAlerts = previousDisplayAlerts
    On Error GoTo 0
End Sub

Private Function BuildTemplateSnapshotSheetName(ByVal wb As Workbook) As String
    ' 一時シート名を衝突しにくい形で生成する（31文字以内）
    Dim baseName As String
    Dim candidateName As String
    Dim suffixNo As Long

    baseName = TEMPLATE_SNAPSHOT_SHEET_PREFIX & Format$(Now, "hhnnss")
    candidateName = Left$(baseName, 31)
    suffixNo = 1

    Do While WorksheetExistsByName(wb, candidateName)
        candidateName = Left$(TEMPLATE_SNAPSHOT_SHEET_PREFIX & Format$(Now, "hhnnss") & "_" & CStr(suffixNo), 31)
        suffixNo = suffixNo + 1
    Loop

    BuildTemplateSnapshotSheetName = candidateName
End Function

Private Function WorksheetExistsByName(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            WorksheetExistsByName = True
            Exit Function
        End If
    Next ws
End Function

Private Function CreateNormalSyntaxEvent(ByVal syntaxKind As String) As Collection
    ' 通常ブロック（if / ternary / for / foreach / while）の文言をまとめたイベント
    Dim ev As Collection
    Dim kindUpper As String

    kindUpper = UCase$(Trim$(syntaxKind))
    Set ev = NewEvent(kindUpper)

    Select Case kindUpper
        Case "IF"
            ev.Add "条件分岐（IF文）の確認", "Title"
            ev.Add "条件が成立する場合", "Cond1"
            ev.Add "IF内の処理が行われること", "Result1"
            ev.Add "条件が成立しない場合", "Cond2"
            ev.Add "IF内の処理が行われないこと", "Result2"

        Case "TERNARY"
            ev.Add "条件分岐（三項演算子）の確認", "Title"
            ev.Add "条件が成立する場合", "Cond1"
            ev.Add "真である場合の処理が行われること", "Result1"
            ev.Add "条件が成立しない場合", "Cond2"
            ev.Add "偽である場合の処理が行われること", "Result2"

        Case "FOR"
            ev.Add "ループ（FOR文）の確認", "Title"
            ev.Add "ループ条件が成立する場合", "Cond1"
            ev.Add "ループ内の処理が行われること", "Result1"
            ev.Add "ループ条件が成立しない場合", "Cond2"
            ev.Add "ループ処理を抜けて、以降の処理が行われること", "Result2"

        Case "FOREACH"
            ev.Add "ループ（FOREACH文）の確認", "Title"
            ev.Add "ループ条件が成立する場合", "Cond1"
            ev.Add "ループ内の処理が行われること", "Result1"
            ev.Add "ループ条件が成立しない場合", "Cond2"
            ev.Add "ループ処理を抜けて、以降の処理が行われること", "Result2"

        Case "WHILE"
            ev.Add "ループ（WHILE文）の確認", "Title"
            ev.Add "ループ条件が成立する場合", "Cond1"
            ev.Add "ループ内の処理が行われること", "Result1"
            ev.Add "ループ条件が成立しない場合", "Cond2"
            ev.Add "ループ処理を抜けて、以降の処理が行われること", "Result2"

        Case Else
            ' 想定外の種類が来ても最低限の形で返す
            ev.Add "確認", "Title"
            ev.Add "条件1", "Cond1"
            ev.Add "期待結果1", "Result1"
            ev.Add "条件2", "Cond2"
            ev.Add "期待結果2", "Result2"
    End Select

    Set CreateNormalSyntaxEvent = ev
End Function

Private Function CreateFunctionEvent(ByVal functionName As String) As Collection
    ' function検出イベント（個別シートでは新しい処理セクション開始に使用）
    Dim ev As Collection
    Set ev = NewEvent("FUNCTION")
    ev.Add functionName, "FunctionName"
    Set CreateFunctionEvent = ev
End Function

Private Function CollectSwitchEvent( _
    ByRef sourceTextValues As Variant, _
    ByVal switchRow As Long, _
    ByVal lastRow As Long, _
    ByRef endRow As Long) As Collection

    ' switch行を起点に、後続のcase/defaultを簡易的に収集して1イベントにまとめる
    ' 終端判定は厳密にせず、以下のような簡易条件で打ち切る:
    ' - 次のfunctionが来た
    ' - 次のswitchが来た
    ' - case/defaultを拾った後に空行が2行続いた
    Dim ev As Collection
    Dim caseValues As Collection
    Dim switchArg As String
    Dim hasDefault As Boolean
    Dim r As Long
    Dim lineText As String
    Dim trimmedText As String
    Dim blankStreak As Long
    Dim foundBranch As Boolean
    Dim parsedCase As String

    Set ev = NewEvent("SWITCH")
    Set caseValues = New Collection
    switchArg = ParseSwitchArgument(GetCellTextFromValue(sourceTextValues(switchRow, 1)))

    For r = switchRow + 1 To lastRow
        lineText = GetCellTextFromValue(sourceTextValues(r, 1))
        trimmedText = Trim$(lineText)

        ' コメント行（# / // 先頭）はswitch収集対象外
        If IsCommentLine(lineText) Then
            blankStreak = 0
            GoTo ContinueSwitchLoop
        End If

        ' 現行ソースシートの検索打ち切り条件（HTML開始タグ）
        If IsSourceSearchStopLine(lineText) Then
            Exit For
        End If

        ' 次のfunction / switch は次の構文として扱いたいので、ここで打ち切る
        If IsFunctionLine(lineText) Then
            Exit For
        End If
        If IsSwitchLine(lineText) Then
            Exit For
        End If

        If Len(trimmedText) = 0 Then
            blankStreak = blankStreak + 1
            If foundBranch And blankStreak >= 2 Then
                Exit For
            End If
        Else
            blankStreak = 0

            If IsCaseLine(lineText) Then
                parsedCase = ParseCaseValue(lineText)
                caseValues.Add parsedCase
                foundBranch = True
            ElseIf IsDefaultLine(lineText) Then
                hasDefault = True
                foundBranch = True
            End If
        End If

ContinueSwitchLoop:
    Next r

    ev.Add switchArg, "SwitchArg"
    ev.Add hasDefault, "HasDefault"
    ev.Add caseValues, "CaseValues"

    ' endRow は、外側ループで再判定したくない範囲の最後の行
    If r > lastRow Then
        endRow = lastRow
    Else
        endRow = r - 1
        If endRow < switchRow Then
            endRow = switchRow
        End If
    End If

    Set CollectSwitchEvent = ev
End Function

Private Function ParseFunctionName(ByVal lineText As String) As String
    ' `function` の後ろの識別子を簡易抽出
    ' 例: function foo(XXX){  -> foo
    Dim posFunction As Long
    Dim restText As String
    Dim i As Long
    Dim ch As String
    Dim nameBuffer As String

    posFunction = InStr(1, lineText, "function", vbTextCompare)
    If posFunction = 0 Then
        ParseFunctionName = "UNKNOWN"
        Exit Function
    End If

    restText = Mid$(lineText, posFunction + Len("function"))
    restText = Trim$(restText)

    If Len(restText) = 0 Then
        ParseFunctionName = "UNKNOWN"
        Exit Function
    End If

    For i = 1 To Len(restText)
        ch = Mid$(restText, i, 1)

        If ch = "(" Or ch = " " Or ch = vbTab Or ch = "{" Then
            If Len(nameBuffer) > 0 Then
                Exit For
            End If
        Else
            nameBuffer = nameBuffer & ch
        End If
    Next i

    If Len(nameBuffer) = 0 Then
        ParseFunctionName = "UNKNOWN"
    Else
        ParseFunctionName = nameBuffer
    End If
End Function

Private Function ParseSwitchArgument(ByVal lineText As String) As String
    ' `switch(YYY)` の括弧内を簡易抽出
    Dim posSwitch As Long
    Dim posOpen As Long
    Dim posClose As Long
    Dim argText As String

    posSwitch = InStr(1, lineText, "switch", vbTextCompare)
    If posSwitch = 0 Then
        ParseSwitchArgument = "UNKNOWN"
        Exit Function
    End If

    posOpen = InStr(posSwitch, lineText, "(", vbBinaryCompare)
    If posOpen = 0 Then
        ParseSwitchArgument = "UNKNOWN"
        Exit Function
    End If

    posClose = InStr(posOpen + 1, lineText, ")", vbBinaryCompare)
    If posClose = 0 Then
        ParseSwitchArgument = "UNKNOWN"
        Exit Function
    End If

    argText = Mid$(lineText, posOpen + 1, posClose - posOpen - 1)
    argText = Trim$(argText)

    If Len(argText) = 0 Then
        ParseSwitchArgument = "UNKNOWN"
    Else
        ParseSwitchArgument = argText
    End If
End Function

Private Function ParseCaseValue(ByVal lineText As String) As String
    ' `case XXX:` の XXX 部分を簡易抽出
    Dim posCase As Long
    Dim restText As String
    Dim posColon As Long

    posCase = InStr(1, lineText, "case", vbTextCompare)
    If posCase = 0 Then
        ParseCaseValue = "UNKNOWN"
        Exit Function
    End If

    restText = Mid$(lineText, posCase + Len("case"))
    posColon = InStr(1, restText, ":", vbBinaryCompare)
    If posColon > 0 Then
        restText = Left$(restText, posColon - 1)
    End If

    restText = Trim$(restText)
    If Len(restText) = 0 Then
        ParseCaseValue = "UNKNOWN"
    Else
        ParseCaseValue = restText
    End If
End Function

Private Function GetLastRow(ByVal ws As Worksheet, ByVal columnIndex As Long) As Long
    ' 指定列の最終行を返す（列が空でも最低1を返す）
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
    If lastRow < 1 Then
        lastRow = 1
    End If

    GetLastRow = lastRow
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

    rawValues = ws.Range(ws.Cells(startRow, columnIndex), ws.Cells(endRow, columnIndex)).Value

    If startRow = endRow Then
        singleCell(1, 1) = rawValues
        ReadColumnValues = singleCell
    Else
        ReadColumnValues = rawValues
    End If
End Function

Private Function GetCellTextFromValue(ByVal cellValue As Variant) As String
    ' エラー値を安全に文字列化するためのヘルパー
    On Error GoTo SafeExit

    If IsError(cellValue) Then
        GetCellTextFromValue = vbNullString
    ElseIf IsEmpty(cellValue) Then
        GetCellTextFromValue = vbNullString
    Else
        GetCellTextFromValue = CStr(cellValue)
    End If
    Exit Function

SafeExit:
    GetCellTextFromValue = vbNullString
End Function

Private Function IsCommentLine(ByVal lineText As String) As Boolean
    ' 先頭（前方空白を除去後）が # または // の行をコメントとみなす
    Dim normalizedText As String

    normalizedText = LTrim$(lineText)

    If Len(normalizedText) = 0 Then Exit Function

    If Left$(normalizedText, 1) = "#" Then
        IsCommentLine = True
        Exit Function
    End If

    If Len(normalizedText) >= 2 Then
        IsCommentLine = (Left$(normalizedText, 2) = "//")
    End If
End Function

Private Function IsMarkTargetLine(ByVal lineText As String) As Boolean
    ' 現行ソースシートのB列マーキング対象
    ' ※ 文字列ベースの部分一致判定を採用
    If Len(Trim$(lineText)) = 0 Then
        Exit Function
    End If

    If IsElseIfLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsElseLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsIfLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsTernaryLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsForeachLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsForLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsWhileLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsSwitchLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsCaseLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsDefaultLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If

    If IsFunctionLine(lineText) Then
        IsMarkTargetLine = True
        Exit Function
    End If
End Function

Private Function IsIfLine(ByVal lineText As String) As Boolean
    ' else if は別扱いなので除外
    If IsElseIfLine(lineText) Then Exit Function
    IsIfLine = ContainsWholeWord(lineText, "if")
End Function

Private Function IsElseIfLine(ByVal lineText As String) As Boolean
    ' 「else if」に加えて「elseif」も対象にする
    If ContainsWholeWord(lineText, "elseif") Then
        IsElseIfLine = True
        Exit Function
    End If

    IsElseIfLine = ContainsText(lineText, "else if") And _
                   ContainsWholeWord(lineText, "else") And _
                   ContainsWholeWord(lineText, "if")
End Function

Private Function IsElseLine(ByVal lineText As String) As Boolean
    IsElseLine = ContainsWholeWord(lineText, "else")
End Function

Private Function IsTernaryLine(ByVal lineText As String) As Boolean
    IsTernaryLine = (InStr(1, lineText, "?", vbBinaryCompare) > 0 And _
                     InStr(1, lineText, ":", vbBinaryCompare) > 0)
End Function

Private Function IsForeachLine(ByVal lineText As String) As Boolean
    IsForeachLine = ContainsWholeWord(lineText, "foreach")
End Function

Private Function IsForLine(ByVal lineText As String) As Boolean
    ' foreach とは区別する
    If IsForeachLine(lineText) Then Exit Function
    IsForLine = ContainsWholeWord(lineText, "for")
End Function

Private Function IsWhileLine(ByVal lineText As String) As Boolean
    IsWhileLine = ContainsWholeWord(lineText, "while")
End Function

Private Function IsSwitchLine(ByVal lineText As String) As Boolean
    IsSwitchLine = ContainsWholeWord(lineText, "switch")
End Function

Private Function IsCaseLine(ByVal lineText As String) As Boolean
    ' 「case」を単語区切りで判定し、かつ「:」を含む行を対象にする
    IsCaseLine = ContainsWholeWord(lineText, "case") And (InStr(1, lineText, ":", vbBinaryCompare) > 0)
End Function

Private Function IsDefaultLine(ByVal lineText As String) As Boolean
    IsDefaultLine = ContainsWholeWord(lineText, "default") And (InStr(1, lineText, ":", vbBinaryCompare) > 0)
End Function

Private Function IsFunctionLine(ByVal lineText As String) As Boolean
    IsFunctionLine = ContainsWholeWord(lineText, "function")
End Function

Private Function IsSourceSearchStopLine(ByVal lineText As String) As Boolean
    ' HTML開始付近に入ったら、現行ソースシートの下側は検索しない
    IsSourceSearchStopLine = _
        ContainsText(lineText, "<!doctype") Or _
        ContainsText(lineText, "<html") Or _
        ContainsText(lineText, "<head")
End Function

Private Function ContainsWholeWord(ByVal sourceText As String, ByVal findWord As String) As Boolean
    ' 単語区切りで一致する場合のみTrue
    ' 例: "for" は "form" ではヒットしない
    Dim searchPos As Long
    Dim hitPos As Long
    Dim endPos As Long
    Dim beforeChar As String
    Dim afterChar As String

    If Len(findWord) = 0 Then Exit Function

    searchPos = 1

    Do
        hitPos = InStr(searchPos, sourceText, findWord, vbTextCompare)
        If hitPos = 0 Then Exit Do

        endPos = hitPos + Len(findWord) - 1
        beforeChar = vbNullString
        afterChar = vbNullString

        If hitPos > 1 Then
            beforeChar = Mid$(sourceText, hitPos - 1, 1)
        End If

        If endPos < Len(sourceText) Then
            afterChar = Mid$(sourceText, endPos + 1, 1)
        End If

        If (Not IsWordChar(beforeChar)) And (Not IsWordChar(afterChar)) Then
            ContainsWholeWord = True
            Exit Function
        End If

        searchPos = hitPos + 1
    Loop
End Function

Private Function IsWordChar(ByVal ch As String) As Boolean
    ' 単語構成文字（ASCII識別子系）を判定
    Dim codePoint As Long

    If Len(ch) = 0 Then Exit Function

    codePoint = AscW(ch)
    IsWordChar = _
        (codePoint >= 48 And codePoint <= 57) Or _
        (codePoint >= 65 And codePoint <= 90) Or _
        (codePoint >= 97 And codePoint <= 122) Or _
        (ch = "_") Or _
        (ch = "$")
End Function

Private Function ContainsText(ByVal sourceText As String, ByVal findText As String) As Boolean
    ' 大文字小文字を無視した部分一致
    If Len(findText) = 0 Then
        ContainsText = False
    Else
        ContainsText = (InStr(1, sourceText, findText, vbTextCompare) > 0)
    End If
End Function

Private Sub ActivateSheetForNextOpen(ByVal wb As Workbook, ByVal targetSheet As Worksheet)
    If wb Is Nothing Then Exit Sub
    If targetSheet Is Nothing Then Exit Sub
    If Not (targetSheet.Parent Is wb) Then Exit Sub

    On Error Resume Next
    wb.Activate
    targetSheet.Activate
    On Error GoTo 0
End Sub

Private Function NewEvent(ByVal eventKind As String) As Collection
    ' 疑似イベントオブジェクト（Collection + Key）を生成
    Dim ev As Collection
    Set ev = New Collection
    ev.Add eventKind, "Kind"
    Set NewEvent = ev
End Function

Private Function EventText(ByVal ev As Collection, ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    ' Collectionのキー取得（文字列）
    On Error GoTo UseDefault
    EventText = CStr(ev.item(keyName))
    Exit Function

UseDefault:
    EventText = defaultValue
End Function

Private Function EventFlag(ByVal ev As Collection, ByVal keyName As String) As Boolean
    ' Collectionのキー取得（Boolean）
    On Error GoTo UseFalse
    EventFlag = CBool(ev.item(keyName))
    Exit Function

UseFalse:
    EventFlag = False
End Function

Private Function EventCollection(ByVal ev As Collection, ByVal keyName As String) As Collection
    ' Collectionのキー取得（Collectionオブジェクト）
    On Error GoTo NoCollection
    Set EventCollection = ev.item(keyName)
    Exit Function

NoCollection:
    Set EventCollection = Nothing
End Function




