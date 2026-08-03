Attribute VB_Name = "EscapePartsMarking"
Option Explicit

Private Const DEFAULT_COMPLETION_MESSAGE As String = "SQLインジェクション対策済み"
Private Const ESCAPE_TARGET_PREFIXES_CSV As String = "pg_escape_string,sqlS,sqlN,sqlLS,sqlC,sqlNZ,sqlInN,sqlF,sqlChk,sqlLikeStr,sqlNum,sqlNum0,sqlStr"
Private Const OPTION_ONLY_A_VALUE_ROW_FILL_TARGET As String = "Both" ' None / Left / Right / Both
Private Const ONLY_A_VALUE_ROW_FILL_COLOR_HEX As String = "#a6a6a6"
Public Type EscapePartsMarkingUiOptions
    Enabled As Boolean
    TargetWorkbookPath As String
    UseCompletionMessage As Boolean
    completionMessage As String
    UseEscapeTargetPrefixesCsv As Boolean
    escapeTargetPrefixesCsv As String
    UseOnlyAValueRowFillTarget As Boolean
    onlyAValueRowFillTarget As String
    UseOnlyAValueRowFillColorHex As Boolean
    onlyAValueRowFillColorHex As String
End Type

Private mUiOptions As EscapePartsMarkingUiOptions

Public Sub RunMainWithUiOptions(ByRef options As EscapePartsMarkingUiOptions)
    ClearUiOptions
    mUiOptions = options
    mUiOptions.Enabled = True

    RunMain

    ClearUiOptions
End Sub

Public Function CreateEscapePartsMarkingUiOptionsForForm() As EscapePartsMarkingUiOptions
    Dim defaults As EscapePartsMarkingUiOptions

    defaults.Enabled = True
    defaults.TargetWorkbookPath = vbNullString

    defaults.UseCompletionMessage = True
    defaults.completionMessage = DEFAULT_COMPLETION_MESSAGE

    defaults.UseEscapeTargetPrefixesCsv = True
    defaults.escapeTargetPrefixesCsv = ESCAPE_TARGET_PREFIXES_CSV

    defaults.UseOnlyAValueRowFillTarget = True
    defaults.onlyAValueRowFillTarget = OPTION_ONLY_A_VALUE_ROW_FILL_TARGET

    defaults.UseOnlyAValueRowFillColorHex = True
    defaults.onlyAValueRowFillColorHex = ONLY_A_VALUE_ROW_FILL_COLOR_HEX

    CreateEscapePartsMarkingUiOptionsForForm = defaults
End Function

Private Sub ClearUiOptions()
    mUiOptions.Enabled = False
    mUiOptions.TargetWorkbookPath = vbNullString
    mUiOptions.UseCompletionMessage = False
    mUiOptions.completionMessage = vbNullString
    mUiOptions.UseEscapeTargetPrefixesCsv = False
    mUiOptions.escapeTargetPrefixesCsv = vbNullString
    mUiOptions.UseOnlyAValueRowFillTarget = False
    mUiOptions.onlyAValueRowFillTarget = vbNullString
    mUiOptions.UseOnlyAValueRowFillColorHex = False
    mUiOptions.onlyAValueRowFillColorHex = vbNullString
End Sub

'============================================================
' xlsmツール（別ファイル）から、選択した xlsx を開いて加工するマクロ
' - シート名に "A1-1-1" を含むシートのみを対象に処理する
' - A列のみ入力（B列空）の行はオプション値（None/Left/Right/Both）に応じて塗りつぶす
' - B列の "sqlX(...)" 部分だけを赤字＋太字（複数ヒット、複数行跨ぎ対応）
' - ただし "DbHelper.sqlX(...)" のようなクラス/インスタンス経由の呼び出しは
'   "DbHelper." も含めて赤字＋太字にする
' - ヒットした行の C列に固定メッセージ（既定: "SQLインジェクション対策済み"）を赤字で書く
'
' 前提:
'  - モジュール先頭の ESCAPE_TARGET_PREFIXES_CSV にエスケープ関数（例: pg_escape_string, sqlS, sqlN）を列挙していること
'============================================================
Public Sub RunMain()
    Dim targetPath As String
    targetPath = PickExcelFilePath()
    If targetPath = "" Then Exit Sub ' キャンセル

    Dim prefixes As Collection
    Set prefixes = LoadPrefixesFromCode()
    If prefixes.Count = 0 Then
        MsgBox "ESCAPE_TARGET_PREFIXES_CSV に prefix（例: pg_escape_string, sqlS, sqlN）を1つ以上設定してください。", vbExclamation
        Exit Sub
    End If

    Dim app As Application
    Set app = Application

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalc As XlCalculation

    ' 高速化＆事故防止（処理後に必ず戻す）
    prevScreenUpdating = app.ScreenUpdating
    prevEnableEvents = app.EnableEvents
    prevDisplayAlerts = app.DisplayAlerts
    prevCalc = app.Calculation

    app.ScreenUpdating = False
    app.EnableEvents = False
    app.DisplayAlerts = False
    app.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    Dim wb As Workbook
    Set wb = app.Workbooks.Open(Filename:=targetPath, ReadOnly:=False)

    '============================================================
    ' ★ シート名に "A1-1-1" を含むシートのみ処理する
    '============================================================
    Dim ws As Worksheet
    Dim processedSheetCount As Long

    For Each ws In wb.Worksheets
        If InStr(1, ws.Name, "A1-1-1", vbBinaryCompare) > 0 Then
            ProcessOneSheet ws, prefixes, ResolveCompletionMessage()
            processedSheetCount = processedSheetCount + 1
        End If
    Next ws

    If processedSheetCount = 0 Then
        wb.Close SaveChanges:=False
        MsgBox "対象ファイルに「A1-1-1」を含むシートが存在しませんでした。", vbExclamation
        GoTo CleanExit
    End If

    wb.Save
    wb.Close SaveChanges:=False

    MsgBox "完了しました。" & vbCrLf & _
           "「A1-1-1」を含むシートを更新しました（件数: " & CStr(processedSheetCount) & "）:" & vbCrLf & targetPath, vbInformation

CleanExit:
    app.ScreenUpdating = prevScreenUpdating
    app.EnableEvents = prevEnableEvents
    app.DisplayAlerts = prevDisplayAlerts
    app.Calculation = prevCalc
    Exit Sub

CleanFail:
    ' 例外時もExcel設定を戻す
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    On Error GoTo 0

    MsgBox "処理中にエラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'============================================================
' 1シート分処理:
' - A列のみ入力（B列空）の行はオプション値（None/Left/Right/Both）に応じて塗りつぶす
' - B列を走査して sqlX(...) / DbHelper.sqlX(...) を装飾（複数行跨ぎ対応）
' - ヒット行のC列に固定メッセージ＆赤字
'============================================================
Private Sub ProcessOneSheet(ByVal ws As Worksheet, ByVal prefixes As Collection, ByVal hitMessage As String)
    ' A/B列の最終行を取得（どちらにもデータが無ければスキップ）
    Dim lastRowA As Long
    Dim lastRowB As Long
    Dim lastRow As Long
    Dim onlyAFillColor As Long
    Dim fillTargetOption As String

    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    lastRow = lastRowA
    If lastRowB > lastRow Then lastRow = lastRowB
    If lastRow < 4 Then Exit Sub

    onlyAFillColor = ResolveOnlyAValueRowFillColor()
    fillTargetOption = ResolveOnlyAValueRowFillTargetOption()

    Dim r As Long
    For r = 4 To lastRow
        If ShouldFillOnlyAValueRow(ws, r) Then
            ApplyOnlyAValueRowFill ws, r, onlyAFillColor, fillTargetOption
        End If
    Next r

    MarkSqlPartsInRows ws, prefixes, hitMessage, 4, lastRow
End Sub

Private Sub MarkSqlPartsInRows( _
    ByVal ws As Worksheet, _
    ByVal prefixes As Collection, _
    ByVal hitMessage As String, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long)

    Dim rowCount As Long
    rowCount = lastRow - firstRow + 1
    If rowCount <= 0 Then Exit Sub

    Dim rowTexts() As String
    Dim rowStartPositions() As Long
    ReDim rowTexts(1 To rowCount)
    ReDim rowStartPositions(1 To rowCount)

    Dim virtualText As String
    Dim rowIndex As Long
    Dim rowNumber As Long

    For rowIndex = 1 To rowCount
        rowNumber = firstRow + rowIndex - 1
        rowTexts(rowIndex) = GetMarkingCellText(ws.Cells(rowNumber, "B").Value2)
        rowStartPositions(rowIndex) = Len(virtualText) + 1
        virtualText = virtualText & rowTexts(rowIndex)
        If rowIndex < rowCount Then virtualText = virtualText & vbLf
    Next rowIndex

    If InStr(1, virtualText, "(", vbBinaryCompare) = 0 Then Exit Sub

    Dim executableCodePositions() As Boolean
    BuildExecutableCodePositionMap virtualText, executableCodePositions

    Dim prefixIndex As Long
    Dim prefix As String

    For prefixIndex = 1 To prefixes.Count
        prefix = CStr(prefixes(prefixIndex))
        MarkAllOccurrencesForOnePrefixAcrossRows ws, virtualText, executableCodePositions, rowTexts, rowStartPositions, firstRow, prefix, hitMessage
    Next prefixIndex
End Sub

Private Function GetMarkingCellText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function

    GetMarkingCellText = CStr(value)
End Function

Private Sub BuildExecutableCodePositionMap( _
    ByVal text As String, _
    ByRef executableCodePositions() As Boolean)

    ReDim executableCodePositions(1 To Len(text))

    Dim scanPos As Long
    Dim ch As String
    Dim nextCh As String
    Dim quoteChar As String
    Dim inQuote As Boolean
    Dim inLineComment As Boolean
    Dim inBlockComment As Boolean

    scanPos = 1

    Do While scanPos <= Len(text)
        ch = Mid$(text, scanPos, 1)
        nextCh = vbNullString
        If scanPos < Len(text) Then nextCh = Mid$(text, scanPos + 1, 1)

        If inLineComment Then
            If ch = vbCr Or ch = vbLf Then inLineComment = False
            scanPos = scanPos + 1
        ElseIf inBlockComment Then
            If ch = "*" And nextCh = "/" Then
                inBlockComment = False
                scanPos = scanPos + 2
            Else
                scanPos = scanPos + 1
            End If
        ElseIf inQuote Then
            If ch = "\" Then
                scanPos = scanPos + 1
                If scanPos <= Len(text) Then scanPos = scanPos + 1
            ElseIf ch = quoteChar Then
                If nextCh = quoteChar Then
                    scanPos = scanPos + 2
                Else
                    inQuote = False
                    quoteChar = vbNullString
                    scanPos = scanPos + 1
                End If
            Else
                scanPos = scanPos + 1
            End If
        ElseIf ch = """" Or ch = "'" Then
            inQuote = True
            quoteChar = ch
            scanPos = scanPos + 1
        ElseIf ch = "/" And nextCh = "/" Then
            inLineComment = True
            scanPos = scanPos + 2
        ElseIf ch = "#" Then
            inLineComment = True
            scanPos = scanPos + 1
        ElseIf ch = "/" And nextCh = "*" Then
            inBlockComment = True
            scanPos = scanPos + 2
        Else
            executableCodePositions(scanPos) = True
            scanPos = scanPos + 1
        End If
    Loop
End Sub

Private Function IsFunctionPrefixCandidate( _
    ByVal text As String, _
    ByRef executableCodePositions() As Boolean, _
    ByVal prefixPos As Long) As Boolean

    If prefixPos < LBound(executableCodePositions) Then Exit Function
    If prefixPos > UBound(executableCodePositions) Then Exit Function
    If Not executableCodePositions(prefixPos) Then Exit Function

    If prefixPos > 1 Then
        Dim previousChar As String
        previousChar = Mid$(text, prefixPos - 1, 1)

        If IsIdentifierChar(previousChar) Or previousChar = "$" Then Exit Function
    End If

    IsFunctionPrefixCandidate = True
End Function

Private Function MarkAllOccurrencesForOnePrefixAcrossRows( _
    ByVal ws As Worksheet, _
    ByVal virtualText As String, _
    ByRef executableCodePositions() As Boolean, _
    ByRef rowTexts() As String, _
    ByRef rowStartPositions() As Long, _
    ByVal firstRow As Long, _
    ByVal prefix As String, _
    ByVal hitMessage As String) As Boolean

    Dim pattern As String
    pattern = prefix & "("

    Dim searchStartPos As Long
    searchStartPos = 1

    Dim hit As Boolean
    hit = False

    Do
        Dim prefixPos As Long
        prefixPos = InStr(searchStartPos, virtualText, pattern, vbTextCompare)
        If prefixPos = 0 Then Exit Do

        If IsFunctionPrefixCandidate(virtualText, executableCodePositions, prefixPos) Then
            Dim openParenPos As Long
            openParenPos = prefixPos + Len(prefix)

            Dim closePos As Long
            closePos = FindMatchingClosingParen(virtualText, executableCodePositions, openParenPos)

            If closePos > 0 Then
                Dim formatStartPos As Long
                formatStartPos = ResolveFormatStartPosition(virtualText, prefixPos)

                ApplyFormattedVirtualRange ws, rowTexts, rowStartPositions, firstRow, formatStartPos, closePos, hitMessage

                hit = True
                searchStartPos = closePos + 1
            Else
                searchStartPos = prefixPos + 1
            End If
        Else
            searchStartPos = prefixPos + 1
        End If
    Loop

    MarkAllOccurrencesForOnePrefixAcrossRows = hit
End Function

Private Sub ApplyFormattedVirtualRange( _
    ByVal ws As Worksheet, _
    ByRef rowTexts() As String, _
    ByRef rowStartPositions() As Long, _
    ByVal firstRow As Long, _
    ByVal formatStartPos As Long, _
    ByVal formatEndPos As Long, _
    ByVal hitMessage As String)

    Dim rowIndex As Long
    Dim rowNumber As Long
    Dim rowTextLength As Long
    Dim rowStartPos As Long
    Dim rowEndPos As Long
    Dim segmentStartPos As Long
    Dim segmentEndPos As Long
    Dim localStartPos As Long
    Dim segmentLength As Long

    For rowIndex = LBound(rowTexts) To UBound(rowTexts)
        rowTextLength = Len(rowTexts(rowIndex))
        If rowTextLength > 0 Then
            rowStartPos = rowStartPositions(rowIndex)
            rowEndPos = rowStartPos + rowTextLength - 1

            segmentStartPos = MaxLong(formatStartPos, rowStartPos)
            segmentEndPos = MinLong(formatEndPos, rowEndPos)

            If segmentStartPos <= segmentEndPos Then
                rowNumber = firstRow + rowIndex - 1
                localStartPos = segmentStartPos - rowStartPos + 1
                segmentLength = segmentEndPos - segmentStartPos + 1

                With ws.Cells(rowNumber, "B").Characters(localStartPos, segmentLength).Font
                    .Color = vbRed
                    .Bold = True
                End With

                With ws.Cells(rowNumber, "C")
                    .Value2 = hitMessage
                    .Font.Color = vbRed
                End With
            End If
        End If
    Next rowIndex
End Sub

Private Function MaxLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        MaxLong = leftValue
    Else
        MaxLong = rightValue
    End If
End Function

Private Function MinLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue <= rightValue Then
        MinLong = leftValue
    Else
        MinLong = rightValue
    End If
End Function

Private Function ShouldFillOnlyAValueRow(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim valueA As Variant
    Dim valueB As Variant

    valueA = ws.Cells(rowNumber, "A").Value2
    valueB = ws.Cells(rowNumber, "B").Value2

    ShouldFillOnlyAValueRow = HasCellValueForOnlyARowRule(valueA) And _
                              (Not HasCellValueForOnlyARowRule(valueB))
End Function

Private Sub ApplyOnlyAValueRowFill(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal fillColor As Long, ByVal fillTargetOption As String)
    Select Case fillTargetOption
        Case "LEFT"
            ws.Cells(rowNumber, "A").Interior.Color = fillColor
        Case "RIGHT"
            ws.Cells(rowNumber, "B").Interior.Color = fillColor
        Case "BOTH"
            ws.Cells(rowNumber, "A").Interior.Color = fillColor
            ws.Cells(rowNumber, "B").Interior.Color = fillColor
        Case Else
            ' NONE または不正値は塗りつぶししない
    End Select
End Sub

Private Function ResolveOnlyAValueRowFillTargetOption() As String
    Dim normalized As String

    normalized = UCase$(Trim$(ResolveOnlyAValueRowFillTargetRaw()))

    Select Case normalized
        Case "NONE", "LEFT", "RIGHT", "BOTH"
            ResolveOnlyAValueRowFillTargetOption = normalized
        Case Else
            ResolveOnlyAValueRowFillTargetOption = "BOTH"
    End Select
End Function

Private Function ResolveOnlyAValueRowFillColor() As Long
    ResolveOnlyAValueRowFillColor = HexColorTextToColorLongOrDefault(ResolveOnlyAValueRowFillColorHexRaw(), RGB(166, 166, 166))
End Function

Private Function HasCellValueForOnlyARowRule(ByVal value As Variant) As Boolean
    If IsError(value) Then
        HasCellValueForOnlyARowRule = True
        Exit Function
    End If

    If IsEmpty(value) Or IsNull(value) Then Exit Function

    If VarType(value) = vbString Then
        HasCellValueForOnlyARowRule = (Len(Trim$(CStr(value))) > 0)
    Else
        HasCellValueForOnlyARowRule = (Len(CStr(value)) > 0)
    End If
End Function

Private Function HexColorTextToColorLongOrDefault(ByVal rawHex As String, ByVal defaultColor As Long) As Long
    Dim t As String
    Dim redPart As Long
    Dim greenPart As Long
    Dim bluePart As Long
    Dim i As Long
    Dim ch As String

    t = UCase$(Trim$(rawHex))
    If Len(t) = 0 Then
        HexColorTextToColorLongOrDefault = defaultColor
        Exit Function
    End If

    If Left$(t, 1) = "#" Then
        t = Mid$(t, 2)
    ElseIf Left$(t, 2) = "0X" Then
        t = Mid$(t, 3)
    End If

    If Len(t) <> 6 Then
        HexColorTextToColorLongOrDefault = defaultColor
        Exit Function
    End If

    For i = 1 To 6
        ch = Mid$(t, i, 1)
        If InStr(1, "0123456789ABCDEF", ch, vbBinaryCompare) = 0 Then
            HexColorTextToColorLongOrDefault = defaultColor
            Exit Function
        End If
    Next i

    On Error GoTo ParseError

    redPart = CLng("&H" & Mid$(t, 1, 2))
    greenPart = CLng("&H" & Mid$(t, 3, 2))
    bluePart = CLng("&H" & Mid$(t, 5, 2))

    HexColorTextToColorLongOrDefault = RGB(redPart, greenPart, bluePart)
    Exit Function

ParseError:
    HexColorTextToColorLongOrDefault = defaultColor
End Function


'============================================================
' 開き括弧に対応する閉じ括弧を探す
' - sqlS(xxx + trim(yyy) + "zzz") のようなネスト括弧に対応する
' - 文字列やコメント内の括弧はコード位置マップで無視する
'============================================================
Private Function FindMatchingClosingParen( _
    ByVal text As String, _
    ByRef executableCodePositions() As Boolean, _
    ByVal openParenPos As Long) As Long

    If openParenPos < 1 Or openParenPos > Len(text) Then Exit Function
    If Mid$(text, openParenPos, 1) <> "(" Then Exit Function
    If Not executableCodePositions(openParenPos) Then Exit Function

    Dim depth As Long
    Dim scanPos As Long
    Dim ch As String

    depth = 0

    For scanPos = openParenPos To Len(text)
        If executableCodePositions(scanPos) Then
            ch = Mid$(text, scanPos, 1)

            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                depth = depth - 1
                If depth = 0 Then
                    FindMatchingClosingParen = scanPos
                    Exit Function
                End If
            End If
        End If
    Next scanPos
End Function

'============================================================
' 装飾開始位置を求める
'
' 例1:
'   sqlS('abc')
'   → 開始位置は "s"
'
' 例2:
'   DbHelper.sqlS('abc')
'   → "." の左側の識別子 "DbHelper" も含めて装飾するため
'     開始位置は "D"
'
' この関数では、
' - prefix の直前が "." かどうかを見る
' - "." の左側にある識別子 [A-Za-z0-9_] を逆向きにたどる
'============================================================
Private Function ResolveFormatStartPosition(ByVal text As String, ByVal prefixPos As Long) As Long
    ResolveFormatStartPosition = prefixPos

    ' prefix の直前に "." が無ければ、通常の関数呼び出しとしてそのまま返す
    If prefixPos <= 1 Then Exit Function
    If Mid$(text, prefixPos - 1, 1) <> "." Then Exit Function

    ' "." の左側にある識別子を含める
    Dim scanPos As Long
    scanPos = prefixPos - 2   ' "." の1文字左から調べ始める

    If scanPos < 1 Then Exit Function

    Do While scanPos >= 1
        If IsIdentifierChar(Mid$(text, scanPos, 1)) Then
            scanPos = scanPos - 1
        Else
            Exit Do
        End If
    Loop

    ' 識別子の先頭位置 = 条件を満たさなくなった位置の次
    ResolveFormatStartPosition = scanPos + 1
End Function

'============================================================
' 識別子に使える文字かどうかを判定する
' - 英字
' - 数字
' - アンダースコア
'
' 想定:
'   DbHelper
'   dbHelper
'   helper_01
'   mDb
'============================================================
Private Function IsIdentifierChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function

    Select Case AscW(ch)
        Case 48 To 57   ' 0-9
            IsIdentifierChar = True
        Case 65 To 90   ' A-Z
            IsIdentifierChar = True
        Case 95         ' _
            IsIdentifierChar = True
        Case 97 To 122  ' a-z
            IsIdentifierChar = True
    End Select
End Function

'============================================================
' モジュール先頭のCSV定数から prefix を読み込む
' - 追加したい場合は ESCAPE_TARGET_PREFIXES_CSV へ追記
'============================================================
Private Function LoadPrefixesFromCode() As Collection
    Dim prefixes As New Collection
    Dim rawPrefixes As String
    Dim prefixItems As Variant
    Dim i As Long
    Dim v As String

    rawPrefixes = Replace(ResolveEscapeTargetPrefixesCsvRaw(), "，", ",")
    prefixItems = Split(rawPrefixes, ",")

    For i = LBound(prefixItems) To UBound(prefixItems)
        v = Trim$(CStr(prefixItems(i)))
        If Len(v) > 0 Then
            prefixes.Add v
        End If
    Next i

    Set LoadPrefixesFromCode = prefixes
End Function

'============================================================
' ファイル選択ダイアログ（Excelファイル用）
'============================================================
Private Function ResolveCompletionMessage() As String
    If mUiOptions.Enabled And mUiOptions.UseCompletionMessage Then
        ResolveCompletionMessage = CStr(mUiOptions.completionMessage)
    Else
        ResolveCompletionMessage = DEFAULT_COMPLETION_MESSAGE
    End If
End Function

Private Function ResolveEscapeTargetPrefixesCsvRaw() As String
    If mUiOptions.Enabled And mUiOptions.UseEscapeTargetPrefixesCsv Then
        ResolveEscapeTargetPrefixesCsvRaw = CStr(mUiOptions.escapeTargetPrefixesCsv)
    Else
        ResolveEscapeTargetPrefixesCsvRaw = ESCAPE_TARGET_PREFIXES_CSV
    End If
End Function

Private Function ResolveOnlyAValueRowFillTargetRaw() As String
    If mUiOptions.Enabled And mUiOptions.UseOnlyAValueRowFillTarget Then
        ResolveOnlyAValueRowFillTargetRaw = CStr(mUiOptions.onlyAValueRowFillTarget)
    Else
        ResolveOnlyAValueRowFillTargetRaw = OPTION_ONLY_A_VALUE_ROW_FILL_TARGET
    End If
End Function

Private Function ResolveOnlyAValueRowFillColorHexRaw() As String
    If mUiOptions.Enabled And mUiOptions.UseOnlyAValueRowFillColorHex Then
        ResolveOnlyAValueRowFillColorHexRaw = CStr(mUiOptions.onlyAValueRowFillColorHex)
    Else
        ResolveOnlyAValueRowFillColorHexRaw = ONLY_A_VALUE_ROW_FILL_COLOR_HEX
    End If
End Function

Private Function PickExcelFilePath() As String
    Dim fd As Object

    If mUiOptions.Enabled Then
        PickExcelFilePath = Trim$(mUiOptions.TargetWorkbookPath)
        Exit Function
    End If

    Set fd = Application.FileDialog(3) ' 3 = msoFileDialogFilePicker

    With fd
        .Title = "加工対象の Excel ファイルを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm;*.xlsb"

        If .Show <> -1 Then
            PickExcelFilePath = ""
            Exit Function
        End If

        PickExcelFilePath = .SelectedItems(1)
    End With
End Function















