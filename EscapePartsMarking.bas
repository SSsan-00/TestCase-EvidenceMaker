Attribute VB_Name = "EscapePartsMarking"
Option Explicit

Private Const DEFAULT_COMPLETION_MESSAGE As String = "SQLインジェクション対策済み"
Private Const ESCAPE_TARGET_PREFIXES_CSV As String = "sqlS,sqlN"
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
' - B列の "sqlX(...)" 部分だけを赤字＋太字（複数ヒット対応）
' - ただし "DbHelper.sqlX(...)" のようなクラス/インスタンス経由の呼び出しは
'   "DbHelper." も含めて赤字＋太字にする
' - ヒットした行の C列に固定メッセージ（既定: "SQLインジェクション対策済み"）を赤字で書く
'
' 前提:
'  - モジュール先頭の ESCAPE_TARGET_PREFIXES_CSV にエスケープ関数（例: sqlS, sqlN）を列挙していること
'============================================================
Public Sub RunMain()
    Dim targetPath As String
    targetPath = PickExcelFilePath()
    If targetPath = "" Then Exit Sub ' キャンセル

    Dim prefixes As Collection
    Set prefixes = LoadPrefixesFromCode()
    If prefixes.Count = 0 Then
        MsgBox "ESCAPE_TARGET_PREFIXES_CSV に prefix（例: sqlS, sqlN）を1つ以上設定してください。", vbExclamation
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
' - B列を走査して sqlX(...) / DbHelper.sqlX(...) を装飾
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

        Dim cell As Range
        Set cell = ws.Cells(r, "B")

        If Len(cell.Value2) > 0 Then
            Dim hit As Boolean
            hit = MarkSqlPartsInCell(cell, prefixes)

            If hit Then
                Dim eCell As Range
                Set eCell = ws.Cells(r, "C")
                eCell.Value2 = hitMessage
                eCell.Font.Color = vbRed
            End If
        End If
    Next r
End Sub

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
' セル内の複数パターンをすべて装飾する
' - prefix + "(" の開始位置を探す
' - 直前が "." の場合は、左側の識別子（例: DbHelper）も含める
' - そこから次の ")" までを赤字＋太字
' - 同一セル内に複数存在してもすべて処理
'
' 戻り値:
'   True  = 1つ以上ヒットして装飾した
'   False = ヒットなし
'============================================================
Private Function MarkSqlPartsInCell(ByVal cell As Range, ByVal prefixes As Collection) As Boolean
    Dim text As String
    Dim anyHit As Boolean
    Dim i As Long
    Dim prefix As String
    Dim prefixCount As Long

    text = CStr(cell.Value2)

    ' 開き括弧が無い文字列は prefix(...) パターンを含まない
    If InStr(1, text, "(", vbBinaryCompare) = 0 Then Exit Function

    anyHit = False
    prefixCount = prefixes.Count

    For i = 1 To prefixCount
        prefix = CStr(prefixes(i))
        anyHit = MarkAllOccurrencesForOnePrefix(cell, text, prefix) Or anyHit
    Next i

    MarkSqlPartsInCell = anyHit
End Function

'============================================================
' 1つの prefix について、セル内の全出現箇所を装飾する
' - 例: prefix="sqlS" なら "sqlS(" をすべて探す
' - 見つけたら直後の ")" を探す
' - 直前が "." の場合は、左側の識別子も装飾範囲に含める
'   例: "DbHelper.sqlS(...)" → "DbHelper.sqlS(...)" 全体を装飾
' - 同じセル内に複数あっても全部処理する
'============================================================
Private Function MarkAllOccurrencesForOnePrefix(ByVal cell As Range, ByVal text As String, ByVal prefix As String) As Boolean
    Dim pattern As String
    pattern = prefix & "("

    Dim searchStartPos As Long
    searchStartPos = 1

    Dim hit As Boolean
    hit = False

    Do
        Dim prefixPos As Long
        prefixPos = InStr(searchStartPos, text, pattern, vbTextCompare)
        If prefixPos = 0 Then Exit Do

        Dim closePos As Long
        closePos = InStr(prefixPos + Len(pattern), text, ")", vbTextCompare)

        If closePos > 0 Then
            '----------------------------------------------------
            ' 装飾開始位置を決める
            '
            ' 通常:
            '   sqlS(...)
            '   ↑ ここから装飾
            '
            ' クラス/インスタンス経由:
            '   DbHelper.sqlS(...)
            '   ↑ ここから装飾
            '----------------------------------------------------
            Dim formatStartPos As Long
            formatStartPos = ResolveFormatStartPosition(text, prefixPos)

            Dim lengthToFormat As Long
            lengthToFormat = (closePos - formatStartPos) + 1

            With cell.Characters(formatStartPos, lengthToFormat).Font
                .Color = vbRed
                .Bold = True
            End With

            hit = True
            searchStartPos = closePos + 1
        Else
            ' 閉じ括弧が無い異常/未完成パターンは、少し進めて次を探す
            searchStartPos = prefixPos + 1
        End If
    Loop

    MarkAllOccurrencesForOnePrefix = hit
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















