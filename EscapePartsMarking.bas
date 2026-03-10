Attribute VB_Name = "EscapePartsMarking"
Option Explicit

Private Const DEFAULT_COMPLETION_MESSAGE As String = "SQLインジェクション対策済み"  ' ここを書き換えるとC列メッセージを変更できます
Public Const OPTION_FILL_ONLY_A_VALUE_ROW_ENABLED As Boolean = True ' True: A列のみ入力行のA/Bを塗りつぶす / False: 塗りつぶししない
Private Const ONLY_A_VALUE_ROW_FILL_COLOR_HEX As String = "#a6a6a6" ' A列のみ入力行の塗りつぶし色（#RRGGBB）

'============================================================
' xlsmツール（別ファイル）から、選択した xlsx を開いて加工するマクロ
' - シート名に "A1-1-1" を含むシートのみを対象に処理する
' - A列のみ入力（B列空）の行は A/B を指定色で塗りつぶす（オプション）
' - B列の "sqlX(...)" 部分だけを赤字＋太字（複数ヒット対応）
' - ヒットした行の C列に固定メッセージ（既定: "SQLインジェクション対策済み"）を赤字で書く
'
' 前提:
'  - このモジュール内の LoadPrefixesFromCode へ エスケープ関数（例: sqlS, sqlN）を列挙していること
'============================================================
Public Sub RunMain()
    Dim targetPath As String
    targetPath = PickExcelFilePath()
    If targetPath = "" Then Exit Sub ' キャンセル

    Dim prefixes As Collection
    Set prefixes = LoadPrefixesFromCode()
    If prefixes.Count = 0 Then
        MsgBox "LoadPrefixesFromCode に prefix（例: sqlS, sqlN）を1つ以上設定してください。", vbExclamation
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
            ProcessOneSheet ws, prefixes, DEFAULT_COMPLETION_MESSAGE
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
' - A列のみ入力（B列空）の行は A/B を指定色で塗りつぶす（オプション）
' - B列を走査して sqlX(...) を装飾
' - ヒット行のC列に固定メッセージ＆赤字
'============================================================
Private Sub ProcessOneSheet(ByVal ws As Worksheet, ByVal prefixes As Collection, ByVal hitMessage As String)
    ' A/B列の最終行を取得（どちらにもデータが無ければスキップ）
    Dim lastRowA As Long
    Dim lastRowB As Long
    Dim lastRow As Long
    Dim onlyAFillColor As Long

    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    lastRow = lastRowA
    If lastRowB > lastRow Then lastRow = lastRowB
    If lastRow < 4 Then Exit Sub

    onlyAFillColor = ResolveOnlyAValueRowFillColor()

    Dim r As Long
    For r = 4 To lastRow
        If ShouldFillOnlyAValueRow(ws, r) Then
            ApplyOnlyAValueRowFill ws, r, onlyAFillColor
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

    If Not OPTION_FILL_ONLY_A_VALUE_ROW_ENABLED Then Exit Function

    valueA = ws.Cells(rowNumber, "A").Value2
    valueB = ws.Cells(rowNumber, "B").Value2

    ShouldFillOnlyAValueRow = HasCellValueForOnlyARowRule(valueA) And _
                              (Not HasCellValueForOnlyARowRule(valueB))
End Function

Private Sub ApplyOnlyAValueRowFill(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal fillColor As Long)
    ws.Cells(rowNumber, "A").Interior.Color = fillColor
    ws.Cells(rowNumber, "B").Interior.Color = fillColor
End Sub

Private Function ResolveOnlyAValueRowFillColor() As Long
    ResolveOnlyAValueRowFillColor = HexColorTextToColorLongOrDefault(ONLY_A_VALUE_ROW_FILL_COLOR_HEX, RGB(166, 166, 166))
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
' - 見つけたら 直後の ")" を探して、その範囲を装飾する
' - 同じセル内に複数あっても全部処理する
'============================================================
Private Function MarkAllOccurrencesForOnePrefix(ByVal cell As Range, ByVal text As String, ByVal prefix As String) As Boolean
    Dim pattern As String
    pattern = prefix & "("

    Dim startPos As Long
    startPos = 1

    Dim hit As Boolean
    hit = False

    Do
        Dim openPos As Long
        openPos = InStr(startPos, text, pattern, vbTextCompare)
        If openPos = 0 Then Exit Do

        Dim closePos As Long
        closePos = InStr(openPos + Len(pattern), text, ")", vbTextCompare)

        If closePos > 0 Then
            Dim lengthToFormat As Long
            lengthToFormat = (closePos - openPos) + 1

            With cell.Characters(openPos, lengthToFormat).Font
                .Color = vbRed
                .Bold = True
            End With

            hit = True
            startPos = closePos + 1
        Else
            startPos = openPos + 1
        End If
    Loop

    MarkAllOccurrencesForOnePrefix = hit
End Function

'============================================================
' ソースコード内のリストから prefix を読み込む
' - 追加したい場合は defaultPrefixes の配列へ追記
'============================================================
Private Function LoadPrefixesFromCode() As Collection
    Dim prefixes As New Collection
    Dim defaultPrefixes As Variant
    Dim i As Long
    Dim v As String

    ' ここにエスケープ関数名を追加する
    defaultPrefixes = Array("sqlS", "sqlN")

    For i = LBound(defaultPrefixes) To UBound(defaultPrefixes)
        v = Trim$(CStr(defaultPrefixes(i)))
        If v <> "" Then
            prefixes.Add v
        End If
    Next i

    Set LoadPrefixesFromCode = prefixes
End Function

'============================================================
' ファイル選択ダイアログ（Excelファイル用）
'============================================================
Private Function PickExcelFilePath() As String
    Dim fd As Object
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