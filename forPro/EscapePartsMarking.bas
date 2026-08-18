Attribute VB_Name = "EscapePartsMarking"
Option Explicit

Private Const DEFAULT_COMPLETION_MESSAGE As String = "SQLインジェクション対策済み"
Private Const ESCAPE_TARGET_PREFIXES_CSV As String = "pg_escape_string,sqlS,sqlN,sqlLS,sqlC,sqlNZ,sqlInN,sqlF,sqlChk,sqlLikeStr,sqlNum,sqlNum0,sqlStr"
Private Const OPTION_ONLY_A_VALUE_ROW_FILL_TARGET As String = "Both" ' None / Left / Right / Both
Private Const ONLY_A_VALUE_ROW_FILL_COLOR_HEX As String = "#a6a6a6"
Private Const FIRST_SCAN_ROW As Long = 4
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

' フォームまたはCONFIGから渡された設定を一時適用し、通常の実行経路でエスケープ箇所を装飾する。
Public Sub RunMainWithUiOptions(ByRef options As EscapePartsMarkingUiOptions)
    ClearUiOptions
    mUiOptions = options
    mUiOptions.Enabled = True

    RunMain

    ClearUiOptions
End Sub

' フォームとCONFIGで共有するエスケープ箇所マーキングの既定設定を作成する。
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

' 前回のUI設定が単体実行へ漏れないよう、モジュール保持値を初期化する。
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
' 対象ブック選択、対象シート走査、装飾、保存までの処理全体を統括する。
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
' 一つのA1-1-1シートについて、対象関数装飾とA列のみ行の塗りつぶしを適用する。
Private Sub ProcessOneSheet(ByVal ws As Worksheet, ByVal prefixes As Collection, ByVal hitMessage As String)
    ' A/B列の最終行を取得（どちらにもデータが無ければスキップ）
    Dim lastRowA As Long
    Dim lastRowB As Long
    Dim lastRow As Long
    Dim onlyAFillColor As Long
    Dim fillTargetOption As String
    Dim sourceValues As Variant

    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    lastRow = lastRowA
    If lastRowB > lastRow Then lastRow = lastRowB
    If lastRow < FIRST_SCAN_ROW Then Exit Sub

    fillTargetOption = ResolveOnlyAValueRowFillTargetOption()
    sourceValues = ws.Range(ws.Cells(FIRST_SCAN_ROW, "A"), ws.Cells(lastRow, "B")).Value2

    If fillTargetOption <> "NONE" Then
        onlyAFillColor = ResolveOnlyAValueRowFillColor()
        Dim rowIndex As Long
        For rowIndex = 1 To UBound(sourceValues, 1)
            If ShouldFillOnlyAValueRowValues(sourceValues(rowIndex, 1), sourceValues(rowIndex, 2)) Then
                ApplyOnlyAValueRowFill ws, FIRST_SCAN_ROW + rowIndex - 1, onlyAFillColor, fillTargetOption
            End If
        Next rowIndex
    End If

    MarkSqlPartsInRows ws, prefixes, hitMessage, FIRST_SCAN_ROW, sourceValues
End Sub

' B列の複数行を仮想ソースとして解析し、対象関数呼び出しへ書式を設定する。
Private Sub MarkSqlPartsInRows( _
    ByVal ws As Worksheet, _
    ByVal prefixes As Collection, _
    ByVal hitMessage As String, _
    ByVal firstRow As Long, _
    ByRef sourceValues As Variant)

    Dim rowCount As Long
    rowCount = UBound(sourceValues, 1)
    If rowCount <= 0 Then Exit Sub

    Dim rowTexts() As String
    Dim rowStartPositions() As Long
    Dim virtualLines() As String
    ReDim rowTexts(1 To rowCount)
    ReDim rowStartPositions(1 To rowCount)
    ReDim virtualLines(1 To rowCount)

    Dim virtualText As String
    Dim rowIndex As Long
    Dim nextStartPosition As Long

    nextStartPosition = 1
    For rowIndex = 1 To rowCount
        rowTexts(rowIndex) = GetMarkingCellText(sourceValues(rowIndex, 2))
        virtualLines(rowIndex) = rowTexts(rowIndex)
        rowStartPositions(rowIndex) = nextStartPosition
        nextStartPosition = nextStartPosition + Len(rowTexts(rowIndex))
        If rowIndex < rowCount Then nextStartPosition = nextStartPosition + Len(vbLf)
    Next rowIndex
    virtualText = Join(virtualLines, vbLf)

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

' エラー値を除外し、装飾解析に使えるB列セル文字列を返す。
Private Function GetMarkingCellText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function

    GetMarkingCellText = CStr(value)
End Function

' C#の文字列・コメントを除外し、実行コードとして検索可能な文字位置だけを記録する。
Private Sub BuildExecutableCodePositionMap( _
    ByVal text As String, _
    ByRef executableCodePositions() As Boolean)

    Dim textLength As Long
    Dim scanPos As Long

    textLength = Len(text)
    ReDim executableCodePositions(1 To textLength)

    scanPos = 1
    ScanCSharpCode text, scanPos, textLength, executableCodePositions, 0
End Sub

' 通常コードを走査し、コメント、通常文字列、補間文字列、raw文字列をそれぞれの規則で読み飛ばす。
Private Sub ScanCSharpCode( _
    ByVal text As String, _
    ByRef scanPos As Long, _
    ByVal textLength As Long, _
    ByRef executableCodePositions() As Boolean, _
    ByVal interpolationEndBraceCount As Long)

    Dim ch As String
    Dim nextCh As String
    Dim closingRun As Long
    Dim codeBraceDepth As Long
    Dim parenthesisDepth As Long
    Dim bracketDepth As Long

    Do While scanPos <= textLength
        ch = Mid$(text, scanPos, 1)
        nextCh = vbNullString
        If scanPos < textLength Then nextCh = Mid$(text, scanPos + 1, 1)

        If interpolationEndBraceCount > 0 And ch = "}" Then
            closingRun = CountConsecutiveCharacter(text, scanPos, "}")

            Do While closingRun > 0 And codeBraceDepth > 0
                executableCodePositions(scanPos) = True
                scanPos = scanPos + 1
                closingRun = closingRun - 1
                codeBraceDepth = codeBraceDepth - 1
            Loop

            If codeBraceDepth = 0 And closingRun >= interpolationEndBraceCount Then
                Exit Sub
            End If

            Do While closingRun > 0
                executableCodePositions(scanPos) = True
                scanPos = scanPos + 1
                closingRun = closingRun - 1
            Loop
        ElseIf TrySkipCSharpStringLiteral(text, scanPos, textLength, executableCodePositions) Then
            ' 文字列内は、補間式として再帰解析されたコード部分だけがTrueになる
        ElseIf ch = "/" And nextCh = "/" Then
            SkipCSharpLineComment text, scanPos, textLength
        ElseIf ch = "#" Then
            ' 既存仕様との互換性のため、#以降も行コメントとして扱う
            SkipCSharpLineComment text, scanPos, textLength
        ElseIf ch = "/" And nextCh = "*" Then
            SkipCSharpBlockComment text, scanPos, textLength
        ElseIf IsInterpolationFormatSeparator( _
            text, scanPos, textLength, interpolationEndBraceCount, _
            codeBraceDepth, parenthesisDepth, bracketDepth) Then

            scanPos = scanPos + 1
            SkipInterpolationFormatText text, scanPos, textLength, interpolationEndBraceCount
            Exit Sub
        Else
            executableCodePositions(scanPos) = True

            Select Case ch
                Case "{"
                    codeBraceDepth = codeBraceDepth + 1
                Case "("
                    parenthesisDepth = parenthesisDepth + 1
                Case ")"
                    If parenthesisDepth > 0 Then parenthesisDepth = parenthesisDepth - 1
                Case "["
                    bracketDepth = bracketDepth + 1
                Case "]"
                    If bracketDepth > 0 Then bracketDepth = bracketDepth - 1
            End Select

            scanPos = scanPos + 1
        End If
    Loop
End Sub

' 現在位置がC#文字列開始なら種類を判別し、対応する終端まで走査位置を進める。
Private Function TrySkipCSharpStringLiteral( _
    ByVal text As String, _
    ByRef scanPos As Long, _
    ByVal textLength As Long, _
    ByRef executableCodePositions() As Boolean) As Boolean

    Dim ch As String
    Dim dollarCount As Long
    Dim quoteCount As Long
    Dim quotePos As Long

    ch = Mid$(text, scanPos, 1)

    If ch = "$" Then
        dollarCount = CountConsecutiveCharacter(text, scanPos, "$")
        quotePos = scanPos + dollarCount
        quoteCount = CountConsecutiveCharacter(text, quotePos, """")

        If quoteCount >= 3 Then
            ScanRawString text, scanPos, textLength, executableCodePositions, dollarCount, quoteCount
            TrySkipCSharpStringLiteral = True
            Exit Function
        End If

        If dollarCount = 1 Then
            If quoteCount = 1 Then
                ScanInterpolatedQuotedString text, scanPos, textLength, executableCodePositions, quotePos, False
                TrySkipCSharpStringLiteral = True
                Exit Function
            End If

            If quotePos <= textLength Then
                If Mid$(text, quotePos, 1) = "@" Then
                    quotePos = quotePos + 1
                    If quotePos <= textLength Then
                        If Mid$(text, quotePos, 1) = """" Then
                            ScanInterpolatedQuotedString text, scanPos, textLength, executableCodePositions, quotePos, True
                            TrySkipCSharpStringLiteral = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    ElseIf ch = "@" Then
        quotePos = scanPos + 1

        If quotePos <= textLength Then
            If Mid$(text, quotePos, 1) = "$" Then quotePos = quotePos + 1

            If quotePos <= textLength Then
                If Mid$(text, quotePos, 1) = """" Then
                    If quotePos = scanPos + 2 Then
                        ScanInterpolatedQuotedString text, scanPos, textLength, executableCodePositions, quotePos, True
                    Else
                        SkipQuotedString text, scanPos, textLength, quotePos, """", True
                    End If

                    TrySkipCSharpStringLiteral = True
                    Exit Function
                End If
            End If
        End If
    ElseIf ch = """" Then
        quoteCount = CountConsecutiveCharacter(text, scanPos, """")

        If quoteCount >= 3 Then
            ScanRawString text, scanPos, textLength, executableCodePositions, 0, quoteCount
        Else
            SkipQuotedString text, scanPos, textLength, scanPos, """", False
        End If

        TrySkipCSharpStringLiteral = True
        Exit Function
    ElseIf ch = "'" Then
        SkipQuotedString text, scanPos, textLength, scanPos, "'", False
        TrySkipCSharpStringLiteral = True
        Exit Function
    End If
End Function

' 補間文字列の本文を除外しつつ、波括弧内の式だけを再帰的なコードとして解析する。
Private Sub ScanInterpolatedQuotedString( _
    ByVal text As String, _
    ByRef scanPos As Long, _
    ByVal textLength As Long, _
    ByRef executableCodePositions() As Boolean, _
    ByVal quotePos As Long, _
    ByVal isVerbatim As Boolean)

    Dim ch As String
    Dim nextCh As String
    Dim braceRun As Long

    scanPos = quotePos + 1

    Do While scanPos <= textLength
        ch = Mid$(text, scanPos, 1)
        nextCh = vbNullString
        If scanPos < textLength Then nextCh = Mid$(text, scanPos + 1, 1)

        If ch = """" Then
            If isVerbatim And nextCh = """" Then
                scanPos = scanPos + 2
            Else
                scanPos = scanPos + 1
                Exit Sub
            End If
        ElseIf (Not isVerbatim) And ch = "\" Then
            scanPos = scanPos + 1
            If scanPos <= textLength Then scanPos = scanPos + 1
        ElseIf ch = "{" Then
            braceRun = CountConsecutiveCharacter(text, scanPos, "{")

            If (braceRun Mod 2) = 1 Then
                scanPos = scanPos + braceRun
                ScanCSharpCode text, scanPos, textLength, executableCodePositions, 1

                If scanPos <= textLength Then
                    If Mid$(text, scanPos, 1) = "}" Then
                        scanPos = scanPos + CountConsecutiveCharacter(text, scanPos, "}")
                    End If
                End If
            Else
                scanPos = scanPos + braceRun
            End If
        ElseIf ch = "}" Then
            scanPos = scanPos + CountConsecutiveCharacter(text, scanPos, "}")
        Else
            scanPos = scanPos + 1
        End If
    Loop
End Sub

' 引用符数とドル数に従ってraw文字列を走査し、有効な補間式だけをコードとして扱う。
Private Sub ScanRawString( _
    ByVal text As String, _
    ByRef scanPos As Long, _
    ByVal textLength As Long, _
    ByRef executableCodePositions() As Boolean, _
    ByVal dollarCount As Long, _
    ByVal quoteCount As Long)

    Dim quoteStart As Long
    Dim currentQuoteRun As Long
    Dim braceRun As Long

    quoteStart = scanPos + dollarCount
    scanPos = quoteStart + quoteCount

    Do While scanPos <= textLength
        If Mid$(text, scanPos, 1) = """" Then
            currentQuoteRun = CountConsecutiveCharacter(text, scanPos, """")

            If currentQuoteRun >= quoteCount Then
                scanPos = scanPos + currentQuoteRun
                Exit Sub
            End If

            scanPos = scanPos + currentQuoteRun
        ElseIf dollarCount > 0 And Mid$(text, scanPos, 1) = "{" Then
            braceRun = CountConsecutiveCharacter(text, scanPos, "{")

            If braceRun >= dollarCount And braceRun < (dollarCount * 2) Then
                scanPos = scanPos + braceRun
                ScanCSharpCode text, scanPos, textLength, executableCodePositions, dollarCount

                If scanPos <= textLength Then
                    If Mid$(text, scanPos, 1) = "}" Then
                        scanPos = scanPos + CountConsecutiveCharacter(text, scanPos, "}")
                    End If
                End If
            Else
                scanPos = scanPos + braceRun
            End If
        Else
            scanPos = scanPos + 1
        End If
    Loop
End Sub

' エスケープと逐語文字列の二重引用符を考慮し、通常文字列または文字リテラルを読み飛ばす。
Private Sub SkipQuotedString( _
    ByVal text As String, _
    ByRef scanPos As Long, _
    ByVal textLength As Long, _
    ByVal quotePos As Long, _
    ByVal quoteChar As String, _
    ByVal isVerbatim As Boolean)

    Dim ch As String
    Dim nextCh As String

    scanPos = quotePos + 1

    Do While scanPos <= textLength
        ch = Mid$(text, scanPos, 1)
        nextCh = vbNullString
        If scanPos < textLength Then nextCh = Mid$(text, scanPos + 1, 1)

        If isVerbatim Then
            If ch = quoteChar And nextCh = quoteChar Then
                scanPos = scanPos + 2
            ElseIf ch = quoteChar Then
                scanPos = scanPos + 1
                Exit Sub
            Else
                scanPos = scanPos + 1
            End If
        ElseIf ch = "\" Then
            scanPos = scanPos + 1
            If scanPos <= textLength Then scanPos = scanPos + 1
        ElseIf ch = quoteChar Then
            scanPos = scanPos + 1
            Exit Sub
        Else
            scanPos = scanPos + 1
        End If
    Loop
End Sub

' 改行直前までをコード対象外として、C#の行コメントを読み飛ばす。
Private Sub SkipCSharpLineComment(ByVal text As String, ByRef scanPos As Long, ByVal textLength As Long)
    Do While scanPos <= textLength
        If Mid$(text, scanPos, 1) = vbCr Then
            scanPos = scanPos + 1
            If scanPos <= textLength Then
                If Mid$(text, scanPos, 1) = vbLf Then scanPos = scanPos + 1
            End If
            Exit Sub
        ElseIf Mid$(text, scanPos, 1) = vbLf Then
            scanPos = scanPos + 1
            Exit Sub
        Else
            scanPos = scanPos + 1
        End If
    Loop
End Sub

' 閉じ記号またはソース末尾までをコード対象外として、ブロックコメントを読み飛ばす。
Private Sub SkipCSharpBlockComment(ByVal text As String, ByRef scanPos As Long, ByVal textLength As Long)
    scanPos = scanPos + 2

    Do While scanPos <= textLength
        If scanPos < textLength Then
            If Mid$(text, scanPos, 2) = "*/" Then
                scanPos = scanPos + 2
                Exit Sub
            End If
        End If

        scanPos = scanPos + 1
    Loop
End Sub

' 補間式のトップレベルにあるコロンだけを、書式指定開始として判定する。
Private Function IsInterpolationFormatSeparator( _
    ByVal text As String, _
    ByVal scanPos As Long, _
    ByVal textLength As Long, _
    ByVal interpolationEndBraceCount As Long, _
    ByVal codeBraceDepth As Long, _
    ByVal parenthesisDepth As Long, _
    ByVal bracketDepth As Long) As Boolean

    If interpolationEndBraceCount <= 0 Then Exit Function
    If Mid$(text, scanPos, 1) <> ":" Then Exit Function
    If codeBraceDepth > 0 Or parenthesisDepth > 0 Or bracketDepth > 0 Then Exit Function

    If scanPos > 1 Then
        If Mid$(text, scanPos - 1, 1) = ":" Then Exit Function
    End If

    If scanPos < textLength Then
        If Mid$(text, scanPos + 1, 1) = ":" Then Exit Function
    End If

    IsInterpolationFormatSeparator = True
End Function

' 補間式の書式指定部分を除外し、対応する閉じ波括弧まで走査位置を進める。
Private Sub SkipInterpolationFormatText( _
    ByVal text As String, _
    ByRef scanPos As Long, _
    ByVal textLength As Long, _
    ByVal interpolationEndBraceCount As Long)

    Do While scanPos <= textLength
        If Mid$(text, scanPos, 1) = "}" Then
            If CountConsecutiveCharacter(text, scanPos, "}") >= interpolationEndBraceCount Then
                Exit Sub
            End If
        End If

        scanPos = scanPos + 1
    Loop
End Sub

' raw文字列の区切り判定に使うため、同一文字が連続する長さを数える。
Private Function CountConsecutiveCharacter( _
    ByVal text As String, _
    ByVal startPos As Long, _
    ByVal targetChar As String) As Long

    Dim scanPos As Long

    If Len(targetChar) <> 1 Then Exit Function

    scanPos = startPos
    Do While scanPos <= Len(text)
        If Mid$(text, scanPos, 1) <> targetChar Then Exit Do
        CountConsecutiveCharacter = CountConsecutiveCharacter + 1
        scanPos = scanPos + 1
    Loop
End Function

' 対象関数名が別識別子の一部でなく、直後に呼出し括弧があるか判定する。
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

' 複数行にまたがる引数も含め、一つの対象関数名の全呼出し範囲を装飾する。
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

' 仮想ソース上の範囲を各B列セルへ分割し、文字単位の書式として適用する。
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

' 範囲端計算で使う二つのLong値の大きい方を返す。
Private Function MaxLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        MaxLong = leftValue
    Else
        MaxLong = rightValue
    End If
End Function

' 範囲端計算で使う二つのLong値の小さい方を返す。
Private Function MinLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue <= rightValue Then
        MinLong = leftValue
    Else
        MinLong = rightValue
    End If
End Function

' A列だけに値があるという塗りつぶし対象条件を、配列値から判定する。
Private Function ShouldFillOnlyAValueRowValues(ByVal valueA As Variant, ByVal valueB As Variant) As Boolean
    ShouldFillOnlyAValueRowValues = HasCellValueForOnlyARowRule(valueA) And _
                                   (Not HasCellValueForOnlyARowRule(valueB))
End Function

' 設定されたNone・Left・Right・Bothに従い、A列だけの行へ塗りつぶしを適用する。
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

' UI値またはソース定数を正規化し、塗りつぶす列の選択肢を確定する。
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

' UI値またはソース定数のカラーコードから、実際に適用するExcel色値を決定する。
Private Function ResolveOnlyAValueRowFillColor() As Long
    ResolveOnlyAValueRowFillColor = HexColorTextToColorLongOrDefault(ResolveOnlyAValueRowFillColorHexRaw(), RGB(166, 166, 166))
End Function

' Empty・空文字・空白だけを未入力とし、A列のみ行判定用の実値有無を返す。
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

' カラーコードをExcel色値へ変換し、不正値なら既定色を返す。
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
' 文字列とネスト括弧を考慮し、対象関数呼び出しに対応する閉じ括弧を検索する。
Private Function FindMatchingClosingParen( _
    ByVal text As String, _
    ByRef executableCodePositions() As Boolean, _
    ByVal openParenPos As Long) As Long

    Dim textLength As Long
    textLength = Len(text)

    If openParenPos < 1 Or openParenPos > textLength Then Exit Function
    If Mid$(text, openParenPos, 1) <> "(" Then Exit Function
    If Not executableCodePositions(openParenPos) Then Exit Function

    Dim depth As Long
    Dim scanPos As Long
    Dim ch As String

    depth = 0

    For scanPos = openParenPos To textLength
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
' 関数呼び出し後の添字記法を考慮し、装飾を開始する文字位置を決める。
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
' 英数字とアンダースコアをC#識別子構成文字として判定する。
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
' 編集しやすいカンマ区切り設定を、空欄と重複を除いた対象関数一覧へ変換する。
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
' UI指定があればそれを使い、なければソース定数の完了メッセージを返す。
Private Function ResolveCompletionMessage() As String
    If mUiOptions.Enabled And mUiOptions.UseCompletionMessage Then
        ResolveCompletionMessage = CStr(mUiOptions.completionMessage)
    Else
        ResolveCompletionMessage = DEFAULT_COMPLETION_MESSAGE
    End If
End Function

' UI指定があればそれを使い、なければソース定数の対象関数一覧を返す。
Private Function ResolveEscapeTargetPrefixesCsvRaw() As String
    If mUiOptions.Enabled And mUiOptions.UseEscapeTargetPrefixesCsv Then
        ResolveEscapeTargetPrefixesCsvRaw = CStr(mUiOptions.escapeTargetPrefixesCsv)
    Else
        ResolveEscapeTargetPrefixesCsvRaw = ESCAPE_TARGET_PREFIXES_CSV
    End If
End Function

' UI指定があればそれを使い、なければソース定数の塗りつぶし対象を返す。
Private Function ResolveOnlyAValueRowFillTargetRaw() As String
    If mUiOptions.Enabled And mUiOptions.UseOnlyAValueRowFillTarget Then
        ResolveOnlyAValueRowFillTargetRaw = CStr(mUiOptions.onlyAValueRowFillTarget)
    Else
        ResolveOnlyAValueRowFillTargetRaw = OPTION_ONLY_A_VALUE_ROW_FILL_TARGET
    End If
End Function

' UI指定があればそれを使い、なければソース定数の塗りつぶし色を返す。
Private Function ResolveOnlyAValueRowFillColorHexRaw() As String
    If mUiOptions.Enabled And mUiOptions.UseOnlyAValueRowFillColorHex Then
        ResolveOnlyAValueRowFillColorHexRaw = CStr(mUiOptions.onlyAValueRowFillColorHex)
    Else
        ResolveOnlyAValueRowFillColorHexRaw = ONLY_A_VALUE_ROW_FILL_COLOR_HEX
    End If
End Function

' 装飾対象となるExcelブックをファイル選択ダイアログから取得する。
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















