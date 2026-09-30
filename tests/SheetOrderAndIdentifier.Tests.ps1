$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$temp = Join-Path $env:TEMP ('SheetOrderTests_' + [guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($temp) | Out-Null
$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $wb = $excel.Workbooks.Add()
    $wb.SaveAs((Join-Path $temp 'Tests.xlsm'), 52)
    foreach ($name in @('BetaTestCaseGenerator', 'ConditionalBranchChecker')) {
        $path = Join-Path $root ('escape\' + $name + '.bas')
        if (-not (Test-Path $path)) { $path = Join-Path $root ('forPro\' + $name + '.bas') }
        $component = $wb.VBProject.VBComponents.Import($path)
        if ($name -eq 'BetaTestCaseGenerator') {
            $harness = @'
Public Function TestSheetOrder(ByVal grouped As Boolean) As String
    Dim matches As New Collection
    Dim rec(1 To 4) As Variant
    Dim target As Workbook
    Dim ws As Worksheet
    Dim count As Long
    Dim i As Long
    Dim result As String
    mUiOptions = CreateBetaTestCaseUiOptionsForForm()
    mUiOptions.groupByFeature = grouped
    For i = 1 To 2
        rec(1) = "file"
        rec(2) = "Feature" & i
        rec(3) = "source" & i
        rec(4) = i
        matches.Add rec
    Next
    Set target = BuildOutputWorkbook(ThisWorkbook, matches, "file", "S00-000-00", count)
    For Each ws In target.Worksheets
        result = result & ws.Name & "|"
        If Left$(ws.Name, Len(OUTPUT_SOURCE_PREFIX)) = OUTPUT_SOURCE_PREFIX Then
            If ws.Range("C4").Value <> "source" & Right$(ws.Name, 1) Then Err.Raise 5, , "Source content mismatch"
        ElseIf Left$(ws.Name, 1) = "【" Then
            If ws.Range("BD3").Value <> "S00-000-00" Then Err.Raise 5, , "Case content mismatch"
        End If
    Next
    If count <> 8 Then Err.Raise 5, , "Sheet count mismatch"
    target.Close False
    TestSheetOrder = result
End Function
Public Function TestGroupedDefault() As Boolean
    Dim options As BetaTestCaseUiOptions
    options = CreateBetaTestCaseUiOptionsForForm()
    TestGroupedDefault = options.groupByFeature
End Function
'@
        } else {
            $harness = @'
Public Function TestIdentifiers(ByVal enabled As Boolean, ByVal leadingBranch As Boolean) As Variant
    Dim ws As Worksheet
    Dim values As Variant
    Dim marked As Long
    Dim firstRow As Long
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Cells.Clear
    mUiOptions = CreateConditionalBranchCheckerUiOptionsForForm()
    If mUiOptions.writeBranchIdentifierEnabled Then Err.Raise 5, , "Identifier default must be off"
    mUiOptions.writeBranchIdentifierEnabled = enabled
    mUiOptions.markFillEnabled = True
    mUiOptions.markFillColorHex = "#a6a6a6"
    firstRow = 1
    If leadingBranch Then
        ws.Range("C1").Value = "if ($x) {"
        firstRow = 2
    End If
    ws.Cells(firstRow, 3).Value = "function foo() {"
    ws.Cells(firstRow + 1, 3).Value = "if ($x) {"
    ws.Cells(firstRow + 2, 3).Value = "function bar() {"
    ws.Range("B1:B4").Value = "stale"
    values = ws.Range("C1:C4").Value
    MarkCurrentSourceSheet ws, values, 4, marked, Not leadingBranch
    If ws.Cells(firstRow, 1).Value <> "★" Then Err.Raise 5, , "Function star missing"
    If ws.Cells(firstRow + 2, 1).Value <> "★" Then Err.Raise 5, , "Second function star missing"
    If enabled Then
        If ws.Cells(firstRow, 2).Value <> "B" & firstRow Then Err.Raise 5, , "First identifier mismatch"
        If ws.Cells(firstRow + 1, 2).Value <> "B" & firstRow & "-" Then Err.Raise 5, , "Branch identifier mismatch"
        If ws.Cells(firstRow + 2, 2).Value <> "B" & firstRow + 1 Then Err.Raise 5, , "Second identifier mismatch"
    Else
        If Application.CountA(ws.Range("B1:B4")) <> 0 Then Err.Raise 5, , "Identifiers were written while disabled"
    End If
    If ws.Cells(firstRow, 2).Interior.Color <> RGB(166, 166, 166) Then Err.Raise 5, , "Function fill missing"
    If ws.Cells(firstRow + 1, 2).Interior.Color <> RGB(166, 166, 166) Then Err.Raise 5, , "Branch fill missing"
    TestIdentifiers = True
    Exit Function
Failed:
    TestIdentifiers = "FAILED: " & Err.Description
End Function
'@
        }
        $component.CodeModule.AddFromString($harness)
    }
    foreach ($name in @('【共通】機能名','【個別】機能名','⇒参考','現行ソース（PHP）','現行画面')) {
        $ws = $wb.Worksheets.Add()
        $ws.Name = $name
    }
    if (-not $excel.Run("'Tests.xlsm'!TestGroupedDefault")) { throw 'Grouped default is off' }
    $expected = @(
        '【共通】Feature1|【個別】Feature1|現行ソース（PHP）Feature1|【共通】Feature2|【個別】Feature2|現行ソース（PHP）Feature2|⇒参考|現行画面|',
        '【共通】Feature1|【個別】Feature1|【共通】Feature2|【個別】Feature2|⇒参考|現行ソース（PHP）Feature1|現行ソース（PHP）Feature2|現行画面|'
    )
    for ($i = 0; $i -lt 2; $i++) {
        $actual = $excel.Run("'Tests.xlsm'!TestSheetOrder", ($i -eq 0))
        if ($actual -ne $expected[$i]) { throw "Sheet order mismatch: $actual" }
        Write-Host "PASS sheet order $i and contents"
    }
    foreach ($enabled in @($false, $true)) {
        foreach ($leading in @($false, $true)) {
            $result = $excel.Run("'Tests.xlsm'!TestIdentifiers", $enabled, $leading)
            if ($result -ne $true) { throw "Identifier test failed: $result" }
            Write-Host "PASS identifiers=$enabled leadingBranch=$leading"
        }
    }
} finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $resolved = [IO.Path]::GetFullPath($temp)
    if (-not $resolved.StartsWith([IO.Path]::GetFullPath($env:TEMP) + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe cleanup path' }
    [IO.Directory]::Delete($resolved, $true)
}
