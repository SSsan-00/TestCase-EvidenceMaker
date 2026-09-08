param([string]$ModulePath = '')

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $ModulePath) {
    $ModulePath = Join-Path $repoRoot 'escape\BetaEvidenceGenerator.bas'
    if (-not (Test-Path -LiteralPath $ModulePath)) {
        $ModulePath = Join-Path $repoRoot 'forPro\BetaEvidenceGenerator.bas'
    }
}
$ModulePath = (Resolve-Path -LiteralPath $ModulePath).Path
$testRoot = Join-Path $env:TEMP ('EvidenceGrayTests_' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($testRoot) | Out-Null
$excel = $null
$workbook = $null
$component = $null
$failures = @()

# Append a test-only entry point to exercise the private planner and writer
# with their real Excel cell formatting and shared UI options.
$harness = @'
' Build source rows in an isolated workbook and report planned/actual output.
Public Function TestEvidenceGrayScan(ByVal skipGray As Boolean, ByVal gapKind As String, _
    ByVal gapCount As Long, ByVal grayColumn As Long) As Variant
    Dim src As Worksheet
    Dim body As Worksheet
    Dim target As Workbook
    Dim plan As Object
    Dim filter As Object
    Dim lastSheet As Worksheet
    Dim created As Long
    Dim summary As String
    Dim i As Long
    Dim lastRow As Long
    Dim savedError As Long
    Dim savedDescription As String
    On Error GoTo Failed

    InitializeBetaEvidenceUiOptionsForForm mUiOptions
    mUiOptions.skipGrayFilledSourceCellEnabled = skipGray
    mUiOptions.excludeOutputSheetByPatternEnabled = False
    mUiOptions.topBorderEnabled = False
    mUiOptions.rightBorderEnabled = False
    mUiOptions.slotHeight = 1
    mSlotHeight = 1
    Set mSkipSourceFillColorMap = BuildSkipSourceFillColorMap()

    Set src = ThisWorkbook.Worksheets("Source")
    Set body = ThisWorkbook.Worksheets("A1")
    src.Cells.Clear
    src.Range("A8").Value = "B1"
    src.Range("E8").Value = "first"
    src.Range("H8").Value = "1-1"
    For i = 9 To 8 + gapCount
        Select Case gapKind
            Case "gray", "other"
                src.Cells(i, grayColumn).Value = "ignored" & CStr(i)
                If gapKind = "gray" Then
                    src.Cells(i, grayColumn).Interior.Color = RGB(166, 166, 166)
                Else
                    src.Cells(i, grayColumn).Interior.Color = RGB(255, 255, 0)
                End If
            Case "split"
                If i = 108 Then
                    src.Cells(i, grayColumn).Value = "ignored"
                    src.Cells(i, grayColumn).Interior.Color = RGB(166, 166, 166)
                End If
            Case "empty-gray"
                src.Cells(i, grayColumn).Interior.Color = RGB(166, 166, 166)
            Case "partial"
                src.Cells(i, 5).Value = "ignored"
                src.Cells(i, 5).Interior.Color = RGB(166, 166, 166)
                src.Cells(i, 8).Value = "visible"
        End Select
    Next i
    lastRow = 9 + gapCount
    src.Cells(lastRow, 1).Value = "B2"
    src.Cells(lastRow, 5).Value = "last"
    src.Cells(lastRow, 8).Value = "2-1"
    Set plan = BuildPlannedEvidenceSheetNameMap(src, "individual", filter)
    Set target = Application.Workbooks.Add(xlWBATWorksheet)
    summary = ProcessReferenceSheet(src, target, body, body, "sample", False, _
        "individual", filter, created)
    Set lastSheet = FindWorksheetExact(target, "B2")
    Dim actualLast As Boolean
    Dim lastValue As String
    actualLast = Not lastSheet Is Nothing
    If actualLast Then lastValue = CStr(lastSheet.Range("A3").Value)
    TestEvidenceGrayScan = Array(plan.Exists("B2"), actualLast, lastValue, _
        Application.CountA(target.Worksheets("B1").Range("A3:B2000")))
    target.Close SaveChanges:=False
    Set mSkipSourceFillColorMap = Nothing
    ClearUiOptions
    Exit Function
Failed:
    savedError = Err.Number
    savedDescription = Err.Description
    On Error Resume Next
    If Not target Is Nothing Then target.Close SaveChanges:=False
    Set mSkipSourceFillColorMap = Nothing
    ClearUiOptions
    On Error GoTo 0
    Err.Raise savedError, "TestEvidenceGrayScan", savedDescription
End Function
'@

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Add()
    $workbook.Worksheets.Item(1).Name = 'Source'
    $workbook.Worksheets.Add().Name = 'A1'
    $workbook.SaveAs((Join-Path $testRoot 'EvidenceGrayTests.xlsm'), 52)
    $component = $workbook.VBProject.VBComponents.Import($ModulePath)
    $component.CodeModule.AddFromString($harness)

    $cases = @(
        @('99 gray E rows', $true, 'gray', 99, 5, $true, 2),
        @('100 gray E rows', $true, 'gray', 100, 5, $true, 2),
        @('101 gray H rows', $true, 'gray', 101, 8, $true, 2),
        @('100 gray A rows', $true, 'gray', 100, 1, $true, 2),
        @('gray between two 99-row gaps', $true, 'split', 199, 5, $true, 2),
        @('99 actual empty rows', $true, 'empty', 99, 5, $true, 2),
        @('100 actual empty rows', $true, 'empty', 100, 5, $true, 2),
        @('1000 actual empty rows', $true, 'empty', 1000, 5, $true, 2),
        @('100 gray empty rows', $true, 'empty-gray', 100, 5, $true, 2),
        @('gray E and visible H on same row', $true, 'partial', 1, 5, $true, 3),
        @('skip disabled', $false, 'gray', 100, 5, $true, 102),
        @('non-gray values', $true, 'other', 100, 5, $true, 102)
    )
    foreach ($case in $cases) {
        $result = $excel.Run("'" + $workbook.Name + "'!TestEvidenceGrayScan", $case[1], $case[2], $case[3], $case[4])
        $errors = @()
        if ([bool]$result[0] -ne $case[5]) { $errors += "planned B2=$($result[0])" }
        if ([bool]$result[1] -ne $case[5]) { $errors += "created B2=$($result[1])" }
        if ($case[5] -and [string]$result[2] -ne 'last') { $errors += "B2 A3=[$($result[2])]" }
        if ([int]$result[3] -ne $case[6]) { $errors += "B1 case count=$($result[3]), expected=$($case[6])" }
        if ($errors.Count) {
            $failures += "$($case[0]): $($errors -join '; ')"
            Write-Host "FAIL $($failures[-1])"
        } else {
            Write-Host "PASS $($case[0])"
        }
    }
    if ($failures.Count) { throw "$($failures.Count) evidence gray regression cases failed." }
    Write-Host 'All BetaEvidenceGenerator gray-skip tests passed.'
}
finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($obj in @($component, $workbook, $excel)) {
        if ($null -ne $obj) { [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempPrefix = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
    if ($resolved.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        [IO.Directory]::Delete($resolved, $true)
    }
}
