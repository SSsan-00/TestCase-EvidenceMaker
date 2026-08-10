param(
    [string]$ModulePath = ''
)

$ErrorActionPreference = 'Stop'

function Write-TestLog([string]$message) {
    Write-Host ((Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' ' + $message)
}

function Release-ComObject($obj) {
    if ($null -ne $obj) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
    }
}

function Assert-True([bool]$condition, [string]$message) {
    if (-not $condition) { throw $message }
}

function Assert-Equal($expected, $actual, [string]$message) {
    if ($expected -ne $actual) {
        throw "$message Expected=[$expected] Actual=[$actual]"
    }
}

function Assert-Contains([string]$text, [string]$expected, [string]$message) {
    Assert-True ($text.Contains($expected)) $message
}

function Assert-NotContains([string]$text, [string]$unexpected, [string]$message) {
    Assert-True (-not $text.Contains($unexpected)) $message
}

function Get-OpenWorkbookByFullName($excel, [string]$path) {
    foreach ($workbook in $excel.Workbooks) {
        if ([string]::Equals($workbook.FullName, $path, [StringComparison]::OrdinalIgnoreCase)) {
            return $workbook
        }
        Release-ComObject $workbook
    }
    return $null
}

function New-ConditionalTargetWorkbook(
    $excel,
    [string]$path,
    [string]$currentSourceSheetName,
    [string]$individualSheetName
) {
    $workbook = $null
    $sourceSheet = $null
    $individualSheet = $null

    try {
        $workbook = $excel.Workbooks.Add()
        while ($workbook.Worksheets.Count -gt 1) {
            $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
        }

        $sourceSheet = $workbook.Worksheets.Item(1)
        $sourceSheet.Name = $currentSourceSheetName
        $sourceSheet.Range('C1').Value2 = 'function foo() {'
        $sourceSheet.Range('C2').Value2 = 'if ($enabled) {'
        $sourceSheet.Range('C3').Value2 = '}'
        $sourceSheet.Range('C4').Value2 = '}'

        $individualSheet = $workbook.Worksheets.Add()
        $individualSheet.Name = $individualSheetName
        for ($row = 9; $row -le 100; $row++) {
            $individualSheet.Range("A$row`:`D$row").Merge()
        }
        $individualSheet.Range('A9').Value2 = 'KEEP_A9'
        $individualSheet.Range('M9').Value2 = 'KEEP_M9'
        $individualSheet.Range('N10').Value2 = 'KEEP_N10'

        $workbook.SaveAs($path, 51)
        $workbook.Close($false)
    }
    finally {
        Release-ComObject $individualSheet
        Release-ComObject $sourceSheet
        Release-ComObject $workbook
    }
}

function Start-MsgBoxDismissHelper([string]$helperPath) {
    return Start-Process -FilePath 'powershell.exe' `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $helperPath) `
        -WindowStyle Hidden `
        -PassThru
}

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ModulePath)) {
    $candidates = @(
        (Join-Path $repoRoot 'escape\ConditionalBranchChecker.bas'),
        (Join-Path $repoRoot 'forPro\ConditionalBranchChecker.bas')
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            $ModulePath = $candidate
            break
        }
    }
}

if ([string]::IsNullOrWhiteSpace($ModulePath)) {
    throw 'ConditionalBranchChecker.bas was not found under escape or forPro.'
}

$moduleFullPath = (Resolve-Path $ModulePath).Path
$cp932 = [Text.Encoding]::GetEncoding(932)
$moduleText = [IO.File]::ReadAllText($moduleFullPath, $cp932)

Assert-NotContains $moduleText 'LEADING_FUNCTION_STARTS_FROM_B1' 'The legacy leading-function constant must be removed.'
Assert-NotContains $moduleText 'OverrideLeadingFunctionStartsFromB1' 'The legacy leading-function override must be removed.'
Assert-NotContains $moduleText 'leadingFunctionStartsFromB1' 'The legacy leading-function option must be removed.'
Assert-Contains $moduleText 'WRITE_INDIVIDUAL_SHEET_ENABLED' 'The individual-sheet output default is missing.'
Assert-Contains $moduleText 'OverrideWriteIndividualSheetEnabled' 'The individual-sheet output override is missing.'
Assert-Contains $moduleText 'writeIndividualSheetEnabled' 'The individual-sheet output option is missing.'

$testRoot = Join-Path $env:TEMP ('ConditionalBranchCheckerTests_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null

$macroWorkbookPath = Join-Path $testRoot 'ConditionalBranchCheckerTests.xlsm'
$targetOnPath = Join-Path $testRoot 'ConditionalOutputOn.xlsx'
$targetOffPath = Join-Path $testRoot 'ConditionalOutputOff.xlsx'
$helperPath = Join-Path $testRoot 'dismiss-msgbox.ps1'

$helperScript = @"
Add-Type -AssemblyName System.Windows.Forms
`$ws = New-Object -ComObject WScript.Shell
`$end = (Get-Date).AddSeconds(90)
while ((Get-Date) -lt `$end) {
    try {
        if (`$ws.AppActivate('Microsoft Excel')) {
            [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
        }
    } catch {}
    Start-Sleep -Milliseconds 500
}
"@
[IO.File]::WriteAllText($helperPath, $helperScript, (New-Object Text.UTF8Encoding($false)))

$harnessCode = @"
Option Explicit

Public Sub RunConditionalBranchCheckerForTest(ByVal targetPath As String, ByVal writeIndividualSheet As Boolean)
    Dim options As ConditionalBranchCheckerUiOptions

    options.Enabled = True
    options.featureName = "FeatureA"
    options.workbookPath = targetPath
    options.OverrideWriteIndividualSheetEnabled = True
    options.writeIndividualSheetEnabled = writeIndividualSheet
    options.OverrideMarkNonFunctionLineWithDash = True
    options.markNonFunctionLineWithDash = True
    options.OverrideMarkFillEnabled = True
    options.markFillEnabled = False
    options.UseMarkFillColorHex = True
    options.markFillColorHex = "#fff2cc"

    RunMainWithUiOptions options
End Sub
"@

$currentSourceSheetName = -join @(
    [char]0x73FE, [char]0x884C, [char]0x30BD, [char]0x30FC, [char]0x30B9
)
$individualPrefix = -join @(
    [char]0x3010, [char]0x500B, [char]0x5225, [char]0x3011
)
$individualSheetName = $individualPrefix + 'FeatureA'
$functionMark = [string][char]0x2605

$excel = $null
$macroWorkbook = $null
$verifyWorkbook = $null
$helperProcess = $null

try {
    Write-TestLog 'Starting ConditionalBranchChecker tests.'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 1 } catch {}

    $macroWorkbook = $excel.Workbooks.Add()
    $macroWorkbook.SaveAs($macroWorkbookPath, 52)
    $vbProject = $macroWorkbook.VBProject
    $vbProject.VBComponents.Import($moduleFullPath) | Out-Null
    $harnessComponent = $vbProject.VBComponents.Add(1)
    $harnessComponent.Name = 'ConditionalTestHarness'
    $harnessComponent.CodeModule.AddFromString($harnessCode)
    $macroWorkbook.Save()

    New-ConditionalTargetWorkbook $excel $targetOnPath $currentSourceSheetName $individualSheetName
    New-ConditionalTargetWorkbook $excel $targetOffPath $currentSourceSheetName $individualSheetName

    $helperProcess = Start-MsgBoxDismissHelper $helperPath
    $excel.Run("'" + $macroWorkbook.Name + "'!RunConditionalBranchCheckerForTest", $targetOnPath, $true)
    if ($helperProcess -and -not $helperProcess.HasExited) { Stop-Process -Id $helperProcess.Id -Force }
    $helperProcess = $null

    $openedTarget = Get-OpenWorkbookByFullName $excel $targetOnPath
    if ($openedTarget) {
        $openedTarget.Close($false)
        Release-ComObject $openedTarget
    }

    $verifyWorkbook = $excel.Workbooks.Open($targetOnPath, $false, $true)
    $sourceSheet = $verifyWorkbook.Worksheets.Item($currentSourceSheetName)
    $individualSheet = $verifyWorkbook.Worksheets.Item($individualSheetName)

    Assert-Equal $functionMark ([string]$sourceSheet.Range('A1').Value2) 'Function mark mismatch when output is enabled.'
    Assert-Equal 'B1' ([string]$sourceSheet.Range('B1').Value2) 'The first function must always start at B1.'
    Assert-Equal 'B1-' ([string]$sourceSheet.Range('B2').Value2) 'Non-function branch mark mismatch.'
    Assert-Equal 'B1' ([string]$individualSheet.Range('A9').Value2) 'Individual-sheet section number mismatch.'
    Assert-True ([string]$individualSheet.Range('M9').Value2 -ne 'KEEP_M9') 'Individual-sheet header was not written.'
    Assert-True (-not [string]::IsNullOrWhiteSpace([string]$individualSheet.Range('N10').Value2)) 'Individual-sheet IF block was not written.'

    $verifyWorkbook.Close($false)
    Release-ComObject $individualSheet
    Release-ComObject $sourceSheet
    Release-ComObject $verifyWorkbook
    $verifyWorkbook = $null

    $helperProcess = Start-MsgBoxDismissHelper $helperPath
    $excel.Run("'" + $macroWorkbook.Name + "'!RunConditionalBranchCheckerForTest", $targetOffPath, $false)
    if ($helperProcess -and -not $helperProcess.HasExited) { Stop-Process -Id $helperProcess.Id -Force }
    $helperProcess = $null

    $openedTarget = Get-OpenWorkbookByFullName $excel $targetOffPath
    if ($openedTarget) {
        $openedTarget.Close($false)
        Release-ComObject $openedTarget
    }

    $verifyWorkbook = $excel.Workbooks.Open($targetOffPath, $false, $true)
    $sourceSheet = $verifyWorkbook.Worksheets.Item($currentSourceSheetName)
    $individualSheet = $verifyWorkbook.Worksheets.Item($individualSheetName)

    Assert-Equal $functionMark ([string]$sourceSheet.Range('A1').Value2) 'Function mark mismatch when output is disabled.'
    Assert-Equal 'B1' ([string]$sourceSheet.Range('B1').Value2) 'The first function must remain B1 when output is disabled.'
    Assert-Equal 'B1-' ([string]$sourceSheet.Range('B2').Value2) 'Branch marking must run when output is disabled.'
    Assert-Equal 'KEEP_A9' ([string]$individualSheet.Range('A9').Value2) 'Individual sheet was modified while output was disabled.'
    Assert-Equal 'KEEP_M9' ([string]$individualSheet.Range('M9').Value2) 'Individual sheet header was modified while output was disabled.'
    Assert-Equal 'KEEP_N10' ([string]$individualSheet.Range('N10').Value2) 'Individual sheet block was modified while output was disabled.'

    Write-TestLog 'All ConditionalBranchChecker tests passed.'
}
finally {
    try { if ($helperProcess -and -not $helperProcess.HasExited) { Stop-Process -Id $helperProcess.Id -Force } } catch {}
    try { if ($verifyWorkbook) { $verifyWorkbook.Close($false) } } catch {}
    try { if ($macroWorkbook) { $macroWorkbook.Close($false) } } catch {}
    try { if ($excel) { $excel.Quit() } } catch {}

    Release-ComObject $individualSheet
    Release-ComObject $sourceSheet
    Release-ComObject $verifyWorkbook
    Release-ComObject $macroWorkbook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
