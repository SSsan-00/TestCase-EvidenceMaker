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

function Get-CellTextLength($cell) {
    return ([string]$cell.Value2).Length
}

function Assert-CellFullyRedBold($cell, [string]$address) {
    $length = Get-CellTextLength $cell
    Assert-True ($length -gt 0) "$address should contain text."

    $font = $cell.Characters(1, $length).Font
    $isBold = [bool]$font.Bold
    $color = [int]$font.Color

    Assert-True $isBold "$address should be bold."
    Assert-Equal 255 $color "$address should be red."
}

function Assert-CharactersRedBold($cell, [int]$start, [int]$length, [string]$label) {
    $font = $cell.Characters($start, $length).Font
    Assert-True ([bool]$font.Bold) "$label should be bold."
    Assert-Equal 255 ([int]$font.Color) "$label should be red."
}

function Assert-CharacterUnmarked($cell, [int]$position, [string]$label) {
    $font = $cell.Characters($position, 1).Font
    $isMarked = ([bool]$font.Bold) -and ([int]$font.Color -eq 255)
    Assert-True (-not $isMarked) "$label should not be marked."
}

function Assert-CellNotHit($sheet, [string]$sourceAddress, [string]$messageAddress) {
    $sourceCell = $sheet.Range($sourceAddress)
    $sourceLength = Get-CellTextLength $sourceCell
    if ($sourceLength -gt 0) {
        for ($position = 1; $position -le $sourceLength; $position++) {
            Assert-CharacterUnmarked $sourceCell $position "$sourceAddress character $position"
        }
    }

    Assert-Equal '' ([string]$sheet.Range($messageAddress).Value2) "$messageAddress should remain empty."
}

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ModulePath)) {
    $candidates = @(
        (Join-Path $repoRoot 'escape\EscapePartsMarking.bas'),
        (Join-Path $repoRoot 'forPro\EscapePartsMarking.bas')
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            $ModulePath = $candidate
            break
        }
    }
}

if ([string]::IsNullOrWhiteSpace($ModulePath)) {
    throw 'EscapePartsMarking.bas was not found under escape or forPro.'
}

$moduleFullPath = (Resolve-Path $ModulePath).Path
$testRoot = Join-Path $env:TEMP ('EscapePartsMarkingTests_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null

$macroWorkbookPath = Join-Path $testRoot 'EscapePartsMarkingTests.xlsm'
$targetWorkbookPath = Join-Path $testRoot 'Target.xlsx'
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
Set-Content -LiteralPath $helperPath -Value $helperScript -Encoding UTF8

$harnessCode = @"
Option Explicit

Public Sub RunEscapePartsMarkingForTest(ByVal targetPath As String)
    Dim options As EscapePartsMarkingUiOptions

    options.Enabled = True
    options.TargetWorkbookPath = targetPath
    options.UseCompletionMessage = True
    options.completionMessage = "HIT"
    options.UseEscapeTargetPrefixesCsv = True
    options.escapeTargetPrefixesCsv = "sqlS"
    options.UseOnlyAValueRowFillTarget = True
    options.onlyAValueRowFillTarget = "Both"
    options.UseOnlyAValueRowFillColorHex = True
    options.onlyAValueRowFillColorHex = "#a6a6a6"

    RunMainWithUiOptions options
End Sub
"@

$excel = $null
$macroWb = $null
$targetWb = $null
$verifyWb = $null
$helperProcess = $null

try {
    Write-TestLog 'Starting EscapePartsMarking tests.'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    try { $excel.AutomationSecurity = 1 } catch {}

    $macroWb = $excel.Workbooks.Add()
    $macroWb.SaveAs($macroWorkbookPath, 52)
    $vbProj = $macroWb.VBProject
    $vbProj.VBComponents.Import($moduleFullPath) | Out-Null
    $harnessComponent = $vbProj.VBComponents.Add(1)
    $harnessComponent.Name = 'EscapePartsMarkingTestHarness'
    $harnessComponent.CodeModule.AddFromString($harnessCode)
    $macroWb.Save()
    Write-TestLog 'Imported EscapePartsMarking module and harness.'

    $targetWb = $excel.Workbooks.Add()
    while ($targetWb.Worksheets.Count -gt 1) {
        $targetWb.Worksheets.Item($targetWb.Worksheets.Count).Delete()
    }
    $ws = $targetWb.Worksheets.Item(1)
    $ws.Name = 'A1-1-1'

    $ws.Range('B4').Value2 = 'sqlS('
    $ws.Range('B5').Value2 = 'XXX'
    $ws.Range('B6').Value2 = ')'

    $ws.Range('B8').Value2 = 'sqlS(xxx + trim(yyy) + "zzz")'

    $ws.Range('B10').Value2 = 'mySqlS(value)'
    $ws.Range('B11').Value2 = 'message = "sqlS(value)"'
    $ws.Range('B12').Value2 = '// sqlS(value)'
    $ws.Range('B13').Value2 = '# sqlS(value)'
    $ws.Range('B14').Value2 = '/* sqlS(value)'
    $ws.Range('B15').Value2 = 'still a comment */'

    $ws.Range('B17').Value2 = 'sqlS('
    $ws.Range('B18').Value2 = 'value // ) ignored'
    $ws.Range('B19').Value2 = ')'

    $ws.Range('B21').Value2 = 'before sqlS(value) after'

    $ws.Range('B23').Value2 = 'DbHelper.sqlS('
    $ws.Range('B24').Value2 = 'trim(value)'
    $ws.Range('B25').Value2 = ')'

    $ws.Range('B27').Formula = '=NA()'
    $ws.Range('B27').Calculate()
    $ws.Range('B28').Value2 = 'sqlS(afterError)'

    $ws.Range('B30').Value2 = 'sqlS(one) + sqlS(two)'
    $ws.Range('B32').Value2 = 'sqlS(")")'

    $ws.Range('A34').Value2 = 'ONLY_A'

    $normalInterpolated = '$"{sqlS(XXX)}"'
    $normalInterpolatedLiteral = '$"literal sqlS(XXX)"'
    $normalInterpolatedEscapedBraces = '$"{{sqlS(XXX)}}"'
    $verbatimInterpolatedDollarFirst = '$@"path {sqlS(XXX)}"'
    $verbatimInterpolatedAtFirst = '@$"path {sqlS(XXX)}"'
    $rawInterpolated = '$"""raw {sqlS(XXX)} text"""'
    $rawInterpolatedLiteral = '$"""sqlS(XXX)"""'
    $rawDoubleDollar = '$$"""json { "value": {{sqlS(XXX)}} }"""'
    $rawDoubleDollarLiteral = '$$"""literal {sqlS(XXX)}"""'
    $rawDoubleDollarOuterBrace = '$$"""literal {{{sqlS(XXX)}}}"""'
    $interpolationNestedString = '$"{GetText("sqlS(XXX)")}"'
    $interpolationTargetWithNestedString = '$"{sqlS(GetText(")"))}"'
    $normalInterpolatedMultiple = '$"{sqlS(one)}-{sqlS(two)}"'
    $interpolationWithComment = '$"{sqlS(value /* ) */)}"'
    $nonInterpolatedRaw = '"""sqlS(XXX)"""'
    $nonInterpolatedVerbatim = '@"sqlS(XXX)"'
    $rawFourQuoteDelimiter = '$""""{sqlS(XXX)}""""'
    $rawTripleDollar = '$$$"""{{{sqlS(XXX)}}}"""'
    $interpolationFormatText = '$"{value:sqlS(XXX)}"'
    $rawInterpolationFormatText = '$"""{value:sqlS(XXX)}"""'
    $nestedInterpolation = '$"{Wrap($"{sqlS(XXX)}")}"'
    $rawInterpolationNestedString = '$"""{GetText("sqlS(XXX)")}"""'
    $normalInterpolationOuterBrace = '$"{{{sqlS(XXX)}}}"'
    $interpolationObjectInitializer = '$"{sqlS(new Item { Value = XXX })}"'
    $rawFourQuoteWithTripleQuoteContent = '$""""raw """ text {sqlS(XXX)}""""'

    $ws.Range('B36').Value2 = $normalInterpolated
    $ws.Range('B38').Value2 = $normalInterpolatedLiteral
    $ws.Range('B40').Value2 = $normalInterpolatedEscapedBraces
    $ws.Range('B42').Value2 = $verbatimInterpolatedDollarFirst
    $ws.Range('B44').Value2 = $verbatimInterpolatedAtFirst
    $ws.Range('B46').Value2 = $rawInterpolated
    $ws.Range('B48').Value2 = $rawInterpolatedLiteral
    $ws.Range('B50').Value2 = $rawDoubleDollar
    $ws.Range('B52').Value2 = $rawDoubleDollarLiteral
    $ws.Range('B54').Value2 = $rawDoubleDollarOuterBrace
    $ws.Range('B56').Value2 = $interpolationNestedString
    $ws.Range('B58').Value2 = $interpolationTargetWithNestedString
    $ws.Range('B60').Value2 = $normalInterpolatedMultiple

    $ws.Range('B62').Value2 = '$@"start {sqlS('
    $ws.Range('B63').Value2 = 'XXX'
    $ws.Range('B64').Value2 = ')} end"'

    $ws.Range('B66').Value2 = '$"""'
    $ws.Range('B67').Value2 = 'before {sqlS('
    $ws.Range('B68').Value2 = 'XXX'
    $ws.Range('B69').Value2 = ')}'
    $ws.Range('B70').Value2 = 'after'
    $ws.Range('B71').Value2 = '"""'

    $ws.Range('B73').Value2 = '$$"""'
    $ws.Range('B74').Value2 = '{ "x": {{sqlS('
    $ws.Range('B75').Value2 = 'XXX'
    $ws.Range('B76').Value2 = ')}} }'
    $ws.Range('B77').Value2 = '"""'

    $ws.Range('B79').Value2 = $interpolationWithComment
    $ws.Range('B81').Value2 = $nonInterpolatedRaw
    $ws.Range('B82').Value2 = $nonInterpolatedVerbatim
    $ws.Range('B83').Value2 = $rawFourQuoteDelimiter
    $ws.Range('B85').Value2 = $rawTripleDollar
    $ws.Range('B87').Value2 = $interpolationFormatText
    $ws.Range('B88').Value2 = $rawInterpolationFormatText
    $ws.Range('B89').Value2 = $nestedInterpolation
    $ws.Range('B91').Value2 = $rawInterpolationNestedString
    $ws.Range('B93').Value2 = $normalInterpolationOuterBrace
    $ws.Range('B95').Value2 = $interpolationObjectInitializer
    $ws.Range('B97').Value2 = $rawFourQuoteWithTripleQuoteContent

    $targetWb.SaveAs($targetWorkbookPath, 51)
    $targetWb.Close($false)
    Release-ComObject $targetWb
    $targetWb = $null
    Write-TestLog 'Created target workbook.'

    $helperProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File', $helperPath) -WindowStyle Hidden -PassThru
    $excel.Run("'" + $macroWb.Name + "'!RunEscapePartsMarkingForTest", $targetWorkbookPath)
    if ($helperProcess -and -not $helperProcess.HasExited) { Stop-Process -Id $helperProcess.Id -Force }
    $helperProcess = $null
    Write-TestLog 'Ran EscapePartsMarking macro.'

    $verifyWb = $excel.Workbooks.Open($targetWorkbookPath, $false, $true)
    $verifyWs = $verifyWb.Worksheets.Item('A1-1-1')

    Assert-CellFullyRedBold $verifyWs.Range('B4') 'B4'
    Assert-CellFullyRedBold $verifyWs.Range('B5') 'B5'
    Assert-CellFullyRedBold $verifyWs.Range('B6') 'B6'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C4').Value2) 'C4 hit message mismatch.'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C5').Value2) 'C5 hit message mismatch.'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C6').Value2) 'C6 hit message mismatch.'

    Assert-CellFullyRedBold $verifyWs.Range('B8') 'B8'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C8').Value2) 'C8 hit message mismatch.'

    Assert-CellNotHit $verifyWs 'B10' 'C10'
    Assert-CellNotHit $verifyWs 'B11' 'C11'
    Assert-CellNotHit $verifyWs 'B12' 'C12'
    Assert-CellNotHit $verifyWs 'B13' 'C13'
    Assert-CellNotHit $verifyWs 'B14' 'C14'
    Assert-CellNotHit $verifyWs 'B15' 'C15'

    Assert-CellFullyRedBold $verifyWs.Range('B17') 'B17'
    Assert-CellFullyRedBold $verifyWs.Range('B18') 'B18'
    Assert-CellFullyRedBold $verifyWs.Range('B19') 'B19'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C17').Value2) 'C17 hit message mismatch.'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C18').Value2) 'C18 hit message mismatch.'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C19').Value2) 'C19 hit message mismatch.'

    Assert-CharacterUnmarked $verifyWs.Range('B21') 1 'B21 leading text'
    Assert-CharactersRedBold $verifyWs.Range('B21') 8 11 'B21 function call'
    Assert-CharacterUnmarked $verifyWs.Range('B21') 20 'B21 trailing text'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C21').Value2) 'C21 hit message mismatch.'

    Assert-CellFullyRedBold $verifyWs.Range('B23') 'B23'
    Assert-CellFullyRedBold $verifyWs.Range('B24') 'B24'
    Assert-CellFullyRedBold $verifyWs.Range('B25') 'B25'

    Assert-Equal '' ([string]$verifyWs.Range('C27').Value2) 'C27 should remain empty.'
    Assert-CellFullyRedBold $verifyWs.Range('B28') 'B28'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C28').Value2) 'C28 hit message mismatch.'

    Assert-CharactersRedBold $verifyWs.Range('B30') 1 9 'B30 first function call'
    Assert-CharacterUnmarked $verifyWs.Range('B30') 11 'B30 operator'
    Assert-CharactersRedBold $verifyWs.Range('B30') 13 9 'B30 second function call'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C30').Value2) 'C30 hit message mismatch.'

    Assert-CellFullyRedBold $verifyWs.Range('B32') 'B32'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C32').Value2) 'C32 hit message mismatch.'

    Assert-Equal 10921638 ([int]$verifyWs.Range('A34').Interior.Color) 'A34 fill color mismatch.'
    Assert-Equal 10921638 ([int]$verifyWs.Range('B34').Interior.Color) 'B34 fill color mismatch.'

    Assert-CharactersRedBold $verifyWs.Range('B36') ($normalInterpolated.IndexOf('sqlS') + 1) 9 'B36 interpolated function call'
    Assert-CharacterUnmarked $verifyWs.Range('B36') 1 'B36 interpolation prefix'
    Assert-Equal 'HIT' ([string]$verifyWs.Range('C36').Value2) 'C36 hit message mismatch.'

    Assert-CellNotHit $verifyWs 'B38' 'C38'
    Assert-CellNotHit $verifyWs 'B40' 'C40'

    Assert-CharactersRedBold $verifyWs.Range('B42') ($verbatimInterpolatedDollarFirst.IndexOf('sqlS') + 1) 9 'B42 verbatim interpolated function call'
    Assert-CharactersRedBold $verifyWs.Range('B44') ($verbatimInterpolatedAtFirst.IndexOf('sqlS') + 1) 9 'B44 reversed verbatim interpolated function call'
    Assert-CharactersRedBold $verifyWs.Range('B46') ($rawInterpolated.IndexOf('sqlS') + 1) 9 'B46 raw interpolated function call'
    Assert-CellNotHit $verifyWs 'B48' 'C48'
    Assert-CharactersRedBold $verifyWs.Range('B50') ($rawDoubleDollar.IndexOf('sqlS') + 1) 9 'B50 double-dollar raw function call'
    Assert-CellNotHit $verifyWs 'B52' 'C52'
    Assert-CharactersRedBold $verifyWs.Range('B54') ($rawDoubleDollarOuterBrace.IndexOf('sqlS') + 1) 9 'B54 raw function call with outer literal brace'
    Assert-CellNotHit $verifyWs 'B56' 'C56'
    Assert-CharactersRedBold $verifyWs.Range('B58') ($interpolationTargetWithNestedString.IndexOf('sqlS') + 1) 18 'B58 interpolated call with nested string'

    $firstInterpolatedCall = $normalInterpolatedMultiple.IndexOf('sqlS') + 1
    $secondInterpolatedCall = $normalInterpolatedMultiple.IndexOf('sqlS', $firstInterpolatedCall) + 1
    Assert-CharactersRedBold $verifyWs.Range('B60') $firstInterpolatedCall 9 'B60 first interpolated function call'
    Assert-CharactersRedBold $verifyWs.Range('B60') $secondInterpolatedCall 9 'B60 second interpolated function call'

    Assert-CharactersRedBold $verifyWs.Range('B62') 11 5 'B62 multiline verbatim call start'
    Assert-CellFullyRedBold $verifyWs.Range('B63') 'B63'
    Assert-CharactersRedBold $verifyWs.Range('B64') 1 1 'B64 multiline verbatim call end'

    Assert-CharactersRedBold $verifyWs.Range('B67') 9 5 'B67 multiline raw call start'
    Assert-CellFullyRedBold $verifyWs.Range('B68') 'B68'
    Assert-CharactersRedBold $verifyWs.Range('B69') 1 1 'B69 multiline raw call end'
    Assert-CellNotHit $verifyWs 'B70' 'C70'

    Assert-CharactersRedBold $verifyWs.Range('B74') 10 5 'B74 multiline double-dollar raw call start'
    Assert-CellFullyRedBold $verifyWs.Range('B75') 'B75'
    Assert-CharactersRedBold $verifyWs.Range('B76') 1 1 'B76 multiline double-dollar raw call end'

    Assert-CharactersRedBold $verifyWs.Range('B79') ($interpolationWithComment.IndexOf('sqlS') + 1) 19 'B79 interpolated call containing comment'
    Assert-CellNotHit $verifyWs 'B81' 'C81'
    Assert-CellNotHit $verifyWs 'B82' 'C82'
    Assert-CharactersRedBold $verifyWs.Range('B83') ($rawFourQuoteDelimiter.IndexOf('sqlS') + 1) 9 'B83 four-quote raw delimiter'
    Assert-CharactersRedBold $verifyWs.Range('B85') ($rawTripleDollar.IndexOf('sqlS') + 1) 9 'B85 triple-dollar raw function call'
    Assert-CellNotHit $verifyWs 'B87' 'C87'
    Assert-CellNotHit $verifyWs 'B88' 'C88'
    Assert-CharactersRedBold $verifyWs.Range('B89') ($nestedInterpolation.IndexOf('sqlS') + 1) 9 'B89 nested interpolated function call'
    Assert-CellNotHit $verifyWs 'B91' 'C91'
    Assert-CharactersRedBold $verifyWs.Range('B93') ($normalInterpolationOuterBrace.IndexOf('sqlS') + 1) 9 'B93 interpolated function call with outer literal brace'
    $objectInitializerCallLength = $interpolationObjectInitializer.LastIndexOf(')') - $interpolationObjectInitializer.IndexOf('sqlS') + 1
    Assert-CharactersRedBold $verifyWs.Range('B95') ($interpolationObjectInitializer.IndexOf('sqlS') + 1) $objectInitializerCallLength 'B95 interpolated call with object initializer'
    Assert-CharactersRedBold $verifyWs.Range('B97') ($rawFourQuoteWithTripleQuoteContent.IndexOf('sqlS') + 1) 9 'B97 raw function call after shorter quote run'

    Write-TestLog 'All EscapePartsMarking tests passed.'
}
finally {
    try { if ($helperProcess -and -not $helperProcess.HasExited) { Stop-Process -Id $helperProcess.Id -Force } } catch {}
    try { if ($verifyWb -ne $null) { $verifyWb.Close($false) } } catch {}
    try { if ($targetWb -ne $null) { $targetWb.Close($false) } } catch {}
    try { if ($macroWb -ne $null) { $macroWb.Close($false) } } catch {}
    try { if ($excel -ne $null) { $excel.Quit() } } catch {}

    Release-ComObject $verifyWb
    Release-ComObject $targetWb
    Release-ComObject $macroWb
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
