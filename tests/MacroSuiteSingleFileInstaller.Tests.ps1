param(
    [string]$InstallerPath = ''
)

$ErrorActionPreference = 'Stop'

function Release-ComObject($obj) {
    if ($null -ne $obj) {
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($InstallerPath)) {
    $InstallerPath = Join-Path $repoRoot 'MacroSuiteSingleFileInstaller.bas'
}

$installerFullPath = (Resolve-Path $InstallerPath).Path
$testRoot = Join-Path $env:TEMP ('MacroSuiteInstallerTests_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null

$workbookPath = Join-Path $testRoot 'InstallerTests.xlsm'
$helperPath = Join-Path $testRoot 'dismiss-msgbox.ps1'
$helperScript = @"
Add-Type -AssemblyName System.Windows.Forms
`$shell = New-Object -ComObject WScript.Shell
`$end = (Get-Date).AddSeconds(60)
while ((Get-Date) -lt `$end) {
    try {
        if (`$shell.AppActivate('Microsoft Excel')) {
            [Windows.Forms.SendKeys]::SendWait('{ENTER}')
        }
    } catch {}
    Start-Sleep -Milliseconds 400
}
"@
Set-Content -LiteralPath $helperPath -Value $helperScript -Encoding UTF8

$excel = $null
$workbook = $null
$helperProcess = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 1 } catch {}

    $workbook = $excel.Workbooks.Add()
    $workbook.SaveAs($workbookPath, 52)
    $workbook.VBProject.VBComponents.Import($installerFullPath) | Out-Null
    $workbook.Save()

    $helperProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $helperPath
    ) -WindowStyle Hidden -PassThru
    $excel.Run("'" + $workbook.Name + "'!InstallMacroSuiteFromSingleFile")
    if ($helperProcess -and -not $helperProcess.HasExited) {
        Stop-Process -Id $helperProcess.Id -Force
    }
    $helperProcess = $null

    $componentNames = @(
        'BetaEvidenceGenerator',
        'BetaTestCaseGenerator',
        'ConditionalBranchChecker',
        'EscapePartsMarking',
        'MacroToolsUserFormInstaller'
    )

    foreach ($componentName in $componentNames) {
        $component = $workbook.VBProject.VBComponents.Item($componentName)
        if ($component.CodeModule.CountOfLines -le 0) {
            throw "Installed component is empty: $componentName"
        }
        Write-Host ("Installed: {0} ({1} lines)" -f $componentName, $component.CodeModule.CountOfLines)
    }

    # Exercise the generated form code and CONFIG defaults, not only the modules.
    $helperProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $helperPath
    ) -WindowStyle Hidden -PassThru
    $excel.Run("'" + $workbook.Name + "'!InstallMacroToolsUserForm")
    $excel.Run("'" + $workbook.Name + "'!CreateMacroConfigSheet")
    if ($helperProcess -and -not $helperProcess.HasExited) {
        Stop-Process -Id $helperProcess.Id -Force
    }
    $helperProcess = $null
    $form = $workbook.VBProject.VBComponents.Item('frmMacroTools')
    $config = $workbook.Worksheets.Item('CONFIG')
    if ($form.Designer.Controls.Item('fraConditional').Controls.Item('chkWriteIndividualSheet').Value) {
        throw 'Form individual-sheet output default must be off.'
    }
    if ([string]$config.Range('B28').Value2 -ne 'FALSE') {
        throw 'CONFIG individual-sheet output default must be off.'
    }
    $orderControls = $form.Designer.Controls.Item('fraTestCase').Controls
    if (-not $orderControls.Item('optSheetOrderFeature').Value -or $orderControls.Item('optSheetOrderSourceLast').Value) {
        throw 'Form sheet-order default mismatch.'
    }
    $orderControls.Item('optSheetOrderSourceLast').Value = $true
    if ($orderControls.Item('optSheetOrderFeature').Value) {
        throw 'Sheet-order radio buttons must be exclusive.'
    }
    if ($config.Range('B24').Value2 -ne ([string][char]0x6A5F + [char]0x80FD + [char]0x5225)) {
        throw 'CONFIG sheet-order default mismatch.'
    }
    if ([string]$config.Range('B29').Value2 -ne 'FALSE') {
        throw 'CONFIG identifier default must be off.'
    }
    $compileControl = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compileControl) {
        throw 'VBE compile command was not found.'
    }

    if ($compileControl.Enabled) {
        $helperProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $helperPath
        ) -WindowStyle Hidden -PassThru
        $compileControl.Execute()
        Start-Sleep -Seconds 2
        if ($helperProcess -and -not $helperProcess.HasExited) {
            Stop-Process -Id $helperProcess.Id -Force
        }
        $helperProcess = $null
    }

    if ($compileControl.Enabled) {
        throw 'VBA project remains uncompiled after compile command.'
    }

    Write-Host 'Single-file installation and VBE compile passed.'
}
finally {
    try {
        if ($helperProcess -and -not $helperProcess.HasExited) {
            Stop-Process -Id $helperProcess.Id -Force
        }
    } catch {}
    try { if ($workbook) { $workbook.Close($false) } } catch {}
    try { if ($excel) { $excel.Quit() } } catch {}

    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}
