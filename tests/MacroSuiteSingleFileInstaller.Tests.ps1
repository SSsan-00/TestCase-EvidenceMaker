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
