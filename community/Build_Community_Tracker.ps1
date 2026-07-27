[CmdletBinding()]
param(
    [string]$TemplatePath = "",
    [string]$OutputPath = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
    $TemplatePath = Join-Path $ScriptDir "SMW_ROM_Hack_Tracker_Community.xlsx"
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $ScriptDir "SMW_ROM_Hack_Tracker_Community.xlsm"
}

$BasPath = Join-Path $ScriptDir "SMWCommunity.bas"
$ThisWorkbookCodePath = Join-Path $ScriptDir "ThisWorkbookCode.txt"

function Release-ComObjectSafely {
    param($ComObject)
    if ($null -ne $ComObject) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) } catch {}
    }
}

function Stop-WithMessage {
    param([string]$Message)
    Write-Host ""
    Write-Host $Message -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to close"
    exit 1
}

foreach ($requiredFile in @($TemplatePath, $BasPath, $ThisWorkbookCodePath)) {
    if (-not (Test-Path -LiteralPath $requiredFile)) {
        Stop-WithMessage "Required file not found: $requiredFile"
    }
}

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Stop-WithMessage "Close every open Excel window, then run BUILD_COMMUNITY_TRACKER.bat again."
}

$probe = $null
$excel = $null
$book = $null
$vbProject = $null
$components = $null
$thisWorkbookComponent = $null
$codeModule = $null
$registryPath = $null
$previousAccessVbom = $null
$hadPreviousAccessVbom = $false

try {
    Write-Host "Detecting Microsoft Excel..." -ForegroundColor Cyan
    $probe = New-Object -ComObject Excel.Application
    $excelVersion = [string]$probe.Version
    $probe.Quit()
    Release-ComObjectSafely $probe
    $probe = $null

    $registryPath = "HKCU:\Software\Microsoft\Office\$excelVersion\Excel\Security"
    if (-not (Test-Path -LiteralPath $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }

    try {
        $previousAccessVbom = (Get-ItemProperty -LiteralPath $registryPath -Name AccessVBOM -ErrorAction Stop).AccessVBOM
        $hadPreviousAccessVbom = $true
    } catch {
        $hadPreviousAccessVbom = $false
    }

    New-ItemProperty -LiteralPath $registryPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null
    Start-Sleep -Milliseconds 750

    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }

    Write-Host "Building the macro-enabled Community Tracker..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 1 } catch {}

    $book = $excel.Workbooks.Open([System.IO.Path]::GetFullPath($TemplatePath))
    $book.CheckCompatibility = $false

    # Convert the clean template into an XLSM before importing VBA.
    $book.SaveAs([System.IO.Path]::GetFullPath($OutputPath), 52)

    $vbProject = $book.VBProject
    $components = $vbProject.VBComponents

    for ($index = $components.Count; $index -ge 1; $index--) {
        $component = $components.Item($index)
        if ($component.Name -eq "SMWCommunity") {
            $components.Remove($component)
        }
        Release-ComObjectSafely $component
    }

    [void]$components.Import([System.IO.Path]::GetFullPath($BasPath))

    $thisWorkbookComponent = $components.Item("ThisWorkbook")
    $codeModule = $thisWorkbookComponent.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    $codeModule.AddFromString((Get-Content -LiteralPath $ThisWorkbookCodePath -Raw))

    $macroName = "'" + $book.Name + "'!SMWCommunity.PrepareCommunityWorkbook"
    [void]$excel.Run($macroName)

    $book.Save()
    $book.Close($true)
    $excel.Quit()

    Release-ComObjectSafely $codeModule
    Release-ComObjectSafely $thisWorkbookComponent
    Release-ComObjectSafely $components
    Release-ComObjectSafely $vbProject
    Release-ComObjectSafely $book
    Release-ComObjectSafely $excel
    $codeModule = $null
    $thisWorkbookComponent = $null
    $components = $null
    $vbProject = $null
    $book = $null
    $excel = $null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    try { Unblock-File -LiteralPath $OutputPath } catch {}

    Write-Host ""
    Write-Host "Build complete:" -ForegroundColor Green
    Write-Host $OutputPath -ForegroundColor Green
    Write-Host ""
    Write-Host "Upload this XLSM as the Community Edition release." -ForegroundColor Cyan
    Write-Host ""
    Start-Process -FilePath $OutputPath
}
catch {
    $message = $_.Exception.Message
    try { if ($null -ne $book) { $book.Close($false) } } catch {}
    try { if ($null -ne $excel) { $excel.Quit() } } catch {}

    Release-ComObjectSafely $codeModule
    Release-ComObjectSafely $thisWorkbookComponent
    Release-ComObjectSafely $components
    Release-ComObjectSafely $vbProject
    Release-ComObjectSafely $book
    Release-ComObjectSafely $excel
    Release-ComObjectSafely $probe

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $fullMessage = "The Community Tracker build failed:`r`n`r`n" + $message + `
        "`r`n`r`nIf the message mentions programmatic access, open Excel > File > Options > " + `
        "Trust Center > Trust Center Settings > Macro Settings, enable 'Trust access to the VBA project object model', close Excel, and rerun the builder."
    Stop-WithMessage $fullMessage
}
finally {
    if ($null -ne $registryPath) {
        try {
            if ($hadPreviousAccessVbom) {
                New-ItemProperty -LiteralPath $registryPath -Name AccessVBOM -PropertyType DWord -Value ([int]$previousAccessVbom) -Force | Out-Null
            } else {
                Remove-ItemProperty -LiteralPath $registryPath -Name AccessVBOM -ErrorAction SilentlyContinue
            }
        } catch {}
    }
}
