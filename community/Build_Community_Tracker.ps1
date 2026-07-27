[CmdletBinding()]
param(
    [string]$RepoSlug = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TemplatePath = Join-Path $ScriptDir "SMW_ROM_Hack_Tracker_Community.xlsx"
$MacroPath = Join-Path $ScriptDir "SMWCommunity.bas"
$VersionPath = Join-Path (Split-Path -Parent $ScriptDir) "data\version.json"
$OutputPath = Join-Path $ScriptDir "SMW_ROM_Hack_Tracker_Community.xlsm"

if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "Template workbook not found: $TemplatePath"
}
if (-not (Test-Path -LiteralPath $MacroPath)) {
    throw "VBA module not found: $MacroPath"
}

if ([string]::IsNullOrWhiteSpace($RepoSlug)) {
    Write-Host ""
    Write-Host "Optional: enter your GitHub repo as owner/repository, for example:" -ForegroundColor Cyan
    Write-Host "FredDOGG23/smwc_tracker" -ForegroundColor Cyan
    Write-Host "Leave blank to keep the URLs already stored in the workbook." -ForegroundColor Cyan
    $RepoSlug = Read-Host "GitHub repo"
}

$excel = $null
$book = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $book = $excel.Workbooks.Open($TemplatePath)

    if (-not [string]::IsNullOrWhiteSpace($RepoSlug)) {
        $repo = $RepoSlug.Trim().Trim("/")
        $rawBase = "https://raw.githubusercontent.com/$repo/main"
        $settings = $book.Worksheets.Item("Settings")
        $settings.Range("B2").Value2 = "$rawBase/data/SMWCentral_All_Moderated_Hacks.csv"
        $settings.Range("B3").Value2 = "$rawBase/data/version.json"
        $settings.Range("B8").Value2 = "https://github.com/$repo"
    } else {
        $settings = $book.Worksheets.Item("Settings")
    }

    $settings.Range("A10").Value2 = "Catalog Sequence"
    $settings.Range("A11").Value2 = "Refresh Mode"
    $settings.Range("B11").Value2 = "Incremental"

    if (Test-Path -LiteralPath $VersionPath) {
        try {
            $manifest = Get-Content -LiteralPath $VersionPath -Raw | ConvertFrom-Json
            $settings.Range("B10").Value2 = [int]$manifest.sequence
            if ($null -ne $manifest.catalog_version) { $settings.Range("B6").Value2 = [string]$manifest.catalog_version }
            if ($null -ne $manifest.hack_count) { $settings.Range("B5").Value2 = [int]$manifest.hack_count }
        } catch {
            $settings.Range("B10").Value2 = 0
        }
    } else {
        $settings.Range("B10").Value2 = 0
    }

    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }

    # 52 = xlOpenXMLWorkbookMacroEnabled
    $book.SaveAs($OutputPath, 52)

    try {
        [void]$book.VBProject.VBComponents.Import($MacroPath)
    } catch {
        throw "Could not import VBA module. In Excel, enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model. Then close Excel and run this builder again."
    }

    $excel.Run("UpgradeCommunityWorkbook")
    $excel.Run("InstallButtons")
    $excel.Run("PrepareWorkbookForRelease")

    $book.Save()
    $book.Close($true)
    $book = $null

    Write-Host ""
    Write-Host "Created:" -ForegroundColor Green
    Write-Host $OutputPath -ForegroundColor Green

    try { Start-Process -FilePath $OutputPath } catch {}
}
finally {
    if ($null -ne $book) {
        try { $book.Close($false) } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
