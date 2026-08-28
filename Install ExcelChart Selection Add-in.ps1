param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$addinFileName = 'ExcelChart Selection.xlam'
$addinTitle = 'ExcelChart Selection'
$source = Join-Path $PSScriptRoot $addinFileName

if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
    throw "$addinFileName was not found beside this installer."
}

$runningExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($runningExcel.Count -gt 0 -and -not $Force) {
    throw 'Close every Microsoft Excel window, then run this installer again. Excel must restart to load application-level chart context menus.'
}

$roamingData = [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)
$addInDirectory = Join-Path $roamingData 'Microsoft\AddIns'
[IO.Directory]::CreateDirectory($addInDirectory) | Out-Null
$destination = Join-Path $addInDirectory $addinFileName

if (Test-Path -LiteralPath $destination -PathType Leaf) {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    Copy-Item -LiteralPath $destination -Destination ($destination + '.backup_' + $stamp)
}
Copy-Item -LiteralPath $source -Destination $destination -Force

$excel = $null
$addIn = $null
$previousSecurity = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $previousSecurity = $excel.AutomationSecurity
    $excel.AutomationSecurity = 1

    try {
        $addIn = $excel.AddIns.Item($addinTitle)
    }
    catch {
        $addIn = $null
    }

    if ($null -ne $addIn -and
        -not [string]::Equals($addIn.FullName, $destination, [StringComparison]::OrdinalIgnoreCase)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addIn)
        $addIn = $null
    }

    if ($null -eq $addIn) {
        $addIn = $excel.AddIns.Add($destination, $false)
    }

    $addIn.Installed = $true
    if (-not $addIn.Installed) {
        throw 'Excel did not mark the add-in as installed.'
    }

    Write-Host "Installed: $($addIn.FullName)"
}
finally {
    if ($null -ne $addIn) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addIn)
    }
    if ($null -ne $excel) {
        if ($null -ne $previousSecurity) {
            $excel.AutomationSecurity = $previousSecurity
        }
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ''
Write-Host 'Installation complete.'
Write-Host 'Start Microsoft Excel normally. Do not also open the development .xlsm file.'
