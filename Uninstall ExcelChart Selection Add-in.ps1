param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$addinFileName = 'ExcelChart Selection.xlam'
$addinTitle = 'ExcelChart Selection'

$runningExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($runningExcel.Count -gt 0 -and -not $Force) {
    throw 'Close every Microsoft Excel window, then run this uninstaller again.'
}

$roamingData = [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)
$addInDirectory = Join-Path $roamingData 'Microsoft\AddIns'
$destination = Join-Path $addInDirectory $addinFileName

$excel = $null
$addIn = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    try {
        $addIn = $excel.AddIns.Item($addinTitle)
    }
    catch {
        $addIn = $null
    }

    if ($null -ne $addIn) {
        $addIn.Installed = $false
        if ($addIn.Installed) {
            throw 'Excel did not disable the registered add-in.'
        }
    }
}
finally {
    if ($null -ne $addIn) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addIn)
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (Test-Path -LiteralPath $destination -PathType Leaf) {
    Remove-Item -LiteralPath $destination -Force
}

if (Test-Path -LiteralPath $addInDirectory -PathType Container) {
    Get-ChildItem -LiteralPath $addInDirectory -File -Filter "$addinFileName.backup_*" |
        Remove-Item -Force
}

$staleRegistration = $false
$optionsPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
if (Test-Path -LiteralPath $optionsPath) {
    $options = Get-ItemProperty -LiteralPath $optionsPath
    foreach ($property in $options.PSObject.Properties) {
        if ($property.Name -like 'OPEN*' -and
            [string]$property.Value -match '(?i)ExcelChart Selection\.xlam') {
            $staleRegistration = $true
        }
    }
}

if ($staleRegistration) {
    throw 'The deployed file was removed, but Excel left a startup registration behind. Open Excel Options > Add-ins > Manage Excel Add-ins and clear ExcelChart Selection.'
}

Write-Host 'Uninstallation complete.'
Write-Host 'The package folder and its source add-in were left intact.'
