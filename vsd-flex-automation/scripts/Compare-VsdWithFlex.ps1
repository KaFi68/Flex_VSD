<#
So sanh snapshot VSD moi nhat voi du lieu HIEN CO TRONG FLEX (dang la mock, xem
Modules/FlexMockStore.psm1) - day la dung logic nghiep vu CR yeu cau:
  - Ma co tren VSD nhung CHUA co trong Flex        -> ma MOI
  - Ma co ca 2 noi nhung khac "Market"              -> can CHUYEN SAN
  - Ma co ca 2 noi nhung khac "Name"                 -> can DOI TEN
#>

param(
    [string]$SnapshotDir = (Join-Path $PSScriptRoot "..\data\snapshots"),
    [string]$FlexStorePath = (Join-Path $PSScriptRoot "..\data\flex-mock\flex_store.json"),
    [string]$ReportDir = (Join-Path $PSScriptRoot "..\data\reports")
)

$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "Modules\FlexMockStore.psm1") -Force

$latestSnapshot = Get-ChildItem -Path $SnapshotDir -Filter "vsd_securities_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $latestSnapshot) { throw "Chua co snapshot VSD nao. Chay Fetch-VsdSecurities.ps1 truoc." }

$vsd  = [object[]](Get-Content $latestSnapshot.FullName -Raw -Encoding UTF8 | ConvertFrom-Json)
$flex = Get-FlexStore -Path $FlexStorePath

Write-Host "VSD snapshot: $($latestSnapshot.Name) ($($vsd.Count) ma)"
Write-Host "Flex mock:    $FlexStorePath ($($flex.Count) ma)"

$flexByCode = @{}
foreach ($item in $flex) { $flexByCode[$item.Code] = $item }

$newCodes     = New-Object System.Collections.Generic.List[object]
$marketChange = New-Object System.Collections.Generic.List[object]
$nameChange   = New-Object System.Collections.Generic.List[object]

foreach ($v in $vsd) {
    if (-not $flexByCode.ContainsKey($v.Code)) {
        $newCodes.Add($v)
        continue
    }
    $f = $flexByCode[$v.Code]
    if ($f.Market -ne $v.Market) {
        $marketChange.Add([pscustomobject]@{ Code = $v.Code; Name = $v.Name; FlexMarket = $f.Market; VsdMarket = $v.Market })
    }
    if ($f.Name -ne $v.Name) {
        $nameChange.Add([pscustomobject]@{ Code = $v.Code; FlexName = $f.Name; VsdName = $v.Name })
    }
}

Write-Host ""
Write-Host "== Ma MOI (co tren VSD, CHUA co tren Flex): $($newCodes.Count) =="
$newCodes | Format-Table Code, Name, Market, Status -AutoSize | Out-String | Write-Host

Write-Host "== Can CHUYEN SAN (Market khac giua Flex va VSD): $($marketChange.Count) =="
$marketChange | Format-Table Code, Name, FlexMarket, VsdMarket -AutoSize | Out-String | Write-Host

Write-Host "== Can DOI TEN (Name khac giua Flex va VSD): $($nameChange.Count) =="
$nameChange | Format-Table Code, FlexName, VsdName -AutoSize | Out-String | Write-Host

if (-not (Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null }
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportFile = Join-Path $ReportDir "flex_diff_$timestamp.json"

[pscustomobject]@{
    ComparedAt    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    VsdSnapshot   = $latestSnapshot.Name
    NewCodes      = $newCodes
    MarketChanges = $marketChange
    NameChanges   = $nameChange
} | ConvertTo-Json -Depth 5 | Out-File -FilePath $reportFile -Encoding utf8

Write-Host "Da luu bao cao: $reportFile"
return $reportFile
