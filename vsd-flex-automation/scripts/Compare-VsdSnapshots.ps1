<#
So sanh snapshot moi nhat voi snapshot truoc do de phat hien:
  - Ma chung khoan MOI xuat hien tren VSD
  - Ma da doi "Market" (Noi giao dich)   -> ung vien CHUYEN SAN
  - Ma da doi "Name" (ten to chuc/ten CK) -> ung vien DOI TEN
  - Ma bien mat khoi danh sach (hiem, dang canh bao de kiem tra thu cong)

LUU Y QUAN TRONG: script nay hien dang so sanh VSD voi chinh VSD lan quet truoc
(tam thoi, vi chua co ket noi doc du lieu tu Flex). Day CHUA PHAI logic cuoi cung
theo CR - CR yeu cau so sanh VSD voi du lieu HIEN CO TRONG FLEX de biet ma nao
"chua co tren Flex". Khi co duoc cach doc du liem Flex (export/API/DB), thay the
ham Get-PreviousSnapshotData ben duoi bang ham doc du lieu Flex.
#>

param(
    [string]$SnapshotDir = (Join-Path $PSScriptRoot "..\data\snapshots"),
    [string]$ReportDir   = (Join-Path $PSScriptRoot "..\data\reports")
)

$ErrorActionPreference = "Stop"

function Get-LatestSnapshots {
    param([string]$Dir, [int]$Count = 2)
    Get-ChildItem -Path $Dir -Filter "vsd_securities_*.json" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First $Count
}

$files = Get-LatestSnapshots -Dir $SnapshotDir -Count 2
if ($files.Count -lt 2) {
    Write-Warning "Can it nhat 2 snapshot de so sanh. Hien co $($files.Count). Chay Fetch-VsdSecurities.ps1 them mot lan nua vao lan sau roi so sanh."
    return
}

$latestFile   = $files[0]
$previousFile = $files[1]

Write-Host "So sanh:"
Write-Host "  Truoc : $($previousFile.Name)"
Write-Host "  Moi   : $($latestFile.Name)"

$latest   = Get-Content $latestFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
$previous = Get-Content $previousFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json

$latestByCode   = @{}
$previousByCode = @{}
foreach ($item in $latest)   { $latestByCode[$item.Code]   = $item }
foreach ($item in $previous) { $previousByCode[$item.Code] = $item }

$newCodes     = New-Object System.Collections.Generic.List[object]
$marketChange = New-Object System.Collections.Generic.List[object]
$nameChange   = New-Object System.Collections.Generic.List[object]
$statusChange = New-Object System.Collections.Generic.List[object]
$removedCodes = New-Object System.Collections.Generic.List[object]

foreach ($code in $latestByCode.Keys) {
    $cur = $latestByCode[$code]
    if (-not $previousByCode.ContainsKey($code)) {
        $newCodes.Add($cur)
        continue
    }
    $old = $previousByCode[$code]
    if ($old.Market -ne $cur.Market) {
        $marketChange.Add([pscustomobject]@{ Code = $code; OldMarket = $old.Market; NewMarket = $cur.Market; Name = $cur.Name })
    }
    if ($old.Name -ne $cur.Name) {
        $nameChange.Add([pscustomobject]@{ Code = $code; OldName = $old.Name; NewName = $cur.Name })
    }
    if ($old.Status -ne $cur.Status) {
        $statusChange.Add([pscustomobject]@{ Code = $code; OldStatus = $old.Status; NewStatus = $cur.Status; Name = $cur.Name })
    }
}

foreach ($code in $previousByCode.Keys) {
    if (-not $latestByCode.ContainsKey($code)) {
        $removedCodes.Add($previousByCode[$code])
    }
}

Write-Host ""
Write-Host "== Ma MOI (chua tung thay o lan quet truoc): $($newCodes.Count) =="
$newCodes | Format-Table Code, Name, Market, Status -AutoSize | Out-String | Write-Host

Write-Host "== Doi NOI GIAO DICH (ung vien CHUYEN SAN): $($marketChange.Count) =="
$marketChange | Format-Table Code, OldMarket, NewMarket, Name -AutoSize | Out-String | Write-Host

Write-Host "== Doi TEN: $($nameChange.Count) =="
$nameChange | Format-Table Code, OldName, NewName -AutoSize | Out-String | Write-Host

Write-Host "== Doi TRANG THAI: $($statusChange.Count) =="
$statusChange | Format-Table Code, OldStatus, NewStatus, Name -AutoSize | Out-String | Write-Host

if ($removedCodes.Count -gt 0) {
    Write-Warning "$($removedCodes.Count) ma bien mat khoi danh sach so voi lan quet truoc - kiem tra thu cong (co the do loi quet, khong nen tu dong xu ly)."
}

if (-not (Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null }
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportFile = Join-Path $ReportDir "diff_$timestamp.json"

[pscustomobject]@{
    ComparedAt    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    PreviousFile  = $previousFile.Name
    LatestFile    = $latestFile.Name
    NewCodes      = $newCodes
    MarketChanges = $marketChange
    NameChanges   = $nameChange
    StatusChanges = $statusChange
    RemovedCodes  = $removedCodes
} | ConvertTo-Json -Depth 5 | Out-File -FilePath $reportFile -Encoding utf8

Write-Host "Da luu bao cao diff: $reportFile"
