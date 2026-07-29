<#
Tao du lieu Flex GIA LAP tu mot snapshot VSD that, de demo toan bo luong nghiep vu
(mac du chua ket noi duoc Flex that). Kich ban gia lap:
  - Lay N ma dau tien tu snapshot VSD lam "du lieu hien co trong Flex" (Status = Hoat dong).
  - BO BOT vai ma cuoi danh sach  -> gia lap "ma chua co tren Flex" (se bi phat hien la MOI).
  - Doi "Market" cua vai ma       -> gia lap "can chuyen san".
  - Doi "Name" cua 1 ma           -> gia lap "can doi ten".

Day la DU LIEU DEMO, khong phai du lieu Flex that.
#>

param(
    [string]$SnapshotDir = (Join-Path $PSScriptRoot "..\data\snapshots"),
    [string]$OutFile     = (Join-Path $PSScriptRoot "..\data\flex-mock\flex_store.json"),
    [int]$MissingCount   = 5,   # so ma se "thieu" tren Flex (=> bao la MOI)
    [int]$MarketChangeCount = 2,
    [int]$NameChangeCount   = 1
)

$ErrorActionPreference = "Stop"

$latestSnapshot = Get-ChildItem -Path $SnapshotDir -Filter "vsd_securities_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $latestSnapshot) { throw "Chua co snapshot VSD nao. Chay Fetch-VsdSecurities.ps1 truoc." }

Write-Host "Dung snapshot: $($latestSnapshot.Name)"
$vsd = [object[]](Get-Content $latestSnapshot.FullName -Raw -Encoding UTF8 | ConvertFrom-Json)
Write-Host "Tong so ma trong snapshot: $($vsd.Count)"

if ($vsd.Count -le ($MissingCount + $MarketChangeCount + $NameChangeCount + 5)) {
    throw "Snapshot qua nho ($($vsd.Count) ma) de tao du lieu demo co y nghia. Hay chay Fetch-VsdSecurities.ps1 voi -MaxPages lon hon (vi du 30)."
}

# Sao chep sang mang Flex mock, bo cac ma CUOI danh sach de gia lap "chua co tren Flex"
$seedTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
$flexList = $vsd[0..($vsd.Count - 1 - $MissingCount)] | ForEach-Object {
    [pscustomobject]@{
        Code           = $_.Code
        IsinCode       = $_.IsinCode
        Name           = $_.Name
        StockType      = $_.StockType
        Market         = $_.Market
        ManagementArea = $_.ManagementArea
        Status         = "Hoạt động"   # trang thai binh thuong trong Flex
        StatusChangedAt = $seedTime
        LichSuDuyet    = @([pscustomobject]@{ ThoiGian = $seedTime; HanhDong = "Khởi tạo dữ liệu demo"; TrangThai = "Hoạt động" })
    }
}

$missingCodes = $vsd[($vsd.Count - $MissingCount)..($vsd.Count - 1)] | ForEach-Object { $_.Code }
Write-Host "Ma se bi bao la MOI (khong co trong Flex mock): $($missingCodes -join ', ')"

# Gia lap can CHUYEN SAN: doi Market cua vai ma dau danh sach
$oldMarkets = @{}
for ($i = 0; $i -lt $MarketChangeCount; $i++) {
    $item = $flexList[$i]
    $oldMarkets[$item.Code] = $item.Market
    $item.Market = if ($item.Market -eq "HOSE") { "HNX" } elseif ($item.Market -eq "HNX") { "HOSE" } else { "HOSE" }
}
Write-Host "Ma se bi bao can CHUYEN SAN (Flex dang sai san so voi VSD): $($oldMarkets.Keys -join ', ')"

# Gia lap can DOI TEN: doi Name cua 1 ma (offset qua khoi phan da dung cho market change)
$nameChangeIdx = $MarketChangeCount
$oldNames = @{}
for ($i = 0; $i -lt $NameChangeCount; $i++) {
    $item = $flexList[$nameChangeIdx + $i]
    $oldNames[$item.Code] = $item.Name
    $item.Name = $item.Name + " (ten cu - Flex chua cap nhat)"
}
Write-Host "Ma se bi bao can DOI TEN (Flex dang luu ten cu): $($oldNames.Keys -join ', ')"

$outDir = Split-Path $OutFile -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$flexList | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutFile -Encoding utf8

Write-Host ""
Write-Host "Da tao Flex mock: $OutFile ($($flexList.Count) ma)"
