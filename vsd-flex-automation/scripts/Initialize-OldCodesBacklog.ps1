<#
Mo phong "Phan mem cho phep nhap so luong ma chung khoan cu khi mo phan mem" (muc 3
trong CR). Chay MOT LAN khi trien khai lan dau (hoac khi muon nap lai hang doi ton
dong): luu toan bo danh sach ma "chua co tren Flex" tai thoi diem nay thanh HANG DOI
"ma cu", va luu "so luong xu ly moi lan" vao file config.

Sau buoc nay, moi lan quet VSD tiep theo (Compare-VsdWithFlex.ps1) chi tao ra "ma
moi" that su (phat sinh sau thoi diem khoi tao nay) - phan da nam trong hang doi
"ma cu" se duoc Process-OldSecurities.ps1 xu ly dan theo lo, KHONG lap lai o muc 1.
#>

param(
    [Parameter(Mandatory)][string]$ReportFile,   # file flex_diff_*.json tu Compare-VsdWithFlex.ps1 (lay NewCodes hien tai lam ton dong)
    [Parameter(Mandatory)][int]$BatchSize,        # so luong ma cu xu ly moi lan quet - nhan vien nhap khi mo phan mem
    [string]$BacklogPath = (Join-Path $PSScriptRoot "..\data\flex-mock\old_codes_backlog.json"),
    [string]$ConfigPath  = (Join-Path $PSScriptRoot "..\config\old-codes.config.json")
)

$ErrorActionPreference = "Stop"

$report = Get-Content $ReportFile -Raw | ConvertFrom-Json
$newCodes = @($report.NewCodes)

$outDir = Split-Path $BacklogPath -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$newCodes | ConvertTo-Json -Depth 5 | Out-File -FilePath $BacklogPath -Encoding utf8

$cfgDir = Split-Path $ConfigPath -Parent
if (-not (Test-Path $cfgDir)) { New-Item -ItemType Directory -Path $cfgDir -Force | Out-Null }
[pscustomobject]@{ BatchSize = $BatchSize } | ConvertTo-Json | Out-File -FilePath $ConfigPath -Encoding utf8

Write-Host "Da khoi tao hang doi ma cu: $($newCodes.Count) ma (tu $($report.VsdSnapshot))"
Write-Host "So luong xu ly moi lan quet: $BatchSize ma"
Write-Host "Hang doi luu tai: $BacklogPath"
