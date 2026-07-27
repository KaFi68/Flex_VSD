<#
Xoa VINH VIEN cac ma da nam trong tab "Da xoa" (soft-deleted qua man hinh tim kiem
cua Start-FlexMockUI.ps1) qua han giu (mac dinh 30 ngay). Truoc khi qua han, nhan
vien co the "Khoi phuc" (Restore) tai tab "Da xoa" tren UI.

Nen chay dinh ky (VD: 1 lan/ngay) qua Task Scheduler de dam bao viec xoa vinh vien
xay ra dung han, khong phu thuoc vao viec co ai mo trang UI hay khong.

Test nhanh: chay voi -RetentionDays 0 se xoa vinh vien tat ca ma dang trong tab
"Da xoa" ngay lap tuc (dung de kiem tra logic, KHONG dung gia tri nay trong san xuat).
#>

param(
    [int]$RetentionDays = 30,
    [string]$FlexStorePath = (Join-Path $PSScriptRoot "..\data\flex-mock\flex_store.json")
)

$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "Modules\FlexMockStore.psm1") -Force

$flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)
$now = Get-Date
$toPurge = New-Object System.Collections.Generic.List[object]

foreach ($item in $flex) {
    if ($item.Status -ne "Đã xóa" -or -not $item.DeletedAt) { continue }
    try {
        $deletedAt = [datetime]::ParseExact($item.DeletedAt, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
        if (($now - $deletedAt).TotalDays -ge $RetentionDays) {
            $toPurge.Add($item)
        }
    } catch {
        Write-Warning "Bo qua ma $($item.Code): DeletedAt khong dung dinh dang ($($item.DeletedAt))"
    }
}

if ($toPurge.Count -eq 0) {
    Write-Host "Khong co ma nao qua han $RetentionDays ngay trong tab Da xoa."
    return
}

Write-Host "Xoa vinh vien $($toPurge.Count) ma (qua han $RetentionDays ngay):"
foreach ($item in $toPurge) {
    Write-Host "  - $($item.Code) - $($item.Name) (da xoa luc $($item.DeletedAt))"
    $flex.Remove($item) | Out-Null
}

Save-FlexStore -Path $FlexStorePath -Data $flex
Write-Host "Da xoa vinh vien xong."
