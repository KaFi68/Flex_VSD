<#
Quet toan bo Flex (mock) tim cac ma dang "Cho duyet*" nhung da qua ThresholdMinutes
(mac dinh 60 phut = 1 tieng) ma van chua duoc duyet buoc tiep theo. Gui mail nhac
nho nhan vien vao duyet, ap dung cho MOI buoc dang cho duyet (khong chi rieng giua
lan 1 va lan 2).

Trong thuc te nen dat script nay chay dinh ky (VD moi 15-30 phut) qua Task
Scheduler de kiem tra lien tuc, khong chi chay 3 lan/ngay nhu buoc quet VSD chinh.
#>

param(
    [int]$ThresholdMinutes = 60,
    [string]$FlexStorePath = (Join-Path $PSScriptRoot "..\data\flex-mock\flex_store.json"),
    [string]$EmailConfigPath = (Join-Path $PSScriptRoot "..\config\email.config.json"),
    [switch]$SkipEmail
)

$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "Modules\FlexMockStore.psm1") -Force

function Format-Table-Html {
    param($Items, [string[]]$Columns)
    if (-not $Items -or $Items.Count -eq 0) { return "<p><i>Khong co ma nao.</i></p>" }
    $header = ($Columns | ForEach-Object { "<th style='padding:4px 8px;border:1px solid #ddd;background:#f0f0f0'>$_</th>" }) -join ''
    $rows = foreach ($item in $Items) {
        $cells = foreach ($c in $Columns) { "<td style='padding:4px 8px;border:1px solid #ddd'>$([System.Net.WebUtility]::HtmlEncode([string]$item.$c))</td>" }
        "<tr>$($cells -join '')</tr>"
    }
    return "<table style='border-collapse:collapse'><tr>$header</tr>$($rows -join '')</table>"
}

$flex = Get-FlexStore -Path $FlexStorePath
$now = Get-Date
$overdue = New-Object System.Collections.Generic.List[object]

foreach ($item in $flex) {
    if ($item.Status -notlike "Chờ duyệt*") { continue }
    if (-not $item.StatusChangedAt) { continue }
    try {
        $changedAt = [datetime]::ParseExact($item.StatusChangedAt, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        Write-Warning "Bo qua ma $($item.Code): StatusChangedAt khong dung dinh dang ($($item.StatusChangedAt))"
        continue
    }
    $minutesWaiting = ($now - $changedAt).TotalMinutes
    if ($minutesWaiting -ge $ThresholdMinutes) {
        $overdue.Add([pscustomobject]@{
            Code           = $item.Code
            Name           = $item.Name
            Status         = $item.Status
            ChoDuyetTu     = $item.StatusChangedAt
            SoPhutDaCho    = [math]::Round($minutesWaiting, 0)
        })
    }
}

if ($overdue.Count -eq 0) {
    Write-Host "Khong co ma nao cho duyet qua $ThresholdMinutes phut."
    return
}

Write-Host "Co $($overdue.Count) ma cho duyet qua $ThresholdMinutes phut:"
$overdue | Format-Table -AutoSize | Out-String | Write-Host

if ($SkipEmail) {
    Write-Host "[SkipEmail] Khong gui mail nhac nho."
    return
}
if (-not (Test-Path $EmailConfigPath)) {
    Write-Warning "Khong co email.config.json, bo qua gui mail nhac nho."
    return
}

$cfg = Get-Content $EmailConfigPath -Raw | ConvertFrom-Json
$securePassword = ConvertTo-SecureString $cfg.Password -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($cfg.From, $securePassword)

$body = @"
<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px">
<p><b>Nhac nho: co $($overdue.Count) ma chung khoan dang cho duyet qua $ThresholdMinutes phut ma chua duoc xu ly.</b></p>
$(Format-Table-Html -Items $overdue -Columns @('Code','Name','Status','ChoDuyetTu','SoPhutDaCho'))
<p>Vui long vao Flex kiem tra va duyet som de tranh cham tien do khai bao chung khoan.</p>
</body></html>
"@

Send-MailMessage -SmtpServer $cfg.SmtpServer -Port $cfg.Port -UseSsl `
    -From $cfg.From -To $cfg.To -Subject "[Flex] Nhac nho: $($overdue.Count) ma dang cho duyet qua $ThresholdMinutes phut" `
    -Body $body -BodyAsHtml -Credential $cred -Encoding ([System.Text.Encoding]::UTF8)

Write-Host "Da gui mail nhac nho."
