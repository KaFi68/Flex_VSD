<#
Gui email tom tat bao cao diff moi nhat (data/reports/diff_*.json) toi nguoi nhan cau hinh
trong config/email.config.json.

CACH DUNG:
  1. Sua file config/email.config.json (xem huong dan trong Send-VsdDiffReport.README.txt
     cung thu muc) - dien SmtpServer, Port, From, Password (App Password), To.
  2. Chay:  powershell -File Send-VsdDiffReport.ps1
     (mac dinh gui bao cao diff moi nhat trong data/reports)
#>

param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "..\config\email.config.json"),
    [string]$ReportDir  = (Join-Path $PSScriptRoot "..\data\reports"),
    [string]$ReportFile = ""   # de trong = tu dong lay bao cao moi nhat
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $ConfigPath)) {
    throw "Khong tim thay file config: $ConfigPath. Hay tao file nay tu email.config.json.example va dien thong tin SMTP truoc."
}
$cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json

if ([string]::IsNullOrWhiteSpace($ReportFile)) {
    $latest = Get-ChildItem -Path $ReportDir -Filter "diff_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) { throw "Khong co bao cao diff nao trong $ReportDir. Chay Fetch-VsdSecurities.ps1 va Compare-VsdSnapshots.ps1 truoc." }
    $ReportFile = $latest.FullName
}

Write-Host "Dang doc bao cao: $ReportFile"
$report = Get-Content $ReportFile -Raw | ConvertFrom-Json

function Format-Section {
    param([string]$Title, $Items, [string[]]$Columns)
    if (-not $Items -or $Items.Count -eq 0) {
        return "<h3>$Title (0)</h3><p><i>Khong co thay doi.</i></p>"
    }
    $rows = foreach ($item in $Items) {
        $cells = foreach ($c in $Columns) { "<td style='padding:4px 8px;border:1px solid #ddd'>$([System.Net.WebUtility]::HtmlEncode([string]$item.$c))</td>" }
        "<tr>$($cells -join '')</tr>"
    }
    $header = ($Columns | ForEach-Object { "<th style='padding:4px 8px;border:1px solid #ddd;background:#f0f0f0'>$_</th>" }) -join ''
    return "<h3>$Title ($($Items.Count))</h3><table style='border-collapse:collapse'><tr>$header</tr>$($rows -join '')</table>"
}

$body = @"
<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px">
<p>Bao cao quet ma chung khoan VSD - so sanh <b>$($report.PreviousFile)</b> voi <b>$($report.LatestFile)</b> (luc $($report.ComparedAt)).</p>
$(Format-Section -Title "Ma MOI xuat hien tren VSD" -Items $report.NewCodes -Columns @('Code','Name','Market','Status'))
$(Format-Section -Title "Doi Noi giao dich (ung vien CHUYEN SAN)" -Items $report.MarketChanges -Columns @('Code','OldMarket','NewMarket','Name'))
$(Format-Section -Title "Doi Ten" -Items $report.NameChanges -Columns @('Code','OldName','NewName'))
$(Format-Section -Title "Doi Trang thai" -Items $report.StatusChanges -Columns @('Code','OldStatus','NewStatus','Name'))
<p style="color:#888">Mail tu dong tu VSD-Flex automation (dang trong giai doan test).</p>
</body></html>
"@

$securePassword = ConvertTo-SecureString $cfg.Password -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($cfg.From, $securePassword)

$mailParams = @{
    SmtpServer = $cfg.SmtpServer
    Port       = $cfg.Port
    UseSsl     = $true
    From       = $cfg.From
    To         = $cfg.To
    Subject    = "[TEST] VSD - Bao cao thay doi ma chung khoan - $($report.ComparedAt)"
    Body       = $body
    BodyAsHtml = $true
    Credential = $cred
    Encoding   = [System.Text.Encoding]::UTF8
}

Write-Host "Dang gui mail toi $($cfg.To) qua $($cfg.SmtpServer):$($cfg.Port) ..."
Send-MailMessage @mailParams
Write-Host "Da gui xong."
