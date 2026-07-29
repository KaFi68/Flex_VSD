<#
Script tong (orchestrator) - CHAY DINH KY (Task Scheduler) hoac thu cong, dung cho ca
DEMO lan mo phong quy trinh that theo dung 4 muc CR:

  1. Quet VSD (Fetch-VsdSecurities.ps1)
  2. So sanh voi Flex (mock) (Compare-VsdWithFlex.ps1)
  3. Xu ly ma moi - muc 1 CR (Process-NewSecurities.ps1 -AutoApprove)
  3b. Xu ly ma cu ton dong - muc 3 CR (Process-OldSecurities.ps1 -AutoApprove), chay SAU muc 1
  4. Xu ly chuyen san - muc 2 CR (Process-MarketTransfer.ps1)
  4b. Xu ly doi ten - muc 4 CR (Process-NameChanges.ps1 -AutoApprove)
  5. Gui mail tong hop thanh cong (muc 5.1 CR)

LUU Y: -AutoApprove la MO PHONG hanh dong "nhan vien bam duyet" de demo/tu dong chay
het mot luot. Trong Flex that day PHAI la hanh dong thu cong cua con nguoi tren man
hinh Flex - khi noi Flex that, bo cac co -AutoApprove va de nhan vien tu duyet qua
UI mock (hoac Flex that).
#>

param(
    [int]$FetchMaxPages = 30,      # gioi han trang quet VSD cho nhanh (demo). San xuat that: de 0 (quet full)
    [switch]$SkipFetch,            # bo qua buoc 1, dung snapshot VSD moi nhat da co san
    [switch]$SkipEmail,            # khong gui mail that (chi in ra console) - dung khi test nhanh
    [switch]$AutoApprove           # CHI dung de DEMO tu duyet het 1 luot. San xuat that: KHONG bat co nay,
                                    # de nhan vien tu duyet qua UI mock (hoac Flex that) nhu quy trinh CR yeu cau.
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$emailCfg  = Join-Path $scriptDir "..\config\email.config.json"

function Send-SummaryEmail {
    param($Report, $NewCount, $MarketCount, $NameCount)
    if ($SkipEmail) { Write-Host "[SkipEmail] Bo qua gui mail tong hop."; return }
    if (-not (Test-Path $emailCfg)) { Write-Warning "Chua co email.config.json, bo qua mail tong hop."; return }
    $cfg = Get-Content $emailCfg -Raw | ConvertFrom-Json
    $securePassword = ConvertTo-SecureString $cfg.Password -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($cfg.From, $securePassword)

    $body = @"
<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px">
<p><b>Báo cáo tổng hợp - quét VSD và cập nhật Flex (mock demo)</b><br/>Thời gian: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
<ul>
  <li>Khai mã chứng khoán mới: <b>$NewCount</b> mã (xem chi tiết trong 2 mail duyệt tiến trình 1/2 trước đó)</li>
  <li>Sửa mã chuyển sàn: <b>$MarketCount</b> mã</li>
  <li>Sửa mã cần chỉnh tên: <b>$NameCount</b> mã</li>
</ul>
<p style="color:#888">Đây là mail tổng hợp tương ứng mục 5.1 trong CR gửi phòng Lưu ký. Dữ liệu Flex đang là MOCK (demo), chưa nối thật với hệ thống Flex.</p>
</body></html>
"@
    Send-MailMessage -SmtpServer $cfg.SmtpServer -Port $cfg.Port -UseSsl `
        -From $cfg.From -To $cfg.To -Subject "[TEST] VSD-Flex - Báo cáo tổng hợp tự động hóa" `
        -Body $body -BodyAsHtml -Credential $cred -Encoding ([System.Text.Encoding]::UTF8)
    Write-Host "Da gui mail tong hop."
}

Write-Host "############################################"
Write-Host "# BUOC 1: Quet du lieu VSD"
Write-Host "############################################"
if ($SkipFetch) {
    Write-Host "(Bo qua - dung snapshot VSD moi nhat da co)"
} else {
    & (Join-Path $scriptDir "Fetch-VsdSecurities.ps1") -MaxPages $FetchMaxPages
}

Write-Host ""
Write-Host "############################################"
Write-Host "# BUOC 2: So sanh VSD voi Flex (mock)"
Write-Host "############################################"
$reportFile = & (Join-Path $scriptDir "Compare-VsdWithFlex.ps1")
$reportFile = ($reportFile | Select-Object -Last 1).ToString().Trim()
Write-Host "Report file: $reportFile"

$report = Get-Content $reportFile -Raw | ConvertFrom-Json
$newCount    = @($report.NewCodes).Count
$marketCount = @($report.MarketChanges).Count
$nameCount   = @($report.NameChanges).Count

Write-Host ""
Write-Host "############################################"
Write-Host "# BUOC 3: Xu ly ma moi (muc 1 CR)"
Write-Host "############################################"
if ($newCount -gt 0) {
    & (Join-Path $scriptDir "Process-NewSecurities.ps1") -ReportFile $reportFile -AutoApprove:$AutoApprove -SkipEmail:$SkipEmail
} else {
    Write-Host "Khong co ma moi."
}

Write-Host ""
Write-Host "############################################"
Write-Host "# BUOC 3b: Xu ly ma cu ton dong (muc 3 CR) - chay SAU ma moi"
Write-Host "############################################"
$oldCodesConfig = Join-Path $scriptDir "..\config\old-codes.config.json"
if (Test-Path $oldCodesConfig) {
    & (Join-Path $scriptDir "Process-OldSecurities.ps1") -AutoApprove:$AutoApprove -SkipEmail:$SkipEmail
} else {
    Write-Host "Chua khoi tao hang doi ma cu (chay Initialize-OldCodesBacklog.ps1 truoc neu can). Bo qua buoc nay."
}

Write-Host ""
Write-Host "############################################"
Write-Host "# BUOC 4: Xu ly chuyen san (muc 2 CR)"
Write-Host "############################################"
if ($marketCount -gt 0) {
    & (Join-Path $scriptDir "Process-MarketTransfer.ps1") -ReportFile $reportFile -SkipEmail:$SkipEmail
} else {
    Write-Host "Khong co ma can chuyen san."
}

Write-Host ""
Write-Host "############################################"
Write-Host "# BUOC 4b: Xu ly doi ten (muc 4 CR)"
Write-Host "############################################"
if ($nameCount -gt 0) {
    & (Join-Path $scriptDir "Process-NameChanges.ps1") -ReportFile $reportFile -AutoApprove:$AutoApprove -SkipEmail:$SkipEmail
} else {
    Write-Host "Khong co ma can doi ten."
}

Write-Host ""
Write-Host "############################################"
Write-Host "# BUOC 5: Gui mail tong hop"
Write-Host "############################################"
Send-SummaryEmail -Report $report -NewCount $newCount -MarketCount $marketCount -NameCount $nameCount

Write-Host ""
Write-Host "=== HOAN TAT MOT LUOT CHAY DEMO ==="
Write-Host "Ma moi: $newCount | Chuyen san: $marketCount | Doi ten (chua xu ly tu dong): $nameCount"
