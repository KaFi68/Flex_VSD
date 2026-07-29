<#
Xu ly MUC 2 trong CR - "Kiem tra ma chung khoan chuyen san":
  Buoc 1-3: cap nhat "Noi giao dich" trong Flex (mock) theo gia tri moi tren VSD
  Buoc 4: gui mail cho nhan vien duyet ngay sau khi thuc hien xong

CR con mo ta mail 5.2 "gui truoc 1 ngay lam viec so voi ngay hieu luc chuyen san"
- phan nay CAN NGAY HIEU LUC CHUYEN SAN tu VSD (hien khong co trong du lieu list
  dang scrape) nen CHUA lam trong ban demo nay. Ghi chu TODO o day.
#>

param(
    [Parameter(Mandatory)][string]$ReportFile,
    [string]$FlexStorePath = (Join-Path $PSScriptRoot "..\data\flex-mock\flex_store.json"),
    [string]$EmailConfigPath = (Join-Path $PSScriptRoot "..\config\email.config.json"),
    [switch]$SkipEmail
)

$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "Modules\FlexMockStore.psm1") -Force

function Send-WorkflowEmail {
    param([string]$Subject, [string]$HtmlBody)
    if ($SkipEmail) { Write-Host "[SkipEmail] (khong gui) Subject: $Subject"; return }
    if (-not (Test-Path $EmailConfigPath)) { Write-Warning "Khong co email.config.json, bo qua gui mail: $Subject"; return }
    $cfg = Get-Content $EmailConfigPath -Raw | ConvertFrom-Json
    $securePassword = ConvertTo-SecureString $cfg.Password -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($cfg.From, $securePassword)
    Send-MailMessage -SmtpServer $cfg.SmtpServer -Port $cfg.Port -UseSsl `
        -From $cfg.From -To $cfg.To -Subject $Subject -Body $HtmlBody -BodyAsHtml `
        -Credential $cred -Encoding ([System.Text.Encoding]::UTF8)
    Write-Host "Da gui mail: $Subject"
}

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

$report = Get-Content $ReportFile -Raw | ConvertFrom-Json
$marketChanges = @($report.MarketChanges)

if ($marketChanges.Count -eq 0) {
    Write-Host "Khong co ma nao can chuyen san."
    return
}

$flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)
$flexByCode = @{}
foreach ($item in $flex) { $flexByCode[$item.Code] = $item }

Write-Host "=== Buoc 1-2 (tra cuu): doi chieu Noi GD tren Flex vs VSD cho $($marketChanges.Count) ma ==="
foreach ($chg in $marketChanges) {
    if ($flexByCode.ContainsKey($chg.Code)) {
        $item = $flexByCode[$chg.Code]

        # Buoc 1-2 CR: "Nhap giao dich 020004 -> Nhap ma CK can chuyen san -> vao tab
        # Chung khoan -> chon Chung khoan co Noi giao dich la San cu" - luu lai ket qua
        # tra cuu nay truoc khi sua, de hien thi ro rang tren UI (khong chi ghi de).
        $item | Add-Member -NotePropertyName TraCuuChuyenSan -NotePropertyValue ([pscustomobject]@{
            NoiGDCu  = $item.Market
            NoiGDMoi = $chg.VsdMarket
            TraCuuLuc = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }) -Force

        # Buoc 3: sua truong Noi GD sang san moi -> vao thang tab "Chung khoan"
        # (khong qua tab TT chung vi ma nay da co tren Flex tu truoc, chi sua 1 truong)
        $item.Market = $chg.VsdMarket
        Set-FlexStatus -Item $item -NewStatus "Chờ duyệt Chứng khoán" -HanhDong "Tra cứu thấy đổi Nơi GD ($($chg.FlexMarket) -> $($chg.VsdMarket)), đã sửa trên Flex"
        Write-Host "  $($chg.Code): tra cuu thay Noi GD Flex=$($chg.FlexMarket) khac VSD=$($chg.VsdMarket) -> da sua sang $($chg.VsdMarket)"
    }
}
Save-FlexStore -Path $FlexStorePath -Data $flex

Send-WorkflowEmail -Subject "[Flex] Yêu cầu duyệt giao dịch CHUYỂN SÀN ($($marketChanges.Count) mã)" `
    -HtmlBody "<p>Phần mềm đã thực hiện: tra cứu trên VSD phát hiện $($marketChanges.Count) mã có Nơi giao dịch khác với Flex, đã tự động sửa trường Nơi GD trên tab Chứng khoán như sau (cột 'FlexMarket' là giá trị CŨ trước khi sửa):</p>$(Format-Table-Html -Items $marketChanges -Columns @('Code','Name','FlexMarket','VsdMarket'))<p>Đề nghị nhân viên kiểm tra và bấm Duyệt để xác nhận thay đổi.</p>"

Write-Host "Da xu ly xong buoc chuyen san (dang Cho duyet). TODO: mail bao truoc 1 ngay (muc 5.2 CR) can them du lieu ngay hieu luc tu VSD."
