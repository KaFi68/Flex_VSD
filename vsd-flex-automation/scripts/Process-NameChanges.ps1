<#
Xu ly MUC 4 trong CR - "Chinh sua ten to chuc phat hanh / ten giao dich". Ma nay
DA CO SAN tren Flex (dang Hoat dong) - chi sua 4 truong ten tren tab TT chung, gui
mail cho nhan vien DUYET 1 LAN (khac muc 1/3 phai duyet 2 lan TT chung + Chung khoan).

Rule dien ten (viet tat khi ten qua dai...) theo file "Cach thuc khai ma chung
khoan" - CHUA lam trong demo nay, hien dang lay nguyen ten tu VSD.
Ten tieng Anh (Ten giao dich tieng Anh / Ten tieng Anh): VSD scraper hien CHUA lay
duoc truong nay (trang tra cuu VSD khong hien thi) nen de placeholder can nhap tay.
#>

param(
    [Parameter(Mandatory)][string]$ReportFile,
    [string]$FlexStorePath = (Join-Path $PSScriptRoot "..\data\flex-mock\flex_store.json"),
    [string]$EmailConfigPath = (Join-Path $PSScriptRoot "..\config\email.config.json"),
    [switch]$AutoApprove,
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
$nameChanges = @($report.NameChanges)

if ($nameChanges.Count -eq 0) {
    Write-Host "Khong co ma nao can doi ten."
    return
}

$flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)
$flexByCode = @{}
foreach ($item in $flex) { $flexByCode[$item.Code] = $item }

Write-Host "=== Muc 4 - Buoc 1-2: sua ten TCPH/ten giao dich cho $($nameChanges.Count) ma ==="
foreach ($chg in $nameChanges) {
    if (-not $flexByCode.ContainsKey($chg.Code)) { continue }
    $item = $flexByCode[$chg.Code]

    $item | Add-Member -NotePropertyName TraCuuDoiTen -NotePropertyValue ([pscustomobject]@{
        TenCu  = $chg.FlexName
        TenMoi = $chg.VsdName
        TraCuuLuc = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }) -Force

    $item.Name = $chg.VsdName
    # Ma nay DA hoat dong san tren Flex - neu chua tung co "Tabs" (VD: ma seed ban dau,
    # chua qua Process-NewSecurities.ps1) thi phai dien du lieu THAT cho ChungKhoan
    # (khong de null), de UI/approve-handler nhan dung day la ma cu, duyet 1 lan la xong.
    if (-not $item.Tabs) { $item | Add-Member -NotePropertyName Tabs -NotePropertyValue ([pscustomobject]@{ TTChung=$null; ChungKhoan=$null; TTCK=$null }) -Force }
    if (-not $item.Tabs.ChungKhoan) {
        $item.Tabs.ChungKhoan = [pscustomobject]@{
            MaCK = $item.Code; NoiGD = $item.Market; LoaiChungKhoan = $item.StockType; MenhGia = "10000"
        }
    }
    $oldTTChung = $item.Tabs.TTChung
    $item.Tabs.TTChung = [pscustomobject]@{
        MaCK                = $item.Code
        TenTCPH             = $chg.VsdName
        TenChungKhoan       = $chg.VsdName
        ThiTruong           = if ($oldTTChung) { $oldTTChung.ThiTruong } else { "Trong nước" }
        NoiQuanLyVSD        = if ($oldTTChung) { $oldTTChung.NoiQuanLyVSD } else { "Trung tâm lưu ký chứng khoán Việt Nam" }
        MaISIN              = if ($oldTTChung) { $oldTTChung.MaISIN } else { $item.IsinCode }
        TenGiaoDichTiengAnh = if ($oldTTChung -and $oldTTChung.TenGiaoDichTiengAnh -and $oldTTChung.TenGiaoDichTiengAnh -ne "Đang bổ sung") { $oldTTChung.TenGiaoDichTiengAnh } else { "Đang bổ sung" }
        TenTiengAnh         = if ($oldTTChung -and $oldTTChung.TenTiengAnh -and $oldTTChung.TenTiengAnh -ne "Đang bổ sung") { $oldTTChung.TenTiengAnh } else { "Đang bổ sung" }
    }

    Set-FlexStatus -Item $item -NewStatus "Chờ duyệt TT chung" -HanhDong "Sửa tên TCPH/tên giao dịch theo VSD ('$($chg.FlexName)' -> '$($chg.VsdName)'), chờ duyệt 1 lần"
    Write-Host "  $($chg.Code): '$($chg.FlexName)' -> '$($chg.VsdName)'"
}
Save-FlexStore -Path $FlexStorePath -Data $flex

Send-WorkflowEmail -Subject "[Flex] Yêu cầu duyệt ĐỔI TÊN ($($nameChanges.Count) mã)" `
    -HtmlBody "<p>Phần mềm đã thực hiện: phát hiện $($nameChanges.Count) mã có tên khác giữa VSD và Flex, đã tự động sửa tab TT chung như sau:</p>$(Format-Table-Html -Items $nameChanges -Columns @('Code','FlexName','VsdName'))<p>Đề nghị nhân viên kiểm tra và bấm Duyệt (chỉ 1 lần, vì mã đã hoạt động sẵn trên Flex).</p>"

if (-not $AutoApprove) {
    Write-Host "Dang cho duyet (1 lan) tai tab TT chung. Chay lai voi -AutoApprove de mo phong tiep, hoac bam Duyet tren UI mock."
    return
}

Write-Host "=== [MOC PHONG DUYET] -> hoan tat doi ten (ma da hoat dong, khong qua tab Chung khoan) ==="
foreach ($chg in $nameChanges) {
    if ($flexByCode.ContainsKey($chg.Code)) {
        Set-FlexStatus -Item $flexByCode[$chg.Code] -NewStatus "Hoạt động" -HanhDong "Duyệt đổi tên xong - hoàn tất"
    }
}
Save-FlexStore -Path $FlexStorePath -Data $flex

Send-WorkflowEmail -Subject "[Flex] Đã duyệt xong ĐỔI TÊN ($($nameChanges.Count) mã)" `
    -HtmlBody "<p>Các mã sau đã được duyệt xong việc đổi tên:</p>$(Format-Table-Html -Items $nameChanges -Columns @('Code','VsdName'))"

Write-Host "Da hoan tat doi ten cho $($nameChanges.Count) ma."
