<#
Xu ly MUC 3 trong CR - "Bo sung cac ma chung khoan cu vao 020004". Dung LAI Y HET
luong duyet 2 cong (TT chung -> Chung khoan) nhu Process-NewSecurities.ps1 (muc 1),
chi khac:
  - Nguon du lieu: hang doi ton dong (old_codes_backlog.json) do
    Initialize-OldCodesBacklog.ps1 tao ra, KHONG phai NewCodes cua lan quet hien tai.
  - Chi xu ly TOI DA "BatchSize" ma MOI LAN CHAY (theo config/old-codes.config.json),
    de tranh don qua nhieu ma cu vao hang cho duyet cung luc.
  - Luon chay SAU Process-NewSecurities.ps1 trong 1 lan quet (dung thu tu CR yeu cau).
#>

param(
    [string]$FlexStorePath = (Join-Path $PSScriptRoot "..\data\flex-mock\flex_store.json"),
    [string]$EmailConfigPath = (Join-Path $PSScriptRoot "..\config\email.config.json"),
    [string]$BacklogPath = (Join-Path $PSScriptRoot "..\data\flex-mock\old_codes_backlog.json"),
    [string]$ConfigPath  = (Join-Path $PSScriptRoot "..\config\old-codes.config.json"),
    [switch]$AutoApprove,
    [switch]$SkipEmail
)

$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "Modules\FlexMockStore.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "Modules\VsdDetail.psm1") -Force

if (-not (Test-Path $BacklogPath)) {
    Write-Host "Chua co hang doi ma cu. Chay Initialize-OldCodesBacklog.ps1 truoc (chi can 1 lan khi trien khai)."
    return
}
if (-not (Test-Path $ConfigPath)) {
    throw "Khong tim thay $ConfigPath (so luong xu ly moi lan). Chay Initialize-OldCodesBacklog.ps1 truoc."
}
$cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$batchSize = [int]$cfg.BatchSize

$backlog = [object[]](Get-Content $BacklogPath -Raw -Encoding UTF8 | ConvertFrom-Json)
if ($backlog.Count -eq 0) {
    Write-Host "Hang doi ma cu da het (0 ma con lai)."
    return
}

$flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)
$existingCodes = @($flex | ForEach-Object { $_.Code })

# Bo qua ma da co tren Flex (VD: nhan vien da tu tay them truoc do) - van coi la "da xu ly", rut khoi hang doi
$stillMissing = $backlog | Where-Object { $_.Code -notin $existingCodes }
$batch = @($stillMissing | Select-Object -First $batchSize)
$handledCodes = $backlog | Where-Object { $_.Code -in $existingCodes -or $_.Code -in ($batch | ForEach-Object { $_.Code }) }
$remaining = @($backlog | Where-Object { $_.Code -notin ($handledCodes | ForEach-Object { $_.Code }) })

Write-Host "Hang doi ma cu con lai truoc khi xu ly: $($backlog.Count) ma. Xu ly lo nay: $($batch.Count) ma (BatchSize=$batchSize)."

if ($batch.Count -eq 0) {
    Write-Host "Khong co ma cu can nhap trong lo nay (co the da duoc them tren Flex qua duong khac)."
    $remaining | ConvertTo-Json -Depth 5 | Out-File -FilePath $BacklogPath -Encoding utf8
    return
}

function Send-WorkflowEmail {
    param([string]$Subject, [string]$HtmlBody)
    if ($SkipEmail) { Write-Host "[SkipEmail] (khong gui) Subject: $Subject"; return }
    if (-not (Test-Path $EmailConfigPath)) { Write-Warning "Khong co email.config.json, bo qua gui mail: $Subject"; return }
    $mailCfg = Get-Content $EmailConfigPath -Raw | ConvertFrom-Json
    $securePassword = ConvertTo-SecureString $mailCfg.Password -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($mailCfg.From, $securePassword)
    Send-MailMessage -SmtpServer $mailCfg.SmtpServer -Port $mailCfg.Port -UseSsl `
        -From $mailCfg.From -To $mailCfg.To -Subject $Subject -Body $HtmlBody -BodyAsHtml `
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

function Get-NoiQuanLyVSDText {
    param([string]$ManagementArea)
    if ($ManagementArea -like "*Chi nhánh*") { return "Trung tâm lưu ký chứng khoán Việt Nam - Chi nhánh TP Hồ Chí Minh" }
    return "Trung tâm lưu ký chứng khoán Việt Nam"
}

Write-Host "=== Muc 3 - Cong 1 (Tab TT chung): nhap $($batch.Count) ma CU ==="
foreach ($code in $batch) {
    $detail = $null
    if ($code.DetailId) {
        Write-Host "  ... dang lay chi tiet $($code.Code) tu /s-detail/$($code.DetailId)"
        $detail = Get-VsdSecurityDetail -DetailId $code.DetailId
    }

    $ttChung = [pscustomobject]@{
        MaCK                = $code.Code
        TenTCPH             = if ($detail -and $detail.TenTCPH) { $detail.TenTCPH } else { $code.Name }
        TenChungKhoan       = if ($detail -and $detail.TenChungKhoan) { $detail.TenChungKhoan } else { $code.Name }
        TenGiaoDichTiengAnh = "Đang bổ sung"
        TenTiengAnh         = "Đang bổ sung"
        ThiTruong           = "Trong nước"
        NoiQuanLyVSD        = Get-NoiQuanLyVSDText -ManagementArea $code.ManagementArea
        MaISIN              = $code.IsinCode
    }
    $now = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $newItem = [pscustomobject]@{
        Code            = $code.Code
        IsinCode        = $code.IsinCode
        Name            = $code.Name
        StockType       = $code.StockType
        Market          = $code.Market
        ManagementArea  = $code.ManagementArea
        DetailId        = $code.DetailId
        MenhGiaVSD      = if ($detail) { $detail.MenhGia } else { $null }
        TongSoDangKyVSD = if ($detail) { $detail.TongSoDangKy } else { $null }
        LoaiTraiPhieuVSD = if ($detail) { $detail.LoaiChungKhoanVSD } else { $null }
        LoaiKyHanVSD    = if ($detail) { $detail.LoaiKyHan } else { $null }
        KyHanVSD        = if ($detail) { $detail.KyHan } else { $null }
        MaCKCS_VSD      = if ($detail) { $detail.MaCKCS } else { $null }
        TenTCPHCKCS_VSD = if ($detail) { $detail.TenTCPHCKCS } else { $null }
        LoaiChungQuyenVSD = if ($detail) { $detail.LoaiChungQuyen } else { $null }
        PhuongThucThanhToanVSD = if ($detail) { $detail.PhuongThucThanhToan } else { $null }
        GiaThucHienVSD  = if ($detail) { $detail.GiaThucHien } else { $null }
        TyLeChuyenDoiVSD = if ($detail) { $detail.TyLeChuyenDoi } else { $null }
        ThoiHanCWThangVSD = if ($detail) { $detail.ThoiHanCWThang } else { $null }
        NgayDaoHanVSD   = if ($detail) { $detail.NgayDaoHan } else { $null }
        Status          = "Chờ duyệt TT chung"
        Source          = "Mã cũ (tồn đọng)"
        StatusChangedAt = $now
        LichSuDuyet     = @([pscustomobject]@{ ThoiGian = $now; HanhDong = "Nhập tab TT chung (mã cũ tồn đọng, Tên TCPH lấy từ trang chi tiết VSD)"; TrangThai = "Chờ duyệt TT chung" })
        Tabs            = [pscustomobject]@{
            TTChung    = $ttChung
            ChungKhoan = $null
            TTCK       = $null
        }
    }
    $flex.Add($newItem)
    Write-Host "  + $($code.Code) - TCPH: $($ttChung.TenTCPH)"
}
Save-FlexStore -Path $FlexStorePath -Data $flex
$remaining | ConvertTo-Json -Depth 5 | Out-File -FilePath $BacklogPath -Encoding utf8
Write-Host "Con lai trong hang doi sau lo nay: $($remaining.Count) ma."

$ttChungDisplay = $flex | Where-Object { $_.Code -in ($batch | ForEach-Object { $_.Code }) } | ForEach-Object {
    [pscustomobject]@{ Code = $_.Code; TenTCPH = $_.Tabs.TTChung.TenTCPH; ThiTruong = $_.Tabs.TTChung.ThiTruong; NoiQuanLyVSD = $_.Tabs.TTChung.NoiQuanLyVSD; MaISIN = $_.Tabs.TTChung.MaISIN }
}
Send-WorkflowEmail -Subject "[Flex] Yêu cầu duyệt tab TT chung - MÃ CŨ TỒN ĐỌNG ($($batch.Count) mã)" `
    -HtmlBody "<p>Phần mềm đã thực hiện: lấy $($batch.Count) mã từ hàng đợi mã chứng khoán cũ tồn đọng (còn lại $($remaining.Count) mã trong hàng đợi), đã tự động nhập vào tab TT chung:</p>$(Format-Table-Html -Items $ttChungDisplay -Columns @('Code','TenTCPH','ThiTruong','NoiQuanLyVSD','MaISIN'))<p>Đề nghị nhân viên kiểm tra và bấm Duyệt để chuyển sang bước nhập tab Chứng khoán.</p>"

if (-not $AutoApprove) {
    Write-Host "Dang o tab 'TT chung' - cho duyet. Chay lai voi -AutoApprove de mo phong tiep."
    return
}

Write-Host ""
Write-Host "=== [MOC PHONG DUYET TT CHUNG] -> chuyen sang tab Chung khoan ==="
$codesSet = $batch | ForEach-Object { $_.Code }
foreach ($item in $flex) {
    if ($item.Code -in $codesSet) {
        $item.Tabs.ChungKhoan = New-ChungKhoanTabData -Code $item.Code -Market $item.Market -StockType $item.StockType -MenhGiaVSD $item.MenhGiaVSD -LoaiTraiPhieuVSD $item.LoaiTraiPhieuVSD -LoaiKyHan $item.LoaiKyHanVSD -KyHan $item.KyHanVSD -MaCKCS $item.MaCKCS_VSD -TenTCPHCKCS $item.TenTCPHCKCS_VSD -LoaiChungQuyen $item.LoaiChungQuyenVSD -PhuongThucThanhToan $item.PhuongThucThanhToanVSD -GiaThucHien $item.GiaThucHienVSD -TyLeChuyenDoi $item.TyLeChuyenDoiVSD -ThoiHanCWThang $item.ThoiHanCWThangVSD -NgayDaoHan $item.NgayDaoHanVSD
        Set-FlexStatus -Item $item -NewStatus "Chờ duyệt Chứng khoán" -HanhDong "Duyệt TT chung xong (mã cũ) -> chuyển sang tab Chứng khoán"
    }
}
Save-FlexStore -Path $FlexStorePath -Data $flex

$chungKhoanDisplay = $flex | Where-Object { $_.Code -in $codesSet } | ForEach-Object {
    [pscustomobject]@{ Code = $_.Code; NoiGD = $_.Tabs.ChungKhoan.NoiGD; LoaiChungKhoan = $_.Tabs.ChungKhoan.LoaiChungKhoan; MenhGia = $_.Tabs.ChungKhoan.MenhGia }
}
Send-WorkflowEmail -Subject "[Flex] Yêu cầu duyệt tab Chứng khoán - MÃ CŨ TỒN ĐỌNG ($($batch.Count) mã)" `
    -HtmlBody "<p>Phần mềm đã thực hiện: sau khi tab TT chung được duyệt, đã nhập tiếp tab Chứng khoán cho các mã cũ sau:</p>$(Format-Table-Html -Items $chungKhoanDisplay -Columns @('Code','NoiGD','LoaiChungKhoan','MenhGia'))<p>Đề nghị nhân viên kiểm tra và bấm Duyệt để hoàn tất.</p>"

if (-not $AutoApprove) { return }

Write-Host "=== [MOC PHONG DUYET CHUNG KHOAN] -> hoan tat khai ma cu ==="
foreach ($item in $flex) {
    if ($item.Code -in $codesSet) {
        $item.Tabs.TTCK = [pscustomobject]@{
            DonViGiaoDich         = "1000"
            KhoiLuongNiemYet      = if ($item.TongSoDangKyVSD) { $item.TongSoDangKyVSD } else { "N/A (theo tổng số đăng ký tại VSD)" }
            NgayNiemYet           = (Get-Date -Format "yyyy-MM-dd")
            MuaBanCungNgay        = "Có"
            CheckRoomNDTNuocNgoai = "Có"
        }
        Set-FlexStatus -Item $item -NewStatus "Hoạt động" -HanhDong "Duyet Chung khoan xong (ma cu) - hoan tat"
    }
}
Save-FlexStore -Path $FlexStorePath -Data $flex
Write-Host "Da hoan tat khai $($batch.Count) ma cu (Status = Hoat dong). Con lai $($remaining.Count) ma trong hang doi cho lan quet sau."
