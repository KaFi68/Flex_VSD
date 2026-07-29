<#
Xu ly MUC 1 trong CR - "Khai ma chung khoan moi". Theo dung mo hinh nguoi dung xac
nhan: day la 2 CONG DUYET RIENG, moi cong ung voi 1 TAB/TRANG danh sach tren man
hinh Flex (020004):

  Cong 1 - Tab "TT chung": nhap Ma CK, Ten TCPH, Ten chung khoan, Thi truong, Noi
           quan ly VSD, Ma ISIN -> gui mail cho nhan vien -> Status = "Cho duyet
           TT chung". Nhan vien duyet xong tren man hinh -> ma chuyen sang tab
           "Chung khoan".
  Cong 2 - Tab "Chung khoan": nhap Noi GD, Loai chung khoan, Menh gia -> gui mail
           -> Status = "Cho duyet Chung khoan". Nhan vien duyet xong -> he thong
           tu dien tab TTCK (khong hien thi rieng tren UI danh sach) -> Status =
           "Hoat dong" (hoan tat).

Theo dung bang mapping trong file "Cach thuc khai ma chung khoan" (Excel dinh kem
CR):
  Flex "Ten viet tat"          <- VSD "Ma chung khoan"          (da co: Code)
  Flex "To chuc phat hanh"     <- VSD "Ten TCPH"                (RIENG BIET voi ten CK!)
  Flex "Ten giao dich"         <- VSD "Ten chung khoan"
  Flex "Ten giao dich TA"      <- VSD "Ten chung khoan (TA)"    (VSD thuong CHUA CO, cham hon)
  Flex "Ten TA"                <- VSD "Ten TCPH (TA)"           (VSD thuong CHUA CO, cham hon)
  Flex "Thi truong"            <- luon chon "Trong nuoc"
  Flex "Noi quan ly tai VSD"   <- tuy VSD hien "chi nhanh" hay "tru so chinh"
  Flex "Ma ISIN"               <- VSD "Ma ISIN"

"Ten TCPH" KHONG co o trang danh sach VSD (chi co "Ten chung khoan") - phai fetch
THEM trang chi tiet tung ma (/s-detail/{id}) qua Modules/VsdDetail.psm1 de lay dung.
Trang chi tiet nay cung co Menh gia va Tong so dang ky (dung cho tab Chung khoan/TTCK).

Vi CHUA CO Flex that, buoc "nhan vien bam duyet" duoc mo phong bang tham so
-AutoApprove (chi dung de DEMO chay het 1 luot, KHONG dai dien cho hanh vi thuc te
- trong Flex that day la hanh dong thu cong cua nhan vien tren man hinh).
#>

param(
    [Parameter(Mandatory)][string]$ReportFile,     # file flex_diff_*.json tu Compare-VsdWithFlex.ps1
    [string]$FlexStorePath = (Join-Path $PSScriptRoot "..\data\flex-mock\flex_store.json"),
    [string]$EmailConfigPath = (Join-Path $PSScriptRoot "..\config\email.config.json"),
    [switch]$AutoApprove,
    [switch]$SkipEmail,   # dung khi test khong muon gui mail that
    [string]$Source = "Mã mới"   # hien thi tren UI ("Nguon") - Process-OldSecurities.ps1 dat "Mã cũ (tồn đọng)"
)

$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "Modules\FlexMockStore.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "Modules\VsdDetail.psm1") -Force

function Send-WorkflowEmail {
    param([string]$Subject, [string]$HtmlBody)
    if ($SkipEmail) {
        Write-Host "[SkipEmail] (khong gui) Subject: $Subject"
        return
    }
    if (-not (Test-Path $EmailConfigPath)) {
        Write-Warning "Khong co email.config.json, bo qua gui mail: $Subject"
        return
    }
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

function Get-NoiQuanLyVSDText {
    param([string]$ManagementArea)
    if ($ManagementArea -like "*Chi nhánh*") { return "Trung tâm lưu ký chứng khoán Việt Nam - Chi nhánh TP Hồ Chí Minh" }
    return "Trung tâm lưu ký chứng khoán Việt Nam"
}

$report = Get-Content $ReportFile -Raw | ConvertFrom-Json
$newCodes = @($report.NewCodes)

if ($newCodes.Count -eq 0) {
    Write-Host "Khong co ma moi can xu ly."
    return
}

$flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)
$existingCodes = @($flex | ForEach-Object { $_.Code })

# Chong trung: neu chay lai script voi cung report (VD: chay 2 lan lien tiep truoc khi
# quet lai VSD), bo qua ma da co san tren Flex thay vi tao ban ghi trung lap.
$skippedExisting = $newCodes | Where-Object { $_.Code -in $existingCodes }
$newCodes = @($newCodes | Where-Object { $_.Code -notin $existingCodes })
if ($skippedExisting) {
    Write-Warning "Bo qua $($skippedExisting.Count) ma da co san tren Flex (tranh trung lap): $(($skippedExisting | ForEach-Object { $_.Code }) -join ', ')"
}
if ($newCodes.Count -eq 0) {
    Write-Host "Khong con ma moi nao can xu ly sau khi loc trung."
    return
}

Write-Host "=== Cong 1 - Tab TT chung: nhap $($newCodes.Count) ma moi ==="
foreach ($code in $newCodes) {
    $detail = $null
    if ($code.DetailId) {
        Write-Host "  ... dang lay chi tiet $($code.Code) tu /s-detail/$($code.DetailId)"
        $detail = Get-VsdSecurityDetail -DetailId $code.DetailId
    }

    $ttChung = [pscustomobject]@{
        MaCK                = $code.Code
        TenTCPH             = if ($detail -and $detail.TenTCPH) { $detail.TenTCPH } else { $code.Name }
        TenChungKhoan       = if ($detail -and $detail.TenChungKhoan) { $detail.TenChungKhoan } else { $code.Name }
        TenGiaoDichTiengAnh = "Đang bổ sung"   # VSD thuong cap nhat tieng Anh cham hon tieng Viet
        TenTiengAnh         = "Đang bổ sung"
        ThiTruong           = "Trong nước"
        NoiQuanLyVSD        = Get-NoiQuanLyVSDText -ManagementArea $code.ManagementArea
        MaISIN              = $code.IsinCode
    }

    $now = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $newItem = [pscustomobject]@{
        Code           = $code.Code
        IsinCode       = $code.IsinCode
        Name           = $code.Name
        StockType      = $code.StockType
        Market         = $code.Market
        ManagementArea = $code.ManagementArea
        DetailId       = $code.DetailId
        MenhGiaVSD     = if ($detail) { $detail.MenhGia } else { $null }        # luu lai de dung cho tab Chung khoan
        TongSoDangKyVSD = if ($detail) { $detail.TongSoDangKy } else { $null }   # luu lai de dung cho tab TTCK
        LoaiTraiPhieuVSD = if ($detail) { $detail.LoaiChungKhoanVSD } else { $null }  # chi dung cho Trai phieu
        LoaiKyHanVSD   = if ($detail) { $detail.LoaiKyHan } else { $null }       # chi dung cho Trai phieu
        KyHanVSD       = if ($detail) { $detail.KyHan } else { $null }           # chi dung cho Trai phieu
        MaCKCS_VSD     = if ($detail) { $detail.MaCKCS } else { $null }               # chi dung cho Chung quyen
        TenTCPHCKCS_VSD = if ($detail) { $detail.TenTCPHCKCS } else { $null }          # chi dung cho Chung quyen
        LoaiChungQuyenVSD = if ($detail) { $detail.LoaiChungQuyen } else { $null }     # chi dung cho Chung quyen
        PhuongThucThanhToanVSD = if ($detail) { $detail.PhuongThucThanhToan } else { $null }  # chi dung cho Chung quyen
        GiaThucHienVSD = if ($detail) { $detail.GiaThucHien } else { $null }           # chi dung cho Chung quyen
        TyLeChuyenDoiVSD = if ($detail) { $detail.TyLeChuyenDoi } else { $null }       # chi dung cho Chung quyen
        ThoiHanCWThangVSD = if ($detail) { $detail.ThoiHanCWThang } else { $null }     # chi dung cho Chung quyen
        NgayDaoHanVSD  = if ($detail) { $detail.NgayDaoHan } else { $null }            # chi dung cho Chung quyen
        Status         = "Chờ duyệt TT chung"
        Source         = $Source
        StatusChangedAt = $now
        LichSuDuyet    = @([pscustomobject]@{ ThoiGian = $now; HanhDong = "Nhap tab TT chung (Ten TCPH lay tu trang chi tiet VSD)"; TrangThai = "Chờ duyệt TT chung" })
        Tabs           = [pscustomobject]@{
            TTChung     = $ttChung
            ChungKhoan  = $null
            TTCK        = $null
        }
    }
    $flex.Add($newItem)
    Write-Host "  + $($code.Code) - TCPH: $($ttChung.TenTCPH)"
}
Save-FlexStore -Path $FlexStorePath -Data $flex

$ttChungDisplay = $flex | Where-Object { $_.Code -in ($newCodes | ForEach-Object { $_.Code }) } | ForEach-Object {
    [pscustomobject]@{
        Code         = $_.Code
        TenTCPH      = $_.Tabs.TTChung.TenTCPH
        ThiTruong    = $_.Tabs.TTChung.ThiTruong
        NoiQuanLyVSD = $_.Tabs.TTChung.NoiQuanLyVSD
        MaISIN       = $_.Tabs.TTChung.MaISIN
    }
}
Send-WorkflowEmail -Subject "[Flex] Yeu cau duyet tab TT chung ($($newCodes.Count) ma moi)" `
    -HtmlBody "<p>Phan mem da thuc hien: kiem tra tren VSD phat hien $($newCodes.Count) ma chung khoan MOI chua co tren Flex, va da tu dong nhap cac thong tin sau vao tab TT chung (man hinh 020004):</p>$(Format-Table-Html -Items $ttChungDisplay -Columns @('Code','TenTCPH','ThiTruong','NoiQuanLyVSD','MaISIN'))<p>De nghi nhan vien kiem tra va bam Duyet de chuyen sang buoc nhap tab Chung khoan.</p>"

if (-not $AutoApprove) {
    Write-Host ""
    Write-Host "Dang o tab 'TT chung' - cho duyet. Nhan vien can duyet tren Flex that (hoac bam Duyet tren UI mock), sau do chay lai voi -AutoApprove de mo phong tiep."
    return
}

# ---- Mo phong: nhan vien da duyet tab TT chung -> chuyen sang tab Chung khoan ----
Write-Host ""
Write-Host "=== [MOC PHONG DUYET TT CHUNG] -> chuyen sang tab Chung khoan ==="
$codesSet = $newCodes | ForEach-Object { $_.Code }
foreach ($item in $flex) {
    if ($item.Code -in $codesSet) {
        $item.Tabs.ChungKhoan = New-ChungKhoanTabData -Code $item.Code -Market $item.Market -StockType $item.StockType -MenhGiaVSD $item.MenhGiaVSD -LoaiTraiPhieuVSD $item.LoaiTraiPhieuVSD -LoaiKyHan $item.LoaiKyHanVSD -KyHan $item.KyHanVSD -MaCKCS $item.MaCKCS_VSD -TenTCPHCKCS $item.TenTCPHCKCS_VSD -LoaiChungQuyen $item.LoaiChungQuyenVSD -PhuongThucThanhToan $item.PhuongThucThanhToanVSD -GiaThucHien $item.GiaThucHienVSD -TyLeChuyenDoi $item.TyLeChuyenDoiVSD -ThoiHanCWThang $item.ThoiHanCWThangVSD -NgayDaoHan $item.NgayDaoHanVSD
        Set-FlexStatus -Item $item -NewStatus "Chờ duyệt Chứng khoán" -HanhDong "Duyet TT chung xong -> chuyen sang tab Chung khoan"
    }
}
Save-FlexStore -Path $FlexStorePath -Data $flex

$chungKhoanDisplay = $flex | Where-Object { $_.Code -in $codesSet } | ForEach-Object {
    [pscustomobject]@{
        Code           = $_.Code
        NoiGD          = $_.Tabs.ChungKhoan.NoiGD
        LoaiChungKhoan = $_.Tabs.ChungKhoan.LoaiChungKhoan
        MenhGia        = $_.Tabs.ChungKhoan.MenhGia
    }
}
Send-WorkflowEmail -Subject "[Flex] Yeu cau duyet tab Chung khoan ($($newCodes.Count) ma moi)" `
    -HtmlBody "<p>Phan mem da thuc hien: sau khi tab TT chung duoc duyet, da tu dong nhap tiep cac thong tin sau vao tab Chung khoan:</p>$(Format-Table-Html -Items $chungKhoanDisplay -Columns @('Code','NoiGD','LoaiChungKhoan','MenhGia'))<p>De nghi nhan vien kiem tra va bam Duyet de hoan tat khai ma moi.</p>"

if (-not $AutoApprove) { return }

# ---- Mo phong: nhan vien da duyet tab Chung khoan -> hoan tat (dien TTCK ngam) ----
Write-Host "=== [MOC PHONG DUYET CHUNG KHOAN] -> hoan tat khai ma moi ==="
foreach ($item in $flex) {
    if ($item.Code -in $codesSet) {
        $item.Tabs.TTCK = [pscustomobject]@{
            DonViGiaoDich         = "1000"
            KhoiLuongNiemYet      = if ($item.TongSoDangKyVSD) { $item.TongSoDangKyVSD } else { "N/A (theo tổng số đăng ký tại VSD)" }
            NgayNiemYet           = (Get-Date -Format "yyyy-MM-dd")
            MuaBanCungNgay        = "Có"
            CheckRoomNDTNuocNgoai = "Có"
        }
        Set-FlexStatus -Item $item -NewStatus "Hoạt động" -HanhDong "Duyet Chung khoan xong - hoan tat"
    }
}
Save-FlexStore -Path $FlexStorePath -Data $flex
Write-Host "Da hoan tat khai $($newCodes.Count) ma moi (Status = Hoat dong)."
