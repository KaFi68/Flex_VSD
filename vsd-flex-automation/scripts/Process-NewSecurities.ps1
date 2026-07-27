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

$report = Get-Content $ReportFile -Raw | ConvertFrom-Json
$newCodes = @($report.NewCodes)

if ($newCodes.Count -eq 0) {
    Write-Host "Khong co ma moi can xu ly."
    return
}

$flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)

Write-Host "=== Cong 1 - Tab TT chung: nhap $($newCodes.Count) ma moi ==="
foreach ($code in $newCodes) {
    $ttChung = [pscustomobject]@{
        MaCK              = $code.Code
        TenTCPH           = $code.Name
        TenChungKhoan     = $code.Name
        ThiTruong         = "Trong nước"
        NoiQuanLyVSD      = "Trung tâm lưu ký chứng khoán Việt Nam"
        MaISIN            = $code.IsinCode
    }

    $now = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $newItem = [pscustomobject]@{
        Code           = $code.Code
        IsinCode       = $code.IsinCode
        Name           = $code.Name
        StockType      = $code.StockType
        Market         = $code.Market
        ManagementArea = $code.ManagementArea
        Status         = "Chờ duyệt TT chung"
        Source         = $Source
        StatusChangedAt = $now
        LichSuDuyet    = @([pscustomobject]@{ ThoiGian = $now; HanhDong = "Nhap tab TT chung"; TrangThai = "Chờ duyệt TT chung" })
        Tabs           = [pscustomobject]@{
            TTChung     = $ttChung
            ChungKhoan  = $null
            TTCK        = $null
        }
    }
    $flex.Add($newItem)
    Write-Host "  + $($code.Code) - $($code.Name)"
}
Save-FlexStore -Path $FlexStorePath -Data $flex

$ttChungDisplay = $newCodes | ForEach-Object {
    [pscustomobject]@{
        Code         = $_.Code
        TenTCPH      = $_.Name
        ThiTruong    = "Trong nước"
        NoiQuanLyVSD = "Trung tâm lưu ký chứng khoán Việt Nam"
        MaISIN       = $_.IsinCode
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
        $item.Tabs.ChungKhoan = [pscustomobject]@{
            MaCK           = $item.Code
            NoiGD          = $item.Market
            LoaiChungKhoan = $item.StockType
            MenhGia        = "10000"
        }
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
            KhoiLuongNiemYet      = "N/A (theo tổng số đăng ký tại VSD)"
            NgayNiemYet           = (Get-Date -Format "yyyy-MM-dd")
            MuaBanCungNgay        = "Có"
            CheckRoomNDTNuocNgoai = "Có"
        }
        Set-FlexStatus -Item $item -NewStatus "Hoạt động" -HanhDong "Duyet Chung khoan xong - hoan tat"
    }
}
Save-FlexStore -Path $FlexStorePath -Data $flex
Write-Host "Da hoan tat khai $($newCodes.Count) ma moi (Status = Hoat dong)."
