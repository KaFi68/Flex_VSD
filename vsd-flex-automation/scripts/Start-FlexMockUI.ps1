<#
Giao dien web don gian (chay local, khong can cai them gi) mo phong man hinh 020004
cua Flex, dung Flex GIA LAP (mock) - vi CHUA noi duoc Flex that.

KIEN TRUC (theo mo ta thuc te cua nguoi dung, khac voi ban dau mình hieu nham):
  Day la CAC TAB O CAP TRANG (page-level), moi tab la 1 DANH SACH nhieu ma dang o
  dung giai doan do - KHONG PHAI mo rong rieng tung dong.

  Tab "TT chung"      : danh sach ma MOI dang cho duyet phan TT chung.
                        Duyet xong -> gui mail -> ma bien mat khoi tab nay, xuat
                        hien o tab "Chung khoan".
  Tab "Chung khoan"   : danh sach ma dang cho duyet phan Chung khoan - gom ca ma
                        moi (vua qua tu tab TT chung) VA ma CHUYEN SAN (vi Noi GD
                        la du lieu thuoc tab Chung khoan). Duyet xong -> gui mail
                        -> hoan tat (tab TTCK duoc dien tu dong ben trong, khong
                        hien thi rieng o day).
  Tab "D/S Kiem soat" : tab doc lap, hien lich su duyet GAN DAY cua TAT CA ma
                        (khong phai 1 buoc trong luong duyet).

Co man hinh dang nhap don gian (tai khoan demo trong config/auth.config.json).
LUU Y: day la co che dang nhap DEMO, KHONG dung lam mau bao mat cho he thong that.

Cach chay:
    powershell -ExecutionPolicy Bypass -File Start-FlexMockUI.ps1
    Sau do mo trinh duyet: http://localhost:8080/
#>

param(
    [int]$Port = 8080,
    [string]$FlexStorePath = (Join-Path $PSScriptRoot "..\data\flex-mock\flex_store.json"),
    [string]$EmailConfigPath = (Join-Path $PSScriptRoot "..\config\email.config.json"),
    [string]$AuthConfigPath = (Join-Path $PSScriptRoot "..\config\auth.config.json"),
    [int]$DeleteRetentionDays = 30,   # so ngay giu trong tab "Da xoa" truoc khi xoa vinh vien (xem Purge-DeletedSecurities.ps1)
    [string]$OldCodesBacklogPath = (Join-Path $PSScriptRoot "..\data\flex-mock\old_codes_backlog.json"),
    [string]$OldCodesConfigPath  = (Join-Path $PSScriptRoot "..\config\old-codes.config.json"),
    [string]$ReportsDir = (Join-Path $PSScriptRoot "..\data\reports")
)

$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "Modules\FlexMockStore.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "Modules\VsdDetail.psm1") -Force

function Get-NoiQuanLyVSDText {
    param([string]$ManagementArea)
    if ($ManagementArea -like "*Chi nhánh*") { return "Trung tâm lưu ký chứng khoán Việt Nam - Chi nhánh TP Hồ Chí Minh" }
    return "Trung tâm lưu ký chứng khoán Việt Nam"
}

if (-not (Test-Path $AuthConfigPath)) {
    throw "Khong tim thay $AuthConfigPath. Can file nay de biet tai khoan dang nhap demo."
}
$authCfg = Get-Content $AuthConfigPath -Raw | ConvertFrom-Json

# Danh sach session dang dang nhap (chi luu trong bo nho, mat khi restart server)
$global:ValidSessions = New-Object System.Collections.Generic.HashSet[string]

function Enc {
    param([string]$Text)
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Enc-Multiline {
    # Giong Enc nhung xuong dong (`n) trong noi dung se thanh <br/> that su - dung cho
    # HanhDong khi liet ke nhieu truong da sua (moi truong 1 dong cho de theo doi).
    param([string]$Text)
    return ([System.Net.WebUtility]::HtmlEncode($Text) -replace "`n", "<br/>")
}

function Send-ApproveNotice {
    param([string]$Code, [string]$NewStatus, [string]$HanhDong = "")
    if (-not (Test-Path $EmailConfigPath)) { return }
    try {
        $cfg = Get-Content $EmailConfigPath -Raw | ConvertFrom-Json
        $securePassword = ConvertTo-SecureString $cfg.Password -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($cfg.From, $securePassword)
        $hanhDongLine = if ($HanhDong) { "<p>Hệ thống đã thực hiện:<br/><b>$(Enc-Multiline $HanhDong)</b></p>" } else { "" }
        Send-MailMessage -SmtpServer $cfg.SmtpServer -Port $cfg.Port -UseSsl `
            -From $cfg.From -To $cfg.To -Subject "[Flex] Cap nhat ma $Code" `
            -Body "<p>Mã <b>$Code</b> vừa có cập nhật. Trạng thái mới: <b>$NewStatus</b></p>$hanhDongLine" -BodyAsHtml `
            -Credential $cred -Encoding ([System.Text.Encoding]::UTF8)
    } catch {
        Write-Warning "Gui mail that bai: $_"
    }
}

function Get-NextStatus {
    param([string]$CurrentStatus)
    switch -Wildcard ($CurrentStatus) {
        "Chờ duyệt TT chung*"      { return "Chờ duyệt Chứng khoán" }
        "Chờ duyệt Chứng khoán*"   { return "Hoạt động" }
        default                   { return $CurrentStatus }
    }
}

function Get-SessionToken {
    param($Request)
    $cookie = $Request.Cookies["session"]
    if ($cookie) { return $cookie.Value }
    return $null
}

function Test-Authenticated {
    param($Request)
    $token = Get-SessionToken -Request $Request
    return ($token -and $global:ValidSessions.Contains($token))
}

function Parse-FormBody {
    param($Request)
    $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    $body = $reader.ReadToEnd()
    $reader.Close()
    $result = @{}
    foreach ($pair in $body -split '&') {
        if ($pair -match '^([^=]+)=(.*)$') {
            $key = [System.Uri]::UnescapeDataString(($Matches[1] -replace '\+',' '))
            $val = [System.Uri]::UnescapeDataString(($Matches[2] -replace '\+',' '))
            $result[$key] = $val
        }
    }
    return $result
}

function Build-LoginHtml {
    param([bool]$ShowError = $false)
    $errorBox = if ($ShowError) { "<div class='err'>Sai tài khoản hoặc mật khẩu.</div>" } else { "" }
    $hintText = "Tài khoản demo: $($authCfg.Username) / $($authCfg.Password)"
    return @"
<!doctype html>
<html><head><meta charset="utf-8"><title>Đăng nhập - Flex (Mock)</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; background:#f6f7f9; display:flex; align-items:center; justify-content:center; height:100vh; margin:0 }
  .box { background:#fff; padding:32px 36px; border-radius:8px; box-shadow:0 2px 10px rgba(0,0,0,0.1); width:300px }
  h1 { font-size:18px; margin:0 0 20px }
  label { font-size:13px; color:#444; display:block; margin-bottom:4px }
  input { width:100%; padding:8px 10px; margin-bottom:14px; border:1px solid #ccc; border-radius:5px; box-sizing:border-box; font-size:14px }
  button { width:100%; background:#2563eb; color:#fff; border:none; padding:10px; border-radius:5px; cursor:pointer; font-size:14px }
  button:hover { background:#1d4ed8 }
  .err { background:#fde8e8; color:#c0392b; padding:8px 10px; border-radius:5px; font-size:13px; margin-bottom:14px }
  .hint { font-size:12px; color:#999; margin-top:14px; text-align:center }
</style></head>
<body>
  <form class="box" method="post" action="/login">
    <h1>Đăng nhập Flex (Mock demo)</h1>
    $errorBox
    <label>Tài khoản</label>
    <input type="text" name="username" autofocus />
    <label>Mật khẩu</label>
    <input type="password" name="password" />
    <button type="submit">Đăng nhập</button>
    <div class="hint">$hintText</div>
  </form>
</body></html>
"@
}

$NoiGDOptions = @("HOSE","HNX","OTC","UPCOM","WFT","DCCNY","BOND")
$NoiLuuKyOptions = @("Trụ sở chính","Chi nhánh")
$LoaiChungKhoanOptions = @("Cổ phiếu","Chứng chỉ quỹ","Tín phiếu","Trái phiếu","Chứng quyền")

function Get-SyntheticId {
    param([string]$Code, [int]$Seed)
    $sum = 0
    foreach ($ch in $Code.ToCharArray()) { $sum += [int]$ch }
    return "{0:D7}" -f (($sum * $Seed) % 9999999)
}

function Build-NameChangesTable {
    param($Items)
    if (-not $Items -or $Items.Count -eq 0) {
        return "<tr><td colspan='3' style='text-align:center;color:#888'>Không có mã nào cần cập nhật tên</td></tr>"
    }
    ($Items | ForEach-Object {
        $tc = $_.TraCuuDoiTen
        "<tr>
            <td>$(Enc $_.Code)</td>
            <td>$(Enc $tc.TenCu)</td>
            <td><b>$(Enc $tc.TenMoi)</b></td>
            <td><form method='post' action='/approve?code=$($_.Code)' style='margin:0'><button type='submit' class='btn'>Duyệt</button></form></td>
        </tr>"
    }) -join "`n"
}

function Build-TTChungTable {
    param($Items)
    if (-not $Items -or $Items.Count -eq 0) {
        return "<tr><td colspan='10' style='text-align:center;color:#888'>Không có mã nào đang chờ duyệt TT chung</td></tr>"
    }
    ($Items | ForEach-Object {
        $t = $_.Tabs.TTChung
        $nguon = if ($_.Source) { Enc $_.Source } else { "Mã mới" }
        "<tr>
            <td>$(Enc $_.Code)</td>
            <td>$(Enc $t.TenTCPH)</td>
            <td>$(Enc $t.TenChungKhoan)</td>
            <td>$(Enc $t.TenGiaoDichTiengAnh)</td>
            <td>$(Enc $t.TenTiengAnh)</td>
            <td>$(Enc $t.ThiTruong)</td>
            <td>$(Enc $t.NoiQuanLyVSD)</td>
            <td>$(Enc $t.MaISIN)</td>
            <td><span class='src'>$nguon</span></td>
            <td><form method='post' action='/approve?code=$($_.Code)' style='margin:0'><button type='submit' class='btn'>Duyệt</button></form></td>
        </tr>"
    }) -join "`n"
}

function Build-ChungKhoanTable-Detailed {
    # Bang chi tiet day du - dung chung cho cac loai DA co bang mapping ro rang
    # (hien: Co phieu va Chung chi quy - 2 loai nay dung chung 1 rule, chi khac
    # "Loai chung khoan").
    param($Items, [string]$EmptyLabel = "chứng khoán")
    if (-not $Items -or $Items.Count -eq 0) {
        return "<tr><td colspan='9' style='text-align:center;color:#888'>Không có mã $EmptyLabel nào đang chờ duyệt Chứng khoán</td></tr>"
    }
    ($Items | ForEach-Object {
        if ($_.Tabs -and $_.Tabs.ChungKhoan) {
            $c = $_.Tabs.ChungKhoan
            $noiGdCell = Enc $c.NoiGD
            $loai = Enc $c.LoaiChungKhoan
            $loaiTP = Enc $c.LoaiTraiPhieu
            $phiLuuKy = Enc $c.CoThuPhiLuuKy
            $menhGia = Enc $c.MenhGia
            $nguon = if ($_.Source) { "$(Enc $_.Source) (từ tab TT chung)" } else { "Mã mới (từ tab TT chung)" }
        } else {
            # Ma chuyen san (da co san tren Flex) - suy ra Loai chung khoan tu chinh
            # StockType da luu tren record, khong hardcode 1 loai co dinh.
            $tc = $_.TraCuuChuyenSan
            $ckTmp = New-ChungKhoanTabData -Code $_.Code -Market $_.Market -StockType $_.StockType -MenhGiaVSD $null
            $noiGdCell = "$(Enc $tc.NoiGDCu) &rarr; <b>$(Enc $tc.NoiGDMoi)</b>"
            $loai = Enc $ckTmp.LoaiChungKhoan
            $loaiTP = Enc $ckTmp.LoaiTraiPhieu
            $phiLuuKy = Enc $ckTmp.CoThuPhiLuuKy
            $menhGia = "-"
            $nguon = "Chuyển sàn (mã đã có trên Flex)"
        }
        "<tr>
            <td>$(Enc $_.Code)</td>
            <td>$(Enc $_.Name)</td>
            <td>$noiGdCell</td>
            <td>$loai</td>
            <td>$loaiTP</td>
            <td>$phiLuuKy</td>
            <td>$menhGia</td>
            <td><span class='src'>$nguon</span></td>
            <td><form method='post' action='/approve?code=$($_.Code)' style='margin:0'><button type='submit' class='btn'>Duyệt</button></form></td>
        </tr>"
    }) -join "`n"
}

function Build-ChungKhoanTable-TraiPhieu {
    # Bang chi tiet rieng cho Trai phieu - them 2 cot "Loai ky han" / "Ky han" so voi bang
    # dung chung (Build-ChungKhoanTable-Detailed), va "Loai trai phieu"/"Noi GD" lay DONG
    # tu VSD (khac Tin phieu la luon fix cung gia tri).
    param($Items)
    if (-not $Items -or $Items.Count -eq 0) {
        return "<tr><td colspan='11' style='text-align:center;color:#888'>Không có mã trái phiếu nào đang chờ duyệt Chứng khoán</td></tr>"
    }
    ($Items | ForEach-Object {
        if ($_.Tabs -and $_.Tabs.ChungKhoan) {
            $c = $_.Tabs.ChungKhoan
            $noiGdCell = Enc $c.NoiGD
            $loai = Enc $c.LoaiChungKhoan
            $loaiTP = Enc $c.LoaiTraiPhieu
            $phiLuuKy = Enc $c.CoThuPhiLuuKy
            $menhGia = Enc $c.MenhGia
            $loaiKyHan = if ($c.LoaiKyHan) { Enc $c.LoaiKyHan } else { "-" }
            $kyHan = if ($c.KyHan) { Enc $c.KyHan } else { "-" }
            $nguon = if ($_.Source) { "$(Enc $_.Source) (từ tab TT chung)" } else { "Mã mới (từ tab TT chung)" }
        } else {
            $tc = $_.TraCuuChuyenSan
            $noiGdCell = "$(Enc $tc.NoiGDCu) &rarr; <b>$(Enc $tc.NoiGDMoi)</b>"
            $loai = "Trái phiếu"
            $loaiTP = Enc $_.StockType
            $phiLuuKy = "Có"
            $menhGia = "-"
            $loaiKyHan = "-"
            $kyHan = "-"
            $nguon = "Chuyển sàn (mã đã có trên Flex)"
        }
        "<tr>
            <td>$(Enc $_.Code)</td>
            <td>$(Enc $_.Name)</td>
            <td>$noiGdCell</td>
            <td>$loai</td>
            <td>$loaiTP</td>
            <td>$phiLuuKy</td>
            <td>$menhGia</td>
            <td>$loaiKyHan</td>
            <td>$kyHan</td>
            <td><span class='src'>$nguon</span></td>
            <td><form method='post' action='/approve?code=$($_.Code)' style='margin:0'><button type='submit' class='btn'>Duyệt</button></form></td>
        </tr>"
    }) -join "`n"
}

function Build-ChungKhoanTable-ChungQuyen {
    # Bang chi tiet rieng cho Chung quyen - nhieu cot dac thu (Ma CKCS, TCPH CKCS, Loai
    # chung quyen, phuong thuc/gia thanh toan, gia thuc hien, ty le chuyen doi, thoi han,
    # ngay dao han, ngay giao dich cuoi cung) khong giong loai nao khac.
    param($Items)
    if (-not $Items -or $Items.Count -eq 0) {
        return "<tr><td colspan='16' style='text-align:center;color:#888'>Không có mã chứng quyền nào đang chờ duyệt Chứng khoán</td></tr>"
    }
    ($Items | ForEach-Object {
        if ($_.Tabs -and $_.Tabs.ChungKhoan) {
            $c = $_.Tabs.ChungKhoan
            $noiGdCell = Enc $c.NoiGD
            $loaiCKCoSo = Enc $c.LoaiChungKhoanCoSo
            $maCKCS = Enc $c.MaCKCS
            $tenTCPHCKCS = Enc $c.TenTCPHCKCS
            $loaiCQ = Enc $c.LoaiChungQuyen
            $ptThanhToan = Enc $c.PhuongThucThanhToan
            $giaThanhToan = Enc $c.GiaThanhToan
            $giaThucHien = Enc $c.GiaThucHien
            $tyLeChuyenDoi = Enc $c.TyLeChuyenDoi
            $thoiHan = if ($c.ThoiHanCWThang) { "$(Enc $c.ThoiHanCWThang) tháng" } else { "-" }
            $ngayDaoHan = Enc $c.NgayDaoHan
            $ngayCuoiCung = Enc $c.NgayGiaoDichCuoiCung
            $nguon = if ($_.Source) { "$(Enc $_.Source) (từ tab TT chung)" } else { "Mã mới (từ tab TT chung)" }
        } else {
            $tc = $_.TraCuuChuyenSan
            $noiGdCell = "$(Enc $tc.NoiGDCu) &rarr; <b>$(Enc $tc.NoiGDMoi)</b>"
            $loaiCKCoSo = "Cổ phiếu"
            $maCKCS = "-"; $tenTCPHCKCS = "-"; $loaiCQ = "-"; $ptThanhToan = "-"
            $giaThanhToan = "0.0000"; $giaThucHien = "-"; $tyLeChuyenDoi = "-"
            $thoiHan = "-"; $ngayDaoHan = "-"; $ngayCuoiCung = "-"
            $nguon = "Chuyển sàn (mã đã có trên Flex)"
        }
        "<tr>
            <td>$(Enc $_.Code)</td>
            <td>$(Enc $_.Name)</td>
            <td>$noiGdCell</td>
            <td>$loaiCKCoSo</td>
            <td>$maCKCS</td>
            <td>$tenTCPHCKCS</td>
            <td>$loaiCQ</td>
            <td>$ptThanhToan</td>
            <td>$giaThanhToan</td>
            <td>$giaThucHien</td>
            <td>$tyLeChuyenDoi</td>
            <td>$thoiHan</td>
            <td>$ngayDaoHan</td>
            <td>$ngayCuoiCung</td>
            <td><span class='src'>$nguon</span></td>
            <td><form method='post' action='/approve?code=$($_.Code)' style='margin:0'><button type='submit' class='btn'>Duyệt</button></form></td>
        </tr>"
    }) -join "`n"
}

function Build-CkSearchResultsHtml {
    param($Flex, [string]$Query)
    # Ma da xoa (tab "Da xoa") KHONG hien trong tim kiem - phai khoi phuc truoc moi thay lai duoc
    $matches = @($Flex | Where-Object { $_.Code -like "*$Query*" -and $_.Status -ne "Đã xóa" } | Select-Object -First 20)
    if ($matches.Count -eq 0) {
        return "<p class='muted'>Không tìm thấy mã chứng khoán nào khớp với `"$(Enc $Query)`".</p>"
    }
    $rows = ($matches | ForEach-Object {
        $maTcph = Get-SyntheticId -Code $_.Code -Seed 31
        $maGiaoDich = Get-SyntheticId -Code $_.Code -Seed 17
        "<tr>
            <td>$maTcph</td>
            <td>$(Enc $_.Code)</td>
            <td>$maGiaoDich</td>
            <td>$(Enc $_.Name)</td>
            <td>$(Enc $_.Market)</td>
            <td>$(Enc $_.StockType)</td>
            <td>-</td>
            <td>
                <a class='minibtn' href='/view-security?code=$([System.Uri]::EscapeDataString($_.Code))&q=$([System.Uri]::EscapeDataString($Query))'>Xem</a>
                <a class='minibtn' href='/edit-security?code=$([System.Uri]::EscapeDataString($_.Code))&q=$([System.Uri]::EscapeDataString($Query))'>Sửa</a>
                <a class='minibtn minibtn-danger' href='#' onclick='if(confirm(&#39;Xoa vinh vien ma $(Enc $_.Code)? Hanh dong nay khong the hoan tac.&#39;)){document.getElementById(&#39;delform-$(Enc $_.Code)&#39;).submit();} return false;'>Xóa</a>
                <form id='delform-$(Enc $_.Code)' method='post' action='/delete-security?code=$([System.Uri]::EscapeDataString($_.Code))' style='display:none'>
                    <input type='hidden' name='q' value='$(Enc $Query)' />
                </form>
            </td>
        </tr>"
    }) -join "`n"
    return @"
<div class="closebar"><a class="minibtn" href="/?tab=ck">&#10003; Đóng</a></div>
<table>
<tr><th>Mã TCPH</th><th>Chứng khoán</th><th>Mã giao dịch</th><th>Tên</th><th>Nơi giao dịch</th><th>Loại chứng khoán</th><th>Chặn không GD lô lẻ</th><th></th></tr>
$rows
</table>
"@
}

function Build-StandalonePage {
    # Trang HTML rieng, TOI GIAN - chi chua noi dung Xem/Sua 1 ma, KHONG kem theo phan
    # dashboard/tab khac ben duoi. Dung khi bam Xem/Sua tu ket qua tim kiem tab Chung khoan.
    param([string]$Title, [string]$BodyHtml)
    return @"
<!doctype html>
<html><head><meta charset="utf-8">
<title>$Title</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; background:#f6f7f9; color:#1a1a1a }
  table { border-collapse: collapse; width:100%; background:#fff; box-shadow:0 1px 3px rgba(0,0,0,0.1) }
  table.mini { margin-bottom:14px }
  th, td { padding:8px 10px; border-bottom:1px solid #eee; text-align:left; font-size:13px }
  th { background:#fafafa }
  .btn { background:#2563eb; color:#fff; border:none; padding:6px 12px; border-radius:5px; cursor:pointer; font-size:13px }
  .btn:hover { background:#1d4ed8 }
  .muted { color:#999 }
  h3 { font-size:13px; color:#666; margin:20px 0 8px }
  .minibtn { display:inline-block; background:#f0f0f0; color:#2563eb; border:1px solid #ddd; padding:3px 9px; border-radius:4px; font-size:12px; text-decoration:none; margin-right:4px }
  .minibtn:hover { background:#e5e7eb }
  .editbox { background:#fff; border:1px solid #e5e7eb; border-radius:6px; padding:14px 18px; margin-bottom:16px }
  .closebar { text-align:right; margin-bottom:8px }
  .closebar .minibtn { background:#eafbea; color:#1a7a1a; border-color:#bfe8bf }
  .editnav { display:flex; gap:4px; border-bottom:1px solid #e5e7eb; margin-bottom:14px }
  .editnav span { padding:6px 14px; font-size:12px }
  .editnav-active { color:#2563eb; font-weight:600; border-bottom:2px solid #2563eb }
  .editnav-disabled { color:#bbb }
  .formrow { display:flex; align-items:center; margin-bottom:8px }
  .formrow label { width:160px; font-size:12px; color:#555 }
  .formrow input, .formrow select { flex:1; max-width:280px; padding:6px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px }
  .formrow input:disabled { background:#f5f5f5; color:#888 }
</style>
</head>
<body>
$BodyHtml
</body></html>
"@
}

function Build-DeletedBlockedHtml {
    param($Item, [string]$Action)
    return "<div class='editbox'><p class='muted'>Mã <b>$(Enc $Item.Code)</b> đã bị xóa nên không thể $Action. Vui lòng vào tab <b>Đã xóa</b> để khôi phục trước.</p><form method='post' action='/restore-security?code=$([System.Uri]::EscapeDataString($Item.Code))' style='margin:0'><button type='submit' class='btn'>Khôi phục ngay</button></form></div>"
}

function Build-SelectOptionsHtml {
    # Dung chung cho cac dropdown trong man Sua. Neu gia tri hien tai KHONG nam trong
    # danh sach option chuan (vd du lieu VSD co bien the la), tu them no vao dau danh
    # sach va danh dau selected - tranh viec bam "Cap nhat" ma khong sua gi lai vo tinh
    # doi sang gia tri khac trong dropdown.
    param([string[]]$Options, [string]$CurrentValue)
    $all = @(if ($CurrentValue -and ($Options -notcontains $CurrentValue)) { @($CurrentValue) + $Options } else { $Options })
    return ($all | ForEach-Object {
        $sel = if ($_ -eq $CurrentValue) { " selected" } else { "" }
        "<option value='$(Enc $_)'$sel>$(Enc $_)</option>"
    }) -join ''
}

function Build-CkEditFormHtml {
    # Sua duoc TAT CA cac truong TRU "Trang thai CK" (truong nay do luong Duyet dieu
    # khien tu dong - cho sua tay se lam sai lech voi Lich su duyet/D-S Kiem soat).
    param($Item, [string]$CloseHref = "/")
    if (-not $Item) { return "<p class='muted'>Không tìm thấy mã để sửa.</p>" }
    if ($Item.Status -eq "Đã xóa") { return Build-DeletedBlockedHtml -Item $Item -Action "sửa" }
    $maQuiUoc = if ($Item.MaQuiUocOverride) { $Item.MaQuiUocOverride } else { Get-SyntheticId -Code $Item.Code -Seed 17 }
    $maTcph   = if ($Item.MaTCPHOverride) { $Item.MaTCPHOverride } else { Get-SyntheticId -Code $Item.Code -Seed 31 }
    $options = ($NoiGDOptions | ForEach-Object {
        $sel = if ($_ -eq $Item.Market) { " selected" } else { "" }
        "<option value='$_'$sel>$_</option>"
    }) -join ''
    $noiLuuKyOptionsHtml = Build-SelectOptionsHtml -Options $NoiLuuKyOptions -CurrentValue $Item.ManagementArea
    $loaiCKOptionsHtml = Build-SelectOptionsHtml -Options $LoaiChungKhoanOptions -CurrentValue $Item.StockType
    $loaiTraiPhieuVal = if ($Item.Tabs -and $Item.Tabs.ChungKhoan) { $Item.Tabs.ChungKhoan.LoaiTraiPhieu } else { $null }
    $loaiTraiPhieuField = if ($null -ne $loaiTraiPhieuVal) {
        "<input type='text' name='loaitraiphieu' value=`"$(Enc $loaiTraiPhieuVal)`" />"
    } else {
        "<input type='text' value='-' disabled title='Chưa có dữ liệu (chỉ sinh sau khi duyệt xong tab Chứng khoán)' />"
    }
    return @"
<div class="editbox">
  <div class="closebar"><a class="minibtn" href="$CloseHref">&#10003; Đóng</a></div>
  <div class="editnav">
    <span class="editnav-active">TT chung</span>
    <span class="editnav-disabled">Bước giá</span>
    <span class="editnav-disabled">TT CK</span>
  </div>
  <form method="post" action="/update-security?code=$([System.Uri]::EscapeDataString($Item.Code))">
    <div class="formrow"><label>Mã qui ước</label><input type="text" name="maquiuoc" value="$(Enc $maQuiUoc)" /></div>
    <div class="formrow"><label>Mã TCPH</label><input type="text" name="matcph" value="$(Enc $maTcph)" /></div>
    <div class="formrow"><label>Mã chứng khoán</label><input type="text" name="code" value="$(Enc $Item.Code)" /></div>
    <div class="formrow"><label>Nơi GD</label><select name="noigd">$options</select></div>
    <div class="formrow"><label>Nơi lưu ký</label><select name="noiluuky">$noiLuuKyOptionsHtml</select></div>
    <div class="formrow"><label>Loại chứng khoán</label><select name="loaick">$loaiCKOptionsHtml</select></div>
    <div class="formrow"><label>Trạng thái CK</label><input type="text" value="$(Enc $Item.Status)" disabled /></div>
    <div class="formrow"><label>Loại trái phiếu</label>$loaiTraiPhieuField</div>
    <button type="submit" class="btn">Cập nhật</button>
  </form>
</div>
"@
}

function Build-CkViewHtml {
    param($Item, [string]$CloseHref = "/")
    if (-not $Item) { return "<p class='muted'>Không tìm thấy mã để xem.</p>" }
    if ($Item.Status -eq "Đã xóa") { return Build-DeletedBlockedHtml -Item $Item -Action "xem" }
    $maQuiUoc = if ($Item.MaQuiUocOverride) { $Item.MaQuiUocOverride } else { Get-SyntheticId -Code $Item.Code -Seed 17 }
    $maTcph   = if ($Item.MaTCPHOverride) { $Item.MaTCPHOverride } else { Get-SyntheticId -Code $Item.Code -Seed 31 }

    $ttChungBlock = if ($Item.Tabs -and $Item.Tabs.TTChung) {
        $t = $Item.Tabs.TTChung
        $enRows = ""
        if ($t.TenGiaoDichTiengAnh -or $t.TenTiengAnh) {
            $enRows = "<tr><td>Tên giao dịch tiếng Anh</td><td>$(Enc $t.TenGiaoDichTiengAnh)</td></tr><tr><td>Tên tiếng Anh</td><td>$(Enc $t.TenTiengAnh)</td></tr>"
        }
        @"
<table class='mini'>
<tr><td>Tổ chức phát hành</td><td>$(Enc $t.TenTCPH)</td></tr>
<tr><td>Tên giao dịch</td><td>$(Enc $t.TenChungKhoan)</td></tr>
$enRows
<tr><td>Thị trường</td><td>$(Enc $t.ThiTruong)</td></tr>
<tr><td>Nơi quản lý tại VSD</td><td>$(Enc $t.NoiQuanLyVSD)</td></tr>
<tr><td>Mã ISIN</td><td>$(Enc $t.MaISIN)</td></tr>
</table>
"@
    } else {
        "<table class='mini'><tr><td>Mã ISIN</td><td>$(Enc $Item.IsinCode)</td></tr><tr><td>Tên</td><td>$(Enc $Item.Name)</td></tr></table>"
    }

    $traCuuDoiTenBlock = if ($Item.TraCuuDoiTen) {
        $td = $Item.TraCuuDoiTen
        @"
<h3>Tra cứu đổi tên (mục 4)</h3>
<table class='mini'>
<tr><td>Tên cũ (Flex)</td><td>$(Enc $td.TenCu)</td></tr>
<tr><td>Tên mới (VSD)</td><td><b>$(Enc $td.TenMoi)</b></td></tr>
<tr><td>Thời điểm tra cứu</td><td>$(Enc $td.TraCuuLuc)</td></tr>
</table>
"@
    } else { "" }

    $ckBlock = if ($Item.Tabs -and $Item.Tabs.ChungKhoan) {
        $c = $Item.Tabs.ChungKhoan
        $kyHanRows = if ($c.LoaiKyHan -or $c.KyHan) {
            "<tr><td>Loại kỳ hạn</td><td>$(Enc $c.LoaiKyHan)</td></tr><tr><td>Kỳ hạn</td><td>$(Enc $c.KyHan)</td></tr>"
        } else { "" }
        $cwRows = if ($c.LoaiChungKhoanCoSo) {
            "<tr><td>Loại chứng khoán cơ sở</td><td>$(Enc $c.LoaiChungKhoanCoSo)</td></tr>
<tr><td>Mã CKCS</td><td>$(Enc $c.MaCKCS)</td></tr>
<tr><td>Tên TCPH CKCS</td><td>$(Enc $c.TenTCPHCKCS)</td></tr>
<tr><td>Loại chứng quyền</td><td>$(Enc $c.LoaiChungQuyen)</td></tr>
<tr><td>Phương thức thanh toán</td><td>$(Enc $c.PhuongThucThanhToan)</td></tr>
<tr><td>Giá thanh toán</td><td>$(Enc $c.GiaThanhToan)</td></tr>
<tr><td>Giá thực hiện</td><td>$(Enc $c.GiaThucHien)</td></tr>
<tr><td>Tỷ lệ chuyển đổi</td><td>$(Enc $c.TyLeChuyenDoi)</td></tr>
<tr><td>Thời hạn CW theo tháng</td><td>$(Enc $c.ThoiHanCWThang)</td></tr>
<tr><td>Ngày đáo hạn</td><td>$(Enc $c.NgayDaoHan)</td></tr>
<tr><td>Ngày giao dịch cuối cùng</td><td>$(Enc $c.NgayGiaoDichCuoiCung)</td></tr>"
        } else { "" }
        @"
<table class='mini'>
<tr><td>Nơi GD</td><td>$(Enc $c.NoiGD)</td></tr>
<tr><td>Loại chứng khoán</td><td>$(Enc $c.LoaiChungKhoan)</td></tr>
<tr><td>Loại trái phiếu</td><td>$(Enc $c.LoaiTraiPhieu)</td></tr>
<tr><td>Có thu phí lưu ký không</td><td>$(Enc $c.CoThuPhiLuuKy)</td></tr>
<tr><td>Mệnh giá</td><td>$(Enc $c.MenhGia)</td></tr>
$kyHanRows
$cwRows
</table>
"@
    } else {
        "<table class='mini'><tr><td>Nơi GD</td><td>$(Enc $Item.Market)</td></tr><tr><td>Loại chứng khoán</td><td>$(Enc $Item.StockType)</td></tr></table>"
    }

    $ttckBlock = if ($Item.Tabs -and $Item.Tabs.TTCK) {
        $k = $Item.Tabs.TTCK
        @"
<table class='mini'>
<tr><td>Đơn vị giao dịch</td><td>$(Enc $k.DonViGiaoDich)</td></tr>
<tr><td>Khối lượng niêm yết</td><td>$(Enc $k.KhoiLuongNiemYet)</td></tr>
<tr><td>Ngày niêm yết</td><td>$(Enc $k.NgayNiemYet)</td></tr>
<tr><td>Mua bán cùng ngày</td><td>$(Enc $k.MuaBanCungNgay)</td></tr>
<tr><td>Check room NĐT nước ngoài</td><td>$(Enc $k.CheckRoomNDTNuocNgoai)</td></tr>
</table>
"@
    } else {
        "<p class='muted'><i>Chưa có dữ liệu TTCK (chỉ sinh sau khi duyệt xong tab Chứng khoán).</i></p>"
    }

    $lichSuBlock = Build-KiemSoatHtmlForItem -Item $Item

    return @"
<div class="editbox">
  <div class="closebar"><a class="minibtn" href="$CloseHref">&#10003; Đóng</a></div>
  <div class="formrow"><label>Mã qui ước</label><input type="text" value="$maQuiUoc" disabled /></div>
  <div class="formrow"><label>Mã TCPH</label><input type="text" value="$maTcph" disabled /></div>
  <div class="formrow"><label>Mã chứng khoán</label><input type="text" value="$(Enc $Item.Code)" disabled /></div>
  <div class="formrow"><label>Trạng thái CK</label><input type="text" value="$(Enc $Item.Status)" disabled /></div>
  <div class="formrow"><label>Nguồn</label><input type="text" value="$(Enc $(if ($Item.Source) { $Item.Source } else { "Mã mới" }))" disabled /></div>

  <h3>TT chung</h3>
  $ttChungBlock
  $traCuuDoiTenBlock

  <h3>Chứng khoán</h3>
  $ckBlock

  <h3>TTCK</h3>
  $ttckBlock

  <h3>Lịch sử duyệt (D/S Kiểm soát)</h3>
  $lichSuBlock
</div>
"@
}

function Build-KiemSoatHtmlForItem {
    param($Item)
    if (-not $Item.LichSuDuyet -or @($Item.LichSuDuyet).Count -eq 0) {
        return "<p class='muted'><i>Chưa có lịch sử duyệt.</i></p>"
    }
    $sorted = @($Item.LichSuDuyet | Sort-Object ThoiGian -Descending)
    $rows = ($sorted | ForEach-Object {
        "<tr><td>$(Enc $_.ThoiGian)</td><td>$(Enc-Multiline $_.HanhDong)</td><td>$(Enc $_.TrangThai)</td></tr>"
    }) -join "`n"
    return "<table class='mini'><tr><td><b>Thời gian</b></td><td><b>Hành động</b></td><td><b>Trạng thái</b></td></tr>$rows</table>"
}

function Get-OldCodesConfig {
    if (Test-Path $OldCodesConfigPath) {
        return Get-Content $OldCodesConfigPath -Raw | ConvertFrom-Json
    }
    return $null
}

function Save-OldCodesConfig {
    param([int]$BatchSize)
    $dir = Split-Path $OldCodesConfigPath -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [pscustomobject]@{ BatchSize = $BatchSize } | ConvertTo-Json | Out-File -FilePath $OldCodesConfigPath -Encoding utf8
}

function Get-OldCodesBacklog {
    if (-not (Test-Path $OldCodesBacklogPath)) { return @() }
    return [object[]](Get-Content $OldCodesBacklogPath -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Save-OldCodesBacklog {
    param($Backlog)
    $dir = Split-Path $OldCodesBacklogPath -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    @($Backlog) | ConvertTo-Json -Depth 5 | Out-File -FilePath $OldCodesBacklogPath -Encoding utf8
}

function Get-LatestDiffNewCodes {
    # Dung de "khoi tao" hang doi ma cu tu ban bao cao so sanh VSD-Flex gan nhat
    if (-not (Test-Path $ReportsDir)) { return @() }
    $latest = Get-ChildItem -Path $ReportsDir -Filter "flex_diff_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) { return @() }
    $report = Get-Content $latest.FullName -Raw | ConvertFrom-Json
    return @($report.NewCodes)
}

function Invoke-OldCodesBatch {
    # Lay toi da BatchSize ma tu hang doi, nhap vao tab TT chung (Source = "Ma cu ton dong")
    # Giong het logic Process-OldSecurities.ps1, dung truc tiep trong web server de nhan vien
    # bam nut tren UI la chay ngay, khong can mo dong lenh.
    $cfg = Get-OldCodesConfig
    if (-not $cfg) { return @{ Ok = $false; Message = "Chua cau hinh so luong xu ly moi lan." } }
    $batchSize = [int]$cfg.BatchSize
    $backlog = Get-OldCodesBacklog
    if ($backlog.Count -eq 0) { return @{ Ok = $false; Message = "Hang doi ma cu dang rong." } }

    $flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)
    $existingCodes = @($flex | ForEach-Object { $_.Code })
    $stillMissing = $backlog | Where-Object { $_.Code -notin $existingCodes }
    $batch = @($stillMissing | Select-Object -First $batchSize)
    $handledCodes = $backlog | Where-Object { $_.Code -in $existingCodes -or $_.Code -in ($batch | ForEach-Object { $_.Code }) }
    $remaining = @($backlog | Where-Object { $_.Code -notin ($handledCodes | ForEach-Object { $_.Code }) })

    if ($batch.Count -eq 0) {
        Save-OldCodesBacklog -Backlog $remaining
        return @{ Ok = $false; Message = "Khong co ma nao can nhap trong lo nay."; Remaining = $remaining.Count }
    }

    foreach ($code in $batch) {
        $detail = $null
        if ($code.DetailId) { $detail = Get-VsdSecurityDetail -DetailId $code.DetailId }
        $now = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $flex.Add([pscustomobject]@{
            Code            = $code.Code
            IsinCode        = $code.IsinCode
            Name            = $code.Name
            StockType       = $code.StockType
            Market          = $code.Market
            ManagementArea  = $code.ManagementArea
            DetailId        = $code.DetailId
            MenhGiaVSD      = if ($detail) { $detail.MenhGia } else { $null }
            TongSoDangKyVSD = if ($detail) { $detail.TongSoDangKy } else { $null }
            Status          = "Chờ duyệt TT chung"
            Source          = "Mã cũ (tồn đọng)"
            StatusChangedAt = $now
            LichSuDuyet     = @([pscustomobject]@{ ThoiGian = $now; HanhDong = "Nhap tab TT chung (ma cu ton dong, qua UI, Ten TCPH tu trang chi tiet VSD)"; TrangThai = "Chờ duyệt TT chung" })
            Tabs            = [pscustomobject]@{
                TTChung = [pscustomobject]@{
                    MaCK = $code.Code
                    TenTCPH = if ($detail -and $detail.TenTCPH) { $detail.TenTCPH } else { $code.Name }
                    TenChungKhoan = if ($detail -and $detail.TenChungKhoan) { $detail.TenChungKhoan } else { $code.Name }
                    TenGiaoDichTiengAnh = "Đang bổ sung"
                    TenTiengAnh = "Đang bổ sung"
                    ThiTruong = "Trong nước"
                    NoiQuanLyVSD = Get-NoiQuanLyVSDText -ManagementArea $code.ManagementArea
                    MaISIN = $code.IsinCode
                }
                ChungKhoan = $null
                TTCK       = $null
            }
        })
    }
    Save-FlexStore -Path $FlexStorePath -Data $flex
    Save-OldCodesBacklog -Backlog $remaining
    $batchCodesStr = ($batch | ForEach-Object { $_.Code }) -join ', '
    Send-ApproveNotice -Code $batchCodesStr -NewStatus "Chờ duyệt TT chung" -HanhDong "Đã nhập $($batch.Count) mã cũ tồn đọng vào tab TT chung qua UI. Còn lại $($remaining.Count) mã trong hàng đợi."
    return @{ Ok = $true; Processed = $batch.Count; Remaining = $remaining.Count }
}

function Build-OldCodesTabHtml {
    $cfg = Get-OldCodesConfig
    $backlog = Get-OldCodesBacklog

    if (-not $cfg) {
        $newCodesAvailable = @(Get-LatestDiffNewCodes).Count
        return @"
<div class="editbox">
  <p class="muted">Chưa khởi tạo hàng đợi mã cũ. Nhân viên nhập số lượng xử lý mỗi lần khi mở phần mềm lần đầu (mục 3, CR).</p>
  <p>Hiện có <b>$newCodesAvailable</b> mã đang thiếu trên Flex theo lần so sánh VSD gần nhất - sẽ được đưa vào hàng đợi làm "mã cũ tồn đọng".</p>
  <form method="post" action="/old-codes/init">
    <div class="formrow"><label>Số lượng xử lý mỗi lần</label><input type="number" name="batchSize" value="5" min="1" required /></div>
    <button type="submit" class="btn">Khởi tạo hàng đợi</button>
  </form>
</div>
"@
    }

    $backlogRows = if ($backlog.Count -eq 0) {
        "<tr><td colspan='3' style='text-align:center;color:#888'>Hàng đợi trống</td></tr>"
    } else {
        ($backlog | ForEach-Object { "<tr><td>$(Enc $_.Code)</td><td>$(Enc $_.Name)</td><td>$(Enc $_.Market)</td></tr>" }) -join "`n"
    }

    return @"
<div class="editbox">
  <div class="formrow"><label>Số lượng xử lý mỗi lần hiện tại</label><input type="text" value="$($cfg.BatchSize) mã" disabled /></div>
  <form method="post" action="/old-codes/set-batch-size" style="margin-bottom:14px">
    <div class="formrow"><label>Đổi số lượng xử lý mỗi lần</label><input type="number" name="batchSize" value="$($cfg.BatchSize)" min="1" required /></div>
    <button type="submit" class="btn">Lưu</button>
  </form>
  <p>Hàng đợi còn lại: <b>$($backlog.Count) mã</b></p>
  <form method="post" action="/old-codes/process-batch" style="margin-bottom:14px">
    <button type="submit" class="btn">Xử lý 1 lô ngay ($([math]::Min($cfg.BatchSize, $backlog.Count)) mã)</button>
  </form>
  <h3>Danh sách còn trong hàng đợi</h3>
  <table><tr><th>Mã CK</th><th>Tên</th><th>Nơi GD</th></tr>$backlogRows</table>
</div>
"@
}

function Build-DeletedTable {
    param($Items, [int]$RetentionDays)
    if (-not $Items -or $Items.Count -eq 0) {
        return "<tr><td colspan='6' style='text-align:center;color:#888'>Không có mã nào trong thùng rác</td></tr>"
    }
    $now = Get-Date
    ($Items | ForEach-Object {
        $daysLeftText = "?"
        try {
            $deletedAt = [datetime]::ParseExact($_.DeletedAt, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
            $daysLeftText = "$([math]::Max(0, $RetentionDays - [math]::Floor(($now - $deletedAt).TotalDays))) ngày"
        } catch {}
        "<tr>
            <td>$(Enc $_.Code)</td>
            <td>$(Enc $_.Name)</td>
            <td>$(Enc $_.DeletedAt)</td>
            <td>$(Enc $_.StatusBeforeDelete)</td>
            <td>$daysLeftText</td>
            <td><form method='post' action='/restore-security?code=$($_.Code)' style='margin:0'><button type='submit' class='btn'>Khôi phục</button></form></td>
        </tr>"
    }) -join "`n"
}

function Get-AllKiemSoatStatuses {
    param($Flex)
    $statuses = New-Object System.Collections.Generic.HashSet[string]
    foreach ($item in $Flex) {
        foreach ($entry in @($item.LichSuDuyet)) {
            if ($entry -and $entry.TrangThai) { $statuses.Add($entry.TrangThai) | Out-Null }
        }
    }
    return @($statuses) | Sort-Object
}

function Build-KiemSoatFilterHtml {
    param([string]$CodeFilter, [string]$StatusFilter, [string[]]$AllStatuses)
    $options = "<option value=''>-- Tất cả --</option>" + (($AllStatuses | ForEach-Object {
        $sel = if ($_ -eq $StatusFilter) { " selected" } else { "" }
        "<option value='$(Enc $_)'$sel>$(Enc $_)</option>"
    }) -join '')
    return @"
<form method="get" action="/" class="ks-filter">
  <input type="hidden" name="tab" value="ks" />
  <input type="text" name="ks_code" placeholder="Lọc theo mã CK..." value="$(Enc $CodeFilter)" />
  <select name="ks_status">$options</select>
  <button type="submit" class="btn">Lọc</button>
  <a class="minibtn" href="/?tab=ks">Xóa lọc</a>
</form>
"@
}

function Build-KiemSoatTable {
    param($Flex, [int]$Limit = 30, [string]$CodeFilter = "", [string]$StatusFilter = "")
    $all = New-Object System.Collections.Generic.List[object]
    foreach ($item in $Flex) {
        foreach ($entry in @($item.LichSuDuyet)) {
            if (-not $entry) { continue }
            $all.Add([pscustomobject]@{ Code = $item.Code; Name = $item.Name; ThoiGian = $entry.ThoiGian; HanhDong = $entry.HanhDong; TrangThai = $entry.TrangThai })
        }
    }
    $filtered = $all
    if (-not [string]::IsNullOrWhiteSpace($CodeFilter)) {
        $filtered = $filtered | Where-Object { $_.Code -like "*$CodeFilter*" }
    }
    if (-not [string]::IsNullOrWhiteSpace($StatusFilter)) {
        $filtered = $filtered | Where-Object { $_.TrangThai -eq $StatusFilter }
    }
    $sorted = @($filtered | Sort-Object ThoiGian -Descending | Select-Object -First $Limit)
    if ($sorted.Count -eq 0) {
        return "<tr><td colspan='5' style='text-align:center;color:#888'>Không có kết quả khớp với bộ lọc.</td></tr>"
    }
    ($sorted | ForEach-Object {
        "<tr><td>$(Enc $_.ThoiGian)</td><td>$(Enc $_.Code)</td><td>$(Enc $_.Name)</td><td>$(Enc-Multiline $_.HanhDong)</td><td>$(Enc $_.TrangThai)</td></tr>"
    }) -join "`n"
}

function Build-Html {
    param([string]$CkSearchQuery = "", [string]$EditCode = "", [string]$ViewCode = "", [string]$ViewFrom = "", [string]$ActiveTab = "", [string]$KsCodeFilter = "", [string]$KsStatusFilter = "")

    $flex = Get-FlexStore -Path $FlexStorePath
    $waitingTTChungAll = $flex | Where-Object { $_.Status -eq "Chờ duyệt TT chung" }
    $waitingNameChange = $waitingTTChungAll | Where-Object { $_.TraCuuDoiTen }
    $waitingTTChung    = $waitingTTChungAll | Where-Object { -not $_.TraCuuDoiTen }
    $waitingChungKhoan = $flex | Where-Object { $_.Status -eq "Chờ duyệt Chứng khoán" }
    $deletedItems      = $flex | Where-Object { $_.Status -eq "Đã xóa" }

    $ttChungRows      = Build-TTChungTable -Items $waitingTTChung
    $nameChangeRows   = Build-NameChangesTable -Items $waitingNameChange

    # Chia ma dang cho duyet Chung khoan thanh 5 nhom theo dung 5 sheet trong Excel
    # "Cach thuc khai ma chung khoan" (Co phieu / Chung chi quy / Tin phieu / Trai
    # phieu / Chung quyen). Hien chi Co phieu co day du bang mapping rieng.
    $ckByCategory = @{ CoPhieu = @(); ChungChiQuy = @(); TinPhieu = @(); TraiPhieu = @(); ChungQuyen = @() }
    foreach ($item in $waitingChungKhoan) {
        $cat = Get-SecurityCategory -StockType $item.StockType
        $ckByCategory[$cat] += $item
    }
    $ckCoPhieuRows     = Build-ChungKhoanTable-Detailed -Items $ckByCategory.CoPhieu -EmptyLabel "cổ phiếu"
    $ckChungChiQuyRows = Build-ChungKhoanTable-Detailed -Items $ckByCategory.ChungChiQuy -EmptyLabel "chứng chỉ quỹ"
    $ckTinPhieuRows    = Build-ChungKhoanTable-Detailed -Items $ckByCategory.TinPhieu -EmptyLabel "tín phiếu"
    $ckTraiPhieuRows   = Build-ChungKhoanTable-TraiPhieu -Items $ckByCategory.TraiPhieu
    $ckChungQuyenRows  = Build-ChungKhoanTable-ChungQuyen -Items $ckByCategory.ChungQuyen

    $ksAllStatuses    = Get-AllKiemSoatStatuses -Flex $flex
    $ksFilterHtml     = Build-KiemSoatFilterHtml -CodeFilter $KsCodeFilter -StatusFilter $KsStatusFilter -AllStatuses $ksAllStatuses
    $kiemSoatRows     = Build-KiemSoatTable -Flex $flex -CodeFilter $KsCodeFilter -StatusFilter $KsStatusFilter
    $deletedRows      = Build-DeletedTable -Items $deletedItems -RetentionDays $DeleteRetentionDays
    $oldCodesBlock    = Build-OldCodesTabHtml

    $countTT = @($waitingTTChung).Count
    $countNameChange = @($waitingNameChange).Count
    $countCK = @($waitingChungKhoan).Count
    $countCkCoPhieu     = @($ckByCategory.CoPhieu).Count
    $countCkChungChiQuy = @($ckByCategory.ChungChiQuy).Count
    $countCkTinPhieu    = @($ckByCategory.TinPhieu).Count
    $countCkTraiPhieu   = @($ckByCategory.TraiPhieu).Count
    $countCkChungQuyen  = @($ckByCategory.ChungQuyen).Count
    $countDeleted = @($deletedItems).Count
    $countOldBacklog = @(Get-OldCodesBacklog).Count

    # Xem/Sua chi truy cap qua tim kiem o tab Chung khoan (khong con nut Xem rieng tren tab TT chung)
    $ckSearchBlock = ""
    if (-not [string]::IsNullOrWhiteSpace($EditCode)) {
        $editItem = $flex | Where-Object { $_.Code -eq $EditCode } | Select-Object -First 1
        $ckSearchBlock = Build-CkEditFormHtml -Item $editItem -CloseHref "/?tab=ck"
    } elseif (-not [string]::IsNullOrWhiteSpace($ViewCode) -and $ViewFrom -ne "ttchung") {
        $viewItem = $flex | Where-Object { $_.Code -eq $ViewCode } | Select-Object -First 1
        $ckSearchBlock = Build-CkViewHtml -Item $viewItem -CloseHref "/?tab=ck"
    } elseif (-not [string]::IsNullOrWhiteSpace($CkSearchQuery)) {
        $ckSearchBlock = Build-CkSearchResultsHtml -Flex $flex -Query $CkSearchQuery
    }

    $ttChungChecked = "checked"
    $ckChecked = ""
    $ksChecked = ""
    $deletedChecked = ""
    $oldCodesChecked = ""
    if ($ActiveTab -eq "ck") {
        $ttChungChecked = ""
        $ckChecked = "checked"
    } elseif ($ActiveTab -eq "ks") {
        $ttChungChecked = ""
        $ksChecked = "checked"
    } elseif ($ActiveTab -eq "deleted") {
        $ttChungChecked = ""
        $deletedChecked = "checked"
    } elseif ($ActiveTab -eq "oldcodes") {
        $ttChungChecked = ""
        $oldCodesChecked = "checked"
    } elseif ($ActiveTab -eq "ttchung") {
        $ttChungChecked = "checked"
    } elseif (-not [string]::IsNullOrWhiteSpace($CkSearchQuery) -or -not [string]::IsNullOrWhiteSpace($EditCode) -or -not [string]::IsNullOrWhiteSpace($ViewCode)) {
        $ttChungChecked = ""
        $ckChecked = "checked"
    }

    return @"
<!doctype html>
<html><head><meta charset="utf-8">
<title>Flex (Mock) - 020004</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; background:#f6f7f9; color:#1a1a1a }
  .topbar { display:flex; justify-content:space-between; align-items:center; margin-bottom:16px }
  h1 { font-size: 20px; margin:0 }
  .logout { font-size:13px; color:#2563eb; text-decoration:none }
  .note { background:#fff3cd; border:1px solid #ffe58f; padding:10px 14px; border-radius:6px; margin-bottom:16px; font-size:13px }
  table { border-collapse: collapse; width:100%; background:#fff; box-shadow:0 1px 3px rgba(0,0,0,0.1) }
  th, td { padding:8px 10px; border-bottom:1px solid #eee; text-align:left; font-size:13px }
  th { background:#fafafa }
  .btn { background:#2563eb; color:#fff; border:none; padding:6px 12px; border-radius:5px; cursor:pointer; font-size:13px }
  .btn:hover { background:#1d4ed8 }
  .src { font-size:11px; color:#888; font-style:italic }

  .pagetabs input[type=radio] { display:none }
  .pagetabs .tabnav { display:flex; border-bottom:2px solid #e5e7eb; margin-bottom:14px }
  .pagetabs .tabnav label { padding:10px 18px; cursor:pointer; font-size:14px; color:#555; border-bottom:3px solid transparent; margin-bottom:-2px; user-select:none }
  .pagetabs .panel { display:none }
  .pagetabs .badge-count { background:#fde68a; color:#8a6d00; border-radius:10px; padding:1px 7px; font-size:11px; margin-left:6px }

  .cktypetabs { margin-top:8px }
  .cktypetabs input[type=radio] { display:none }
  .cktypetabs-nav { display:flex; gap:4px; border-bottom:1px solid #e5e7eb; margin-bottom:12px }
  .cktypetabs-nav label { padding:8px 14px; cursor:pointer; font-size:13px; color:#555; border-bottom:2px solid transparent; margin-bottom:-1px; user-select:none }
  .cktypepanel { display:none }
  #cktab-cophieu:checked ~ .cktypetabs-nav label[for=cktab-cophieu],
  #cktab-ccq:checked ~ .cktypetabs-nav label[for=cktab-ccq],
  #cktab-tinphieu:checked ~ .cktypetabs-nav label[for=cktab-tinphieu],
  #cktab-traiphieu:checked ~ .cktypetabs-nav label[for=cktab-traiphieu],
  #cktab-cq:checked ~ .cktypetabs-nav label[for=cktab-cq] { color:#2563eb; border-bottom-color:#2563eb; font-weight:600 }
  #cktab-cophieu:checked ~ .cktypetabs-panels .cktypepanel-cophieu,
  #cktab-ccq:checked ~ .cktypetabs-panels .cktypepanel-ccq,
  #cktab-tinphieu:checked ~ .cktypetabs-panels .cktypepanel-tinphieu,
  #cktab-traiphieu:checked ~ .cktypetabs-panels .cktypepanel-traiphieu,
  #cktab-cq:checked ~ .cktypetabs-panels .cktypepanel-cq { display:block }

  #ptab-ttchung:checked ~ .tabnav label[for=ptab-ttchung],
  #ptab-ck:checked ~ .tabnav label[for=ptab-ck],
  #ptab-ks:checked ~ .tabnav label[for=ptab-ks],
  #ptab-deleted:checked ~ .tabnav label[for=ptab-deleted],
  #ptab-oldcodes:checked ~ .tabnav label[for=ptab-oldcodes] { color:#2563eb; border-bottom-color:#2563eb; font-weight:600 }
  #ptab-ttchung:checked ~ .panels .panel-ttchung,
  #ptab-ck:checked ~ .panels .panel-ck,
  #ptab-ks:checked ~ .panels .panel-ks,
  #ptab-deleted:checked ~ .panels .panel-deleted,
  #ptab-oldcodes:checked ~ .panels .panel-oldcodes { display:block }

  .muted { color:#999 }
  h3 { font-size:13px; color:#666; margin:20px 0 8px }
  .searchbox { margin-bottom:14px }
  .searchbox input[type=text] { padding:7px 10px; border:1px solid #ccc; border-radius:5px; font-size:13px; width:260px }
  .ks-filter { display:flex; gap:8px; align-items:center; margin-bottom:14px }
  .ks-filter input[type=text] { padding:7px 10px; border:1px solid #ccc; border-radius:5px; font-size:13px; width:180px }
  .ks-filter select { padding:7px 10px; border:1px solid #ccc; border-radius:5px; font-size:13px }
  .minibtn { display:inline-block; background:#f0f0f0; color:#2563eb; border:1px solid #ddd; padding:3px 9px; border-radius:4px; font-size:12px; text-decoration:none; margin-right:4px }
  .minibtn:hover { background:#e5e7eb }
  .minibtn-danger { color:#c0392b }
  .editbox { background:#fff; border:1px solid #e5e7eb; border-radius:6px; padding:14px 18px; margin-bottom:16px }
  .closebar { text-align:right; margin-bottom:8px }
  .closebar .minibtn { background:#eafbea; color:#1a7a1a; border-color:#bfe8bf }
  .editnav { display:flex; gap:4px; border-bottom:1px solid #e5e7eb; margin-bottom:14px }
  .editnav span { padding:6px 14px; font-size:12px }
  .editnav-active { color:#2563eb; font-weight:600; border-bottom:2px solid #2563eb }
  .editnav-disabled { color:#bbb }
  .formrow { display:flex; align-items:center; margin-bottom:8px }
  .formrow label { width:160px; font-size:12px; color:#555 }
  .formrow input, .formrow select { flex:1; max-width:280px; padding:6px 8px; border:1px solid #ccc; border-radius:4px; font-size:13px }
  .formrow input:disabled { background:#f5f5f5; color:#888 }
</style>
</head>
<body>
<div class="topbar">
  <h1>Flex (Mock) - Màn hình 020004 giả lập</h1>
  <a class="logout" href="/logout">Đăng xuất ($($authCfg.Username))</a>
</div>
<div class="note">Đây là Flex GIẢ LẬP dùng để demo. Mỗi tab là 1 danh sách mã đang ở đúng giai đoạn đó. Bấm "Duyệt" sẽ chuyển mã sang tab kế tiếp và gửi mail.</div>

<div class="pagetabs">
  <input type="radio" id="ptab-ttchung" name="ptabs" $ttChungChecked>
  <input type="radio" id="ptab-ck" name="ptabs" $ckChecked>
  <input type="radio" id="ptab-ks" name="ptabs" $ksChecked>
  <input type="radio" id="ptab-deleted" name="ptabs" $deletedChecked>
  <input type="radio" id="ptab-oldcodes" name="ptabs" $oldCodesChecked>
  <div class="tabnav">
    <label for="ptab-ttchung">TT chung <span class="badge-count">$($countTT + $countNameChange)</span></label>
    <label for="ptab-ck">Chứng khoán <span class="badge-count">$countCK</span></label>
    <label for="ptab-ks">D/S Kiểm soát</label>
    <label for="ptab-deleted">Đã xóa <span class="badge-count">$countDeleted</span></label>
    <label for="ptab-oldcodes">Mã cũ <span class="badge-count">$countOldBacklog</span></label>
  </div>
  <div class="panels">
    <div class="panel panel-ttchung">
      <h3>Mã mới / mã cũ đang chờ duyệt TT chung ($countTT)</h3>
      <table>
        <tr><th>Tên viết tắt</th><th>Tổ chức phát hành</th><th>Tên giao dịch</th><th>Tên giao dịch tiếng Anh</th><th>Tên tiếng Anh</th><th>Thị trường</th><th>Nơi quản lý tại VSD</th><th>Mã ISIN</th><th>Nguồn</th><th></th></tr>
        $ttChungRows
      </table>

      <h3>Mã cần cập nhật tên (mục 4 CR) ($countNameChange)</h3>
      <table>
        <tr><th>Mã CK</th><th>Tên cũ (Flex)</th><th>Tên mới (VSD)</th><th></th></tr>
        $nameChangeRows
      </table>
    </div>
    <div class="panel panel-ck">
      <div class="searchbox">
        <form method="get" action="/">
          <input type="text" name="ck_search" placeholder="Nhập mã chứng khoán cần tìm..." value="$(Enc $CkSearchQuery)" />
          <button type="submit" class="btn">Tìm</button>
        </form>
      </div>
      $ckSearchBlock

      <div class="cktypetabs">
        <input type="radio" id="cktab-cophieu" name="cktypetabs" checked>
        <input type="radio" id="cktab-ccq" name="cktypetabs">
        <input type="radio" id="cktab-tinphieu" name="cktypetabs">
        <input type="radio" id="cktab-traiphieu" name="cktypetabs">
        <input type="radio" id="cktab-cq" name="cktypetabs">
        <div class="cktypetabs-nav">
          <label for="cktab-cophieu">Cổ phiếu <span class="badge-count">$countCkCoPhieu</span></label>
          <label for="cktab-ccq">Chứng chỉ quỹ <span class="badge-count">$countCkChungChiQuy</span></label>
          <label for="cktab-tinphieu">Tín phiếu <span class="badge-count">$countCkTinPhieu</span></label>
          <label for="cktab-traiphieu">Trái phiếu <span class="badge-count">$countCkTraiPhieu</span></label>
          <label for="cktab-cq">Chứng quyền <span class="badge-count">$countCkChungQuyen</span></label>
        </div>
        <div class="cktypetabs-panels">
          <div class="cktypepanel cktypepanel-cophieu">
            <table>
              <tr><th>Mã CK</th><th>Tên</th><th>Nơi GD</th><th>Loại chứng khoán</th><th>Loại trái phiếu</th><th>Có thu phí lưu ký</th><th>Mệnh giá</th><th>Nguồn</th><th></th></tr>
              $ckCoPhieuRows
            </table>
          </div>
          <div class="cktypepanel cktypepanel-ccq">
            <table>
              <tr><th>Mã CK</th><th>Tên</th><th>Nơi GD</th><th>Loại chứng khoán</th><th>Loại trái phiếu</th><th>Có thu phí lưu ký</th><th>Mệnh giá</th><th>Nguồn</th><th></th></tr>
              $ckChungChiQuyRows
            </table>
          </div>
          <div class="cktypepanel cktypepanel-tinphieu">
            <table>
              <tr><th>Mã CK</th><th>Tên</th><th>Nơi GD</th><th>Loại chứng khoán</th><th>Loại trái phiếu</th><th>Có thu phí lưu ký</th><th>Mệnh giá</th><th>Nguồn</th><th></th></tr>
              $ckTinPhieuRows
            </table>
          </div>
          <div class="cktypepanel cktypepanel-traiphieu">
            <table>
              <tr><th>Mã CK</th><th>Tên</th><th>Nơi GD</th><th>Loại chứng khoán</th><th>Loại trái phiếu</th><th>Có thu phí lưu ký</th><th>Mệnh giá</th><th>Loại kỳ hạn</th><th>Kỳ hạn</th><th>Nguồn</th><th></th></tr>
              $ckTraiPhieuRows
            </table>
          </div>
          <div class="cktypepanel cktypepanel-cq">
            <div style="overflow-x:auto">
            <table>
              <tr><th>Mã CK</th><th>Tên</th><th>Nơi GD</th><th>Loại CK cơ sở</th><th>Mã CKCS</th><th>Tên TCPH CKCS</th><th>Loại chứng quyền</th><th>Phương thức thanh toán</th><th>Giá thanh toán</th><th>Giá thực hiện</th><th>Tỷ lệ chuyển đổi</th><th>Thời hạn</th><th>Ngày đáo hạn</th><th>Ngày GD cuối cùng</th><th>Nguồn</th><th></th></tr>
              $ckChungQuyenRows
            </table>
            </div>
          </div>
        </div>
      </div>
    </div>
    <div class="panel panel-ks">
      $ksFilterHtml
      <table>
        <tr><th>Thời gian</th><th>Mã CK</th><th>Tên</th><th>Hành động</th><th>Trạng thái</th></tr>
        $kiemSoatRows
      </table>
    </div>
    <div class="panel panel-deleted">
      <div class="note">Mã đã xóa được giữ tối đa <b>$DeleteRetentionDays ngày</b> để khôi phục. Sau thời hạn này sẽ bị xóa vĩnh viễn.</div>
      <table>
        <tr><th>Mã CK</th><th>Tên</th><th>Thời điểm xóa</th><th>Trạng thái trước khi xóa</th><th>Còn lại</th><th></th></tr>
        $deletedRows
      </table>
    </div>
    <div class="panel panel-oldcodes">
      <div class="note">Mục 3 (CR): mã chứng khoán cũ tồn đọng từ trước khi phần mềm chạy. Nhập số lượng xử lý mỗi lần, phần mềm sẽ nhập dần vào tab TT chung theo lô.</div>
      $oldCodesBlock
    </div>
  </div>
</div>

</body></html>
"@
}

function Write-HtmlResponse {
    param($Response, [string]$Html, [int]$StatusCode = 200)
    $Response.StatusCode = $StatusCode
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($Html)
    $Response.ContentType = "text/html; charset=utf-8"
    $Response.ContentLength64 = $buffer.Length
    $Response.OutputStream.Write($buffer, 0, $buffer.Length)
    $Response.Close()
}

function Redirect-To {
    param($Response, [string]$Location, [string]$SetCookie = $null)
    if ($SetCookie) { $Response.Headers.Add("Set-Cookie", $SetCookie) }
    $Response.StatusCode = 302
    $Response.RedirectLocation = $Location
    $Response.Close()
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Flex Mock UI dang chay: http://localhost:$Port/"
Write-Host "Tai khoan demo: $($authCfg.Username) / $($authCfg.Password)"
Write-Host "Nhan Ctrl+C de dung."

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        try {
            $path = $request.Url.AbsolutePath

            if ($path -eq "/login" -and $request.HttpMethod -eq "GET") {
                Write-HtmlResponse -Response $response -Html (Build-LoginHtml)
                continue
            }

            if ($path -eq "/login" -and $request.HttpMethod -eq "POST") {
                $form = Parse-FormBody -Request $request
                if ($form["username"] -eq $authCfg.Username -and $form["password"] -eq $authCfg.Password) {
                    $token = [guid]::NewGuid().ToString("N")
                    $global:ValidSessions.Add($token) | Out-Null
                    Redirect-To -Response $response -Location "/" -SetCookie "session=$token; Path=/; HttpOnly"
                } else {
                    Write-HtmlResponse -Response $response -Html (Build-LoginHtml -ShowError $true) -StatusCode 401
                }
                continue
            }

            if ($path -eq "/logout") {
                $token = Get-SessionToken -Request $request
                if ($token) { $global:ValidSessions.Remove($token) | Out-Null }
                Redirect-To -Response $response -Location "/login" -SetCookie "session=; Path=/; Expires=Thu, 01 Jan 1970 00:00:00 GMT"
                continue
            }

            if (-not (Test-Authenticated -Request $request)) {
                Redirect-To -Response $response -Location "/login"
                continue
            }

            if ($request.HttpMethod -eq "POST" -and $path -eq "/approve") {
                $code = $request.QueryString["code"]
                $flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)
                $item = $flex | Where-Object { $_.Code -eq $code } | Select-Object -First 1
                if ($item) {
                    $oldStatus = $item.Status
                    $newStatus = Get-NextStatus -CurrentStatus $oldStatus
                    $hanhDong = "Nhân viên bấm Duyệt trên màn hình (mock)"

                    if ($oldStatus -eq "Chờ duyệt TT chung" -and $item.Tabs -and $item.Tabs.ChungKhoan) {
                        # Muc 4 (doi ten): ma da hoat dong san tren Flex, tab Chung khoan da co du lieu
                        # tu truoc -> duyet 1 lan la xong, khong can qua lai tab Chung khoan.
                        $newStatus = "Hoạt động"
                        $hanhDong = "Duyệt đổi tên xong - hoàn tất (mã đã hoạt động sẵn, không qua tab Chứng khoán)"
                    } elseif ($oldStatus -eq "Chờ duyệt TT chung") {
                        # Chuyen tu tab TT chung sang tab Chung khoan: nhap tab Chung khoan
                        if (-not $item.Tabs) { $item | Add-Member -NotePropertyName Tabs -NotePropertyValue ([pscustomobject]@{ TTChung=$null; ChungKhoan=$null; TTCK=$null }) -Force }
                        $item.Tabs.ChungKhoan = New-ChungKhoanTabData -Code $item.Code -Market $item.Market -StockType $item.StockType -MenhGiaVSD $item.MenhGiaVSD
                        $hanhDong = "Duyệt TT chung xong -> chuyển sang tab Chứng khoán"
                    } elseif ($oldStatus -eq "Chờ duyệt Chứng khoán") {
                        # Hoan tat: dien tab TTCK ngam (khong hien thi tab rieng theo yeu cau UI)
                        if ($item.Tabs) {
                            $ttck = [pscustomobject]@{
                                DonViGiaoDich         = "1000"
                                KhoiLuongNiemYet      = if ($item.TongSoDangKyVSD) { $item.TongSoDangKyVSD } else { "N/A (theo tổng số đăng ký tại VSD)" }
                                NgayNiemYet           = (Get-Date -Format "yyyy-MM-dd")
                                MuaBanCungNgay        = "Có"
                                CheckRoomNDTNuocNgoai = "Có"
                            }
                            if ($item.Tabs.PSObject.Properties['TTCK']) { $item.Tabs.TTCK = $ttck }
                            else { $item.Tabs | Add-Member -NotePropertyName TTCK -NotePropertyValue $ttck -Force }
                        }
                        $hanhDong = "Duyệt Chứng khoán xong -> hoàn tất"
                    }

                    Set-FlexStatus -Item $item -NewStatus $newStatus -HanhDong $hanhDong
                    Save-FlexStore -Path $FlexStorePath -Data $flex
                    Send-ApproveNotice -Code $code -NewStatus $newStatus -HanhDong $hanhDong
                }
                Redirect-To -Response $response -Location "/"
                continue
            }

            if ($request.HttpMethod -eq "POST" -and $path -eq "/update-security") {
                # Man hinh Sua cho sua duoc TAT CA truong TRU "Trang thai CK" (truong do
                # do luong Duyet dieu khien, khong nhan tu form nay).
                $code = $request.QueryString["code"]
                $form = Parse-FormBody -Request $request
                $flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)
                $item = $flex | Where-Object { $_.Code -eq $code } | Select-Object -First 1
                $redirectCode = $code

                if ($item) {
                    # Gom TAT CA thay doi thuc su (gia tri moi khac gia tri cu) vao 1 danh
                    # sach duy nhat, de gui 1 mail duy nhat liet ke day du - khong chi rieng
                    # Noi GD nhu truoc.
                    $changes = New-Object System.Collections.Generic.List[string]
                    $marketChanged = $false

                    # --- Noi GD: giu nguyen dung logic "muc 2 CR - xu ly ma chuyen san" ---
                    $newMarket = $form["noigd"]
                    if ($newMarket -and $newMarket -ne $item.Market) {
                        $oldMarket = $item.Market
                        $item | Add-Member -NotePropertyName TraCuuChuyenSan -NotePropertyValue ([pscustomobject]@{
                            NoiGDCu   = $oldMarket
                            NoiGDMoi  = $newMarket
                            TraCuuLuc = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                        }) -Force
                        $item.Market = $newMarket
                        $changes.Add("Nơi GD: $oldMarket -> $newMarket")
                        $marketChanged = $true
                    }

                    # --- Cac truong sua truc tiep, khong keo theo luong duyet ---
                    $oldManagementArea = $item.ManagementArea
                    if ($form["noiluuky"] -and $form["noiluuky"] -ne $oldManagementArea) {
                        $item | Add-Member -NotePropertyName ManagementArea -NotePropertyValue $form["noiluuky"] -Force
                        $changes.Add("Nơi lưu ký: $oldManagementArea -> $($form["noiluuky"])")
                    }
                    $oldStockType = $item.StockType
                    if ($form["loaick"] -and $form["loaick"] -ne $oldStockType) {
                        $item | Add-Member -NotePropertyName StockType -NotePropertyValue $form["loaick"] -Force
                        $changes.Add("Loại chứng khoán: $oldStockType -> $($form["loaick"])")
                    }
                    if ($form["loaitraiphieu"] -and $item.Tabs -and $item.Tabs.ChungKhoan -and $form["loaitraiphieu"] -ne $item.Tabs.ChungKhoan.LoaiTraiPhieu) {
                        $oldLoaiTP = $item.Tabs.ChungKhoan.LoaiTraiPhieu
                        $item.Tabs.ChungKhoan.LoaiTraiPhieu = $form["loaitraiphieu"]
                        $changes.Add("Loại trái phiếu: $oldLoaiTP -> $($form["loaitraiphieu"])")
                    }
                    $oldMaQuiUoc = if ($item.MaQuiUocOverride) { $item.MaQuiUocOverride } else { Get-SyntheticId -Code $item.Code -Seed 17 }
                    if ($form["maquiuoc"] -and $form["maquiuoc"] -ne $oldMaQuiUoc) {
                        $item | Add-Member -NotePropertyName MaQuiUocOverride -NotePropertyValue $form["maquiuoc"] -Force
                        $changes.Add("Mã qui ước: $oldMaQuiUoc -> $($form["maquiuoc"])")
                    }
                    $oldMaTcph = if ($item.MaTCPHOverride) { $item.MaTCPHOverride } else { Get-SyntheticId -Code $item.Code -Seed 31 }
                    if ($form["matcph"] -and $form["matcph"] -ne $oldMaTcph) {
                        $item | Add-Member -NotePropertyName MaTCPHOverride -NotePropertyValue $form["matcph"] -Force
                        $changes.Add("Mã TCPH: $oldMaTcph -> $($form["matcph"])")
                    }

                    # --- Sua Ma chung khoan (doi ten khoa chinh) - lam SAU CUNG, kiem tra trung ---
                    $newCode = $form["code"]
                    if ($newCode -and $newCode -ne $item.Code) {
                        $conflict = $flex | Where-Object { $_.Code -eq $newCode }
                        if (-not $conflict) {
                            $oldCode = $item.Code
                            $item.Code = $newCode
                            if ($item.Tabs -and $item.Tabs.TTChung -and $item.Tabs.TTChung.MaCK) { $item.Tabs.TTChung.MaCK = $newCode }
                            if ($item.Tabs -and $item.Tabs.ChungKhoan -and $item.Tabs.ChungKhoan.MaCK) { $item.Tabs.ChungKhoan.MaCK = $newCode }
                            $changes.Add("Mã chứng khoán: $oldCode -> $newCode")
                            $redirectCode = $newCode
                        } else {
                            Write-Warning "Ma '$newCode' da ton tai - bo qua doi ten, cac truong khac van duoc luu."
                        }
                    }

                    if ($changes.Count -gt 0) {
                        $hanhDong = "Nhân viên sửa qua màn hình Sửa:`n$($changes -join "`n")"
                        if ($marketChanged) {
                            # Muc 2 CR: doi Noi GD phai quay lai cho duyet tab Chung khoan
                            Set-FlexStatus -Item $item -NewStatus "Chờ duyệt Chứng khoán" -HanhDong $hanhDong
                        } else {
                            # Cac truong khac: chi ghi log, KHONG doi trang thai/luong duyet
                            $entry = [pscustomobject]@{ ThoiGian = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); HanhDong = $hanhDong; TrangThai = $item.Status }
                            $existing = @(if ($item.LichSuDuyet) { $item.LichSuDuyet } else { @() })
                            $newHistory = $existing + @($entry)
                            if ($item.PSObject.Properties['LichSuDuyet']) { $item.LichSuDuyet = $newHistory }
                            else { $item | Add-Member -NotePropertyName LichSuDuyet -NotePropertyValue $newHistory -Force }
                        }
                        Save-FlexStore -Path $FlexStorePath -Data $flex
                        Send-ApproveNotice -Code $redirectCode -NewStatus $item.Status -HanhDong $hanhDong
                    }
                }
                Redirect-To -Response $response -Location "/view-security?code=$([System.Uri]::EscapeDataString($redirectCode))&q=$([System.Uri]::EscapeDataString($redirectCode))"
                continue
            }

            if ($request.HttpMethod -eq "POST" -and $path -eq "/delete-security") {
                # Xoa MEM (soft delete) - chuyen vao tab "Da xoa", giu $DeleteRetentionDays
                # ngay de khoi phuc. Qua han se bi xoa vinh vien (xem Purge-DeletedSecurities.ps1).
                $code = $request.QueryString["code"]
                $form = Parse-FormBody -Request $request
                $q = $form["q"]
                $flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)
                $item = $flex | Where-Object { $_.Code -eq $code } | Select-Object -First 1
                if ($item) {
                    $item | Add-Member -NotePropertyName StatusBeforeDelete -NotePropertyValue $item.Status -Force
                    $item | Add-Member -NotePropertyName DeletedAt -NotePropertyValue (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") -Force
                    Set-FlexStatus -Item $item -NewStatus "Đã xóa" -HanhDong "Nhan vien xoa ma nay qua man hinh tim kiem (giu $DeleteRetentionDays ngay de khoi phuc)"
                    Save-FlexStore -Path $FlexStorePath -Data $flex
                    Send-ApproveNotice -Code $code -NewStatus "Đã xóa" -HanhDong "Ma da bi xoa (mem), co the khoi phuc trong $DeleteRetentionDays ngay tai tab Da xoa"
                }
                Redirect-To -Response $response -Location "/?ck_search=$([System.Uri]::EscapeDataString($q))"
                continue
            }

            if ($request.HttpMethod -eq "POST" -and $path -eq "/restore-security") {
                $code = $request.QueryString["code"]
                $flex = [System.Collections.Generic.List[object]](Get-FlexStore -Path $FlexStorePath)
                $item = $flex | Where-Object { $_.Code -eq $code } | Select-Object -First 1
                if ($item -and $item.Status -eq "Đã xóa") {
                    $restoreStatus = if ($item.StatusBeforeDelete) { $item.StatusBeforeDelete } else { "Hoạt động" }
                    Set-FlexStatus -Item $item -NewStatus $restoreStatus -HanhDong "Nhan vien khoi phuc ma tu tab Da xoa"
                    $item.PSObject.Properties.Remove("DeletedAt")
                    $item.PSObject.Properties.Remove("StatusBeforeDelete")
                    Save-FlexStore -Path $FlexStorePath -Data $flex
                    Send-ApproveNotice -Code $code -NewStatus $restoreStatus -HanhDong "Da khoi phuc ma nay tu tab Da xoa"
                }
                Redirect-To -Response $response -Location "/?tab=deleted"
                continue
            }

            if ($request.HttpMethod -eq "POST" -and $path -eq "/old-codes/init") {
                $form = Parse-FormBody -Request $request
                $batchSize = [int]$form["batchSize"]
                $newCodes = Get-LatestDiffNewCodes
                Save-OldCodesBacklog -Backlog $newCodes
                Save-OldCodesConfig -BatchSize $batchSize
                Redirect-To -Response $response -Location "/?tab=oldcodes"
                continue
            }

            if ($request.HttpMethod -eq "POST" -and $path -eq "/old-codes/set-batch-size") {
                $form = Parse-FormBody -Request $request
                $batchSize = [int]$form["batchSize"]
                Save-OldCodesConfig -BatchSize $batchSize
                Redirect-To -Response $response -Location "/?tab=oldcodes"
                continue
            }

            if ($request.HttpMethod -eq "POST" -and $path -eq "/old-codes/process-batch") {
                Invoke-OldCodesBatch | Out-Null
                Redirect-To -Response $response -Location "/?tab=ttchung"
                continue
            }

            if ($path -eq "/view-security") {
                $code = $request.QueryString["code"]
                $q = $request.QueryString["q"]
                $flex = Get-FlexStore -Path $FlexStorePath
                $item = $flex | Where-Object { $_.Code -eq $code } | Select-Object -First 1
                $body = Build-CkViewHtml -Item $item -CloseHref "/?ck_search=$([System.Uri]::EscapeDataString($q))"
                Write-HtmlResponse -Response $response -Html (Build-StandalonePage -Title "Xem $code - Flex (Mock)" -BodyHtml $body)
                continue
            }

            if ($path -eq "/edit-security") {
                $code = $request.QueryString["code"]
                $q = $request.QueryString["q"]
                $flex = Get-FlexStore -Path $FlexStorePath
                $item = $flex | Where-Object { $_.Code -eq $code } | Select-Object -First 1
                $body = Build-CkEditFormHtml -Item $item -CloseHref "/?ck_search=$([System.Uri]::EscapeDataString($q))"
                Write-HtmlResponse -Response $response -Html (Build-StandalonePage -Title "Sửa $code - Flex (Mock)" -BodyHtml $body)
                continue
            }

            if ($path -eq "/") {
                $ckSearch = $request.QueryString["ck_search"]
                $editCode = $request.QueryString["edit"]
                $viewCode = $request.QueryString["view"]
                $viewFrom = $request.QueryString["from"]
                $activeTab = $request.QueryString["tab"]
                $ksCode = $request.QueryString["ks_code"]
                $ksStatus = $request.QueryString["ks_status"]
                Write-HtmlResponse -Response $response -Html (Build-Html -CkSearchQuery $ckSearch -EditCode $editCode -ViewCode $viewCode -ViewFrom $viewFrom -ActiveTab $activeTab -KsCodeFilter $ksCode -KsStatusFilter $ksStatus)
                continue
            }

            $response.StatusCode = 404
            $response.Close()
        } catch {
            Write-Warning "Loi xu ly request: $_"
            try {
                $errHtml = "<html><body style='font-family:sans-serif;padding:30px'><h2>Loi server (mock)</h2><pre style='white-space:pre-wrap;color:#c0392b'>$([System.Net.WebUtility]::HtmlEncode($_.ToString()))</pre><p><a href='/'>Ve trang chu</a></p></body></html>"
                Write-HtmlResponse -Response $response -Html $errHtml -StatusCode 500
            } catch {
                try { $response.Close() } catch {}
            }
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
