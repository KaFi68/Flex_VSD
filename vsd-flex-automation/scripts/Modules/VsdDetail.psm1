<#
Lay du lieu tu trang CHI TIET tung ma chung khoan tren vsd.vn (https://vsd.vn/s-detail/{id}).
Trang nay KHONG can dang nhap/token (khac trang danh sach), va co day du cac truong ma
trang danh sach KHONG co:
  - Ten To chuc dang ky chung khoan (Ten TCPH)  - KHAC voi Ten chung khoan
  - Menh gia
  - Tong so chung khoan dang ky (dung cho Khoi luong niem yet o tab TTCK)
  - Noi quan ly tai VSDC (doi chieu voi ManagementArea da scrape o trang danh sach)

DetailId lay tu link <a href="/s-detail/{id}">Ma CK</a> o trang danh sach (xem
Fetch-VsdSecurities.ps1 - truong DetailId).
#>

function Get-VsdSecurityDetail {
    param([Parameter(Mandatory)][string]$DetailId)

    try {
        $resp = Invoke-WebRequest -Uri "https://vsd.vn/s-detail/$DetailId" -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" -UseBasicParsing -TimeoutSec 15
    } catch {
        Write-Warning "Khong fetch duoc trang chi tiet /s-detail/$DetailId : $_"
        return $null
    }

    # Trang nay ghi nhan tieng Viet o dang HTML entity (VD: "T&#xEA;n") ngay ca trong the
    # <div>, nen phai giai ma ENTITY CHO CA TRANG truoc khi tim nhan bang tieng Viet thuong.
    $html = [System.Net.WebUtility]::HtmlDecode($resp.Content)

    function Get-FieldValue {
        param([string]$Html, [string]$Label)
        $pattern = [regex]::Escape($Label) + '\s*(?:<span[^>]*>.*?</span>)?\s*:</div>\s*<div[^>]*>\s*(?:<a[^>]*>)?\s*(.*?)\s*(?:</a>)?\s*</div>'
        $m = [regex]::Match($Html, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if (-not $m.Success) { return $null }
        $raw = [regex]::Replace($m.Groups[1].Value, '<[^>]+>', '').Trim()
        return $raw
    }

    # "Ky han" chi xuat hien o trang chi tiet TRAI PHIEU, dang cau van tu do vi du
    # "04 năm kể từ ngày phát hành" hoac "2520 ngày" - can tach so + don vi de dien vao
    # 2 truong rieng "Loai ky han" va "Ky han" ben Flex. Flex CHI cho chon Tuan/Thang/Nam
    # (khong co "ngay"), nen neu VSD tra ve don vi "ngay" thi quy doi sang NAM (chia 360,
    # lam tron - quy uoc theo yeu cau nghiep vu).
    $kyHanRaw = Get-FieldValue -Html $html -Label "Kỳ hạn"
    $loaiKyHan = $null
    $kyHan = $null
    if ($kyHanRaw) {
        $kyHanMatch = [regex]::Match($kyHanRaw, '(\d+)\s*(tuần|tháng|năm)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($kyHanMatch.Success) {
            $kyHan = [string]([int]$kyHanMatch.Groups[1].Value)
            $donVi = $kyHanMatch.Groups[2].Value.ToLower()
            $loaiKyHan = (Get-Culture).TextInfo.ToTitleCase($donVi)
        } else {
            $ngayMatch = [regex]::Match($kyHanRaw, '(\d+)\s*ngày', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($ngayMatch.Success) {
                $soNgay = [int]$ngayMatch.Groups[1].Value
                $kyHan = [string][math]::Round($soNgay / 360.0)
                $loaiKyHan = "Năm"
            }
        }
    }

    # "Thoi han" chi xuat hien o trang chi tiet CHUNG QUYEN (khac nhan "Ky han" cua Trai
    # phieu), vd "4 tháng". Flex chi co 1 truong "Thoi han CW theo thang" (luon quy ve
    # THANG), nen quy doi moi don vi ve thang: nam*12, ngay/30, tuan*7/30.
    $thoiHanRaw = Get-FieldValue -Html $html -Label "Thời hạn"
    $thoiHanThang = $null
    if ($thoiHanRaw) {
        $thMatch = [regex]::Match($thoiHanRaw, '(\d+)\s*(tuần|tháng|năm|ngày)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($thMatch.Success) {
            $soLuong = [double]$thMatch.Groups[1].Value
            $donViTH = $thMatch.Groups[2].Value.ToLower()
            $soThang = switch ($donViTH) {
                "tháng" { $soLuong }
                "năm"   { $soLuong * 12 }
                "ngày"  { $soLuong / 30.0 }
                "tuần"  { $soLuong * 7.0 / 30.0 }
            }
            $thoiHanThang = [string]([math]::Round($soThang))
        }
    }

    # "Ty le chuyen doi" ben VSD ghi dang "4:1", ben Flex ghi dang "4/1" (theo vi du trong
    # bang mapping) - chi doi dau ":" thanh "/".
    $tyLeChuyenDoiVSD = Get-FieldValue -Html $html -Label "Tỷ lệ chuyển đổi"
    $tyLeChuyenDoi = if ($tyLeChuyenDoiVSD) { $tyLeChuyenDoiVSD -replace ':', '/' } else { $null }

    return [pscustomobject]@{
        TenTCPH           = Get-FieldValue -Html $html -Label "Tên Tổ chức đăng ký chứng khoán"
        TenChungKhoan     = Get-FieldValue -Html $html -Label "Tên chứng khoán"
        LoaiChungKhoanVSD = Get-FieldValue -Html $html -Label "Loại chứng khoán"
        MenhGia           = Get-FieldValue -Html $html -Label "Mệnh giá"
        TongSoDangKy      = Get-FieldValue -Html $html -Label "Tổng số chứng khoán đăng ký"
        NoiQuanLyVSDC     = Get-FieldValue -Html $html -Label "Nơi quản lý tại VSDC"
        KyHanRaw          = $kyHanRaw
        LoaiKyHan         = $loaiKyHan
        KyHan             = $kyHan
        # --- rieng cho Chung quyen ---
        MaCKCS            = Get-FieldValue -Html $html -Label "Mã chứng khoán cơ sở"
        TenTCPHCKCS       = Get-FieldValue -Html $html -Label "Tổ chức phát hành mã chứng khoán cơ sở"
        LoaiChungQuyen    = Get-FieldValue -Html $html -Label "Loại chứng quyền"
        PhuongThucThanhToan = Get-FieldValue -Html $html -Label "Phương thức thực hiện chứng quyền"
        GiaThucHien       = Get-FieldValue -Html $html -Label "Giá thực hiện"
        TyLeChuyenDoi     = $tyLeChuyenDoi
        ThoiHanCWThang    = $thoiHanThang
        NgayDaoHan        = Get-FieldValue -Html $html -Label "Ngày đáo hạn"
    }
}

Export-ModuleMember -Function Get-VsdSecurityDetail
