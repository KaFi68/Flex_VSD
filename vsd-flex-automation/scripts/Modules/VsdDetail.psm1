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

    return [pscustomobject]@{
        TenTCPH         = Get-FieldValue -Html $html -Label "Tên Tổ chức đăng ký chứng khoán"
        TenChungKhoan   = Get-FieldValue -Html $html -Label "Tên chứng khoán"
        MenhGia         = Get-FieldValue -Html $html -Label "Mệnh giá"
        TongSoDangKy    = Get-FieldValue -Html $html -Label "Tổng số chứng khoán đăng ký"
        NoiQuanLyVSDC   = Get-FieldValue -Html $html -Label "Nơi quản lý tại VSDC"
    }
}

Export-ModuleMember -Function Get-VsdSecurityDetail
