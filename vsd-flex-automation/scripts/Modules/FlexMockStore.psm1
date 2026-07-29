<#
Flex GIA LAP (mock). Day KHONG PHAI ket noi that toi he thong Flex - day la mot file
JSON dai dien cho "du lieu ma chung khoan hien co trong Flex", dung de demo dung logic
nghiep vu (phat hien ma moi / chuyen san / doi ten + luong cho duyet) trong khi cho
thong tin ket noi Flex that (API/DB/RPA).

Khi co Flex that: thay Get-FlexStore / Save-FlexStore ben duoi bang ham doc/ghi Flex
that (vi du goi API, hoac doc bang trong DB Flex). Phan logic nghiep vu o cac script
khac (Process-NewSecurities.ps1, Process-MarketTransfer.ps1...) khong can sua vi
chung chi lam viec voi doi tuong PowerShell, khong quan tam du lieu tu dau ra.
#>

function Get-FlexStore {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) {
        throw "Chua co Flex mock store tai $Path. Chay Seed-FlexMock.ps1 truoc."
    }
    return [object[]](Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Save-FlexStore {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)]$Data)
    $Data | ConvertTo-Json -Depth 5 | Out-File -FilePath $Path -Encoding utf8
}

function Set-FlexStatus {
    # LUU Y: PS 5.1 co 2 bug la lien quan array can tranh:
    #   1. Bam @() quanh 1 System.Collections.Generic.List[object] RONG se nem
    #      "Argument types do not match" -> TUYET DOI khong dung kieu List[], luon dung
    #      array thuong (@()).
    #   2. "$x = if (...) { @($y) } else { @() }" - neu $y la mang 1 phan tu, PowerShell
    #      "bung" ket qua cua if-block ve dung 1 object (khong con la mang) khi gan bien ->
    #      phai bam @(...) BAO NGOAI CA if/else, khong bam ben trong tung nhanh.
    param([Parameter(Mandatory)]$Item, [Parameter(Mandatory)][string]$NewStatus, [Parameter(Mandatory)][string]$HanhDong)
    $Item.Status = $NewStatus
    $Item.StatusChangedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = [pscustomobject]@{ ThoiGian = $Item.StatusChangedAt; HanhDong = $HanhDong; TrangThai = $NewStatus }
    $existing = @(if ($Item.LichSuDuyet) { $Item.LichSuDuyet } else { @() })
    $newHistory = $existing + @($entry)
    if ($Item.PSObject.Properties['LichSuDuyet']) {
        $Item.LichSuDuyet = $newHistory
    } else {
        $Item | Add-Member -NotePropertyName LichSuDuyet -NotePropertyValue $newHistory -Force
    }
}

function New-ChungKhoanTabData {
    # Xay dung du lieu tab "Chung khoan" theo dung bang mapping trong file "Cach thuc
    # khai ma chung khoan" (Excel dinh kem CR). MOI LOAI CHUNG KHOAN co bo truong rieng
    # (5 sheet trong Excel: Co phieu / Chung chi quy / Tin phieu / Trai phieu / Chung
    # quyen). Ca 5 loai deu da co day du rule.
    #
    # Rule "Co phieu" va "Chung chi quy" - GIONG HET nhau, chi khac "Loai chung khoan":
    #   Flex "Ma chung khoan"          <- VSD "Ma chung khoan"
    #   Flex "Noi GD"                  <- VSD "San giao dich"
    #   Flex "Loai chung khoan"        <- Co phieu: "Co phieu thuong" (thuong = bonus shares)
    #                                      Chung chi quy: "Chung chi quy"
    #   Flex "Loai trai phieu"         <- Luon chon "Khong phai trai phieu"
    #   Flex "Co thu phi luu ky khong" <- Luon chon "Co"
    #   Flex "Menh gia"                <- VSD "Menh gia"
    #
    # Rule "Tin phieu":
    #   Flex "Ma chung khoan"          <- VSD "Ma chung khoan"
    #   Flex "Noi GD"                  <- Luon chon "HNX" (KHONG lay tu VSD)
    #   Flex "Loai chung khoan"        <- Luon chon "Trai phieu"
    #   Flex "Loai trai phieu"         <- Luon chon "Tin phieu"
    #   Flex "Co thu phi luu ky khong" <- Luon chon "Co"
    #   Flex "Menh gia"                <- VSD "Menh gia"
    #
    # Rule "Trai phieu" (khac Tin phieu o cho Noi GD/Loai trai phieu/Ky han lay dong tu VSD,
    # KHONG fix cung "HNX"/"Tin phieu" nhu Tin phieu):
    #   Flex "Ma chung khoan"          <- VSD "Ma chung khoan"
    #   Flex "Noi GD"                  <- VSD "San giao dich" (chu y: neu la san DCCNY thi
    #                                      he thong Flex that se khong sinh tab Buoc gia/TTCK,
    #                                      chi thuc hien khai+duyet den Buoc 5 - mock nay chua
    #                                      mo phong rule do)
    #   Flex "Loai chung khoan"        <- Luon chon "Trai phieu"
    #   Flex "Loai trai phieu"         <- VSD "Loai chung khoan" (vd "Trai phieu doanh nghiep
    #                                      rieng le", "Trai phieu chinh phu"...)
    #   Flex "Co thu phi luu ky khong" <- Luon chon "Co"
    #   Flex "Menh gia"                <- VSD "Menh gia"
    #   Flex "Loai ky han"             <- tach tu VSD "Ky han" (Tuan/Thang/Nam)
    #   Flex "Ky han"                  <- tach tu VSD "Ky han" (phan so)
    #
    # Rule "Chung quyen":
    #   Flex "Ma chung khoan"          <- VSD "Ma chung quyen"
    #   Flex "Noi GD"                  <- VSD "Noi giao dich"
    #   Flex "Loai chung khoan"        <- Luon chon "Chung quyen"
    #   Flex "Loai trai phieu"         <- Luon chon "Khong phai trai phieu"
    #   Flex "Co thu phi luu ky khong" <- Luon chon "Co"
    #   Flex "Menh gia"                <- mac dinh 10.000 (VSD khong co truong nay cho CW)
    #   Flex "Loai chung khoan co so"  <- Luon chon "Co phieu"
    #   Flex "Ma CKCS"                 <- VSD "Ma chung khoan co so"
    #   Flex "Ten TCPH CKCS"           <- VSD "To chuc phat hanh ma chung khoan co so"
    #   Flex "Loai chung quyen"        <- VSD "Loai chung quyen" (Mua/Ban)
    #   Flex "Phuong thuc thanh toan"  <- VSD "Phuong thuc thuc hien chung quyen"
    #   Flex "Gia thanh toan"          <- mac dinh he thong de 0.0000
    #   Flex "Gia thuc hien"           <- VSD "Gia thuc hien"
    #   Flex "Ty le chuyen doi"        <- VSD "Ty le chuyen doi" (doi ":" -> "/")
    #   Flex "Thoi han CW theo thang"  <- tach tu VSD "Thoi han" (quy ve THANG)
    #   Flex "Ngay dao han"            <- VSD "Ngay dao han"
    #   Flex "Ngay giao dich cuoi cung"<- Ngay dao han - 2 nam (theo xac nhan nghiep vu)
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Market,
        [string]$StockType,
        [string]$MenhGiaVSD,
        [string]$LoaiTraiPhieuVSD,     # VSD "Loai chung khoan" tren trang chi tiet - chi dung cho Trai phieu
        [string]$LoaiKyHan,            # da tach san tu VsdDetail (Tuan/Thang/Nam) - chi dung cho Trai phieu
        [string]$KyHan,                # da tach san tu VsdDetail (phan so) - chi dung cho Trai phieu
        [string]$MaCKCS,               # chi dung cho Chung quyen
        [string]$TenTCPHCKCS,          # chi dung cho Chung quyen
        [string]$LoaiChungQuyen,       # chi dung cho Chung quyen
        [string]$PhuongThucThanhToan,  # chi dung cho Chung quyen
        [string]$GiaThucHien,          # chi dung cho Chung quyen
        [string]$TyLeChuyenDoi,        # da doi ":" -> "/" san tu VsdDetail - chi dung cho Chung quyen
        [string]$ThoiHanCWThang,       # da quy ve thang san tu VsdDetail - chi dung cho Chung quyen
        [string]$NgayDaoHan            # chi dung cho Chung quyen, dang dd/MM/yyyy
    )

    $noiGD = $Market
    $loaiChungKhoan = $StockType
    $loaiTraiPhieu = $null
    $coThuPhiLuuKy = $null
    $outLoaiKyHan = $null
    $outKyHan = $null
    $outLoaiCKCoSo = $null
    $outMaCKCS = $null
    $outTenTCPHCKCS = $null
    $outLoaiChungQuyen = $null
    $outPhuongThucThanhToan = $null
    $outGiaThanhToan = $null
    $outGiaThucHien = $null
    $outTyLeChuyenDoi = $null
    $outThoiHanCWThang = $null
    $outNgayDaoHan = $null
    $outNgayGDCuoiCung = $null

    if ($StockType -eq "Cổ phiếu") {
        $loaiChungKhoan = "Cổ phiếu thưởng"
        $loaiTraiPhieu = "Không phải trái phiếu"
        $coThuPhiLuuKy = "Có"
    }
    elseif ($StockType -like "*Chứng chỉ quỹ*") {
        $loaiChungKhoan = "Chứng chỉ quỹ"
        $loaiTraiPhieu = "Không phải trái phiếu"
        $coThuPhiLuuKy = "Có"
    }
    elseif ($StockType -like "Tín phiếu*") {
        $noiGD = "HNX"
        $loaiChungKhoan = "Trái phiếu"
        $loaiTraiPhieu = "Tín phiếu"
        $coThuPhiLuuKy = "Có"
    }
    elseif ($StockType -like "Trái phiếu*") {
        $loaiChungKhoan = "Trái phiếu"
        $loaiTraiPhieu = if ($LoaiTraiPhieuVSD) { $LoaiTraiPhieuVSD } else { $StockType }
        $coThuPhiLuuKy = "Có"
        $outLoaiKyHan = $LoaiKyHan
        $outKyHan = $KyHan
    }
    elseif ($StockType -like "*Chứng quyền*") {
        $loaiChungKhoan = "Chứng quyền"
        $loaiTraiPhieu = "Không phải trái phiếu"
        $coThuPhiLuuKy = "Có"
        $outLoaiCKCoSo = "Cổ phiếu"
        $outMaCKCS = $MaCKCS
        $outTenTCPHCKCS = $TenTCPHCKCS
        $outLoaiChungQuyen = $LoaiChungQuyen
        $outPhuongThucThanhToan = $PhuongThucThanhToan
        $outGiaThanhToan = "0.0000"
        $outGiaThucHien = $GiaThucHien
        $outTyLeChuyenDoi = $TyLeChuyenDoi
        $outThoiHanCWThang = $ThoiHanCWThang
        $outNgayDaoHan = $NgayDaoHan
        if ($NgayDaoHan) {
            try {
                $d = [datetime]::ParseExact($NgayDaoHan, "dd/MM/yyyy", $null)
                $outNgayGDCuoiCung = $d.AddYears(-2).ToString("dd/MM/yyyy")
            } catch { }
        }
    }

    return [pscustomobject]@{
        MaCK                = $Code
        NoiGD               = $noiGD
        LoaiChungKhoan      = $loaiChungKhoan
        LoaiTraiPhieu       = $loaiTraiPhieu
        CoThuPhiLuuKy       = $coThuPhiLuuKy
        MenhGia             = if ($MenhGiaVSD) { $MenhGiaVSD } else { "10000" }
        LoaiKyHan           = $outLoaiKyHan
        KyHan               = $outKyHan
        LoaiChungKhoanCoSo  = $outLoaiCKCoSo
        MaCKCS              = $outMaCKCS
        TenTCPHCKCS         = $outTenTCPHCKCS
        LoaiChungQuyen      = $outLoaiChungQuyen
        PhuongThucThanhToan = $outPhuongThucThanhToan
        GiaThanhToan        = $outGiaThanhToan
        GiaThucHien         = $outGiaThucHien
        TyLeChuyenDoi       = $outTyLeChuyenDoi
        ThoiHanCWThang      = $outThoiHanCWThang
        NgayDaoHan          = $outNgayDaoHan
        NgayGiaoDichCuoiCung = $outNgayGDCuoiCung
    }
}

# Nhom "loai chung khoan" VSD tra ve vao dung 1 trong 5 danh muc theo Excel (Co phieu /
# Chung chi quy / Tin phieu / Trai phieu / Chung quyen). Dung de chia sub-tab tren UI.
function Get-SecurityCategory {
    param([string]$StockType)
    switch -Wildcard ($StockType) {
        "Cổ phiếu*"        { return "CoPhieu" }
        "*Chứng chỉ quỹ*"  { return "ChungChiQuy" }
        "Tín phiếu*"       { return "TinPhieu" }
        "Trái phiếu*"      { return "TraiPhieu" }
        "*Chứng quyền*"    { return "ChungQuyen" }
        default            { return "CoPhieu" }   # mac dinh, phan lon du lieu la CP
    }
}

Export-ModuleMember -Function Get-FlexStore, Save-FlexStore, Set-FlexStatus, New-ChungKhoanTabData, Get-SecurityCategory
