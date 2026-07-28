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
    param([Parameter(Mandatory)]$Item, [Parameter(Mandatory)][string]$NewStatus, [Parameter(Mandatory)][string]$HanhDong)
    $Item.Status = $NewStatus
    $Item.StatusChangedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    if (-not $Item.LichSuDuyet) {
        $Item | Add-Member -NotePropertyName LichSuDuyet -NotePropertyValue (New-Object System.Collections.Generic.List[object]) -Force
    }
    $entry = [pscustomobject]@{ ThoiGian = $Item.StatusChangedAt; HanhDong = $HanhDong; TrangThai = $NewStatus }
    # LichSuDuyet co the la array thuong (sau khi doc tu JSON) chu khong phai List - dung + de cong dong bo
    $Item.LichSuDuyet = @($Item.LichSuDuyet) + @($entry)
}

function New-ChungKhoanTabData {
    # Xay dung du lieu tab "Chung khoan" theo dung bang mapping trong file "Cach thuc
    # khai ma chung khoan" (Excel dinh kem CR). MOI LOAI CHUNG KHOAN co bo truong rieng
    # (5 sheet trong Excel: Co phieu / Chung chi quy / Tin phieu / Trai phieu / Chung
    # quyen). HIEN co day du rule cho "Co phieu" va "Chung chi quy" - cac loai con lai
    # (Tin phieu / Trai phieu / Chung quyen) dung tam truong chung cho den khi co yeu cau.
    #
    # Rule "Co phieu" va "Chung chi quy" - GIONG HET nhau, chi khac "Loai chung khoan":
    #   Flex "Ma chung khoan"          <- VSD "Ma chung khoan"
    #   Flex "Noi GD"                  <- VSD "San giao dich"
    #   Flex "Loai chung khoan"        <- Co phieu: "Co phieu pho thong"
    #                                      Chung chi quy: "Chung chi quy"
    #   Flex "Loai trai phieu"         <- Luon chon "Khong phai trai phieu"
    #   Flex "Co thu phi luu ky khong" <- Luon chon "Co"
    #   Flex "Menh gia"                <- VSD "Menh gia"
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Market,
        [string]$StockType,
        [string]$MenhGiaVSD
    )

    $isCoPhieu = $StockType -eq "Cổ phiếu"
    $isChungChiQuy = $StockType -like "*Chứng chỉ quỹ*"

    $loaiChungKhoan = $StockType
    if ($isCoPhieu) { $loaiChungKhoan = "Cổ phiếu phổ thông" }
    elseif ($isChungChiQuy) { $loaiChungKhoan = "Chứng chỉ quỹ" }

    $apDungRuleChung = $isCoPhieu -or $isChungChiQuy

    return [pscustomobject]@{
        MaCK           = $Code
        NoiGD          = $Market
        LoaiChungKhoan = $loaiChungKhoan
        LoaiTraiPhieu  = if ($apDungRuleChung) { "Không phải trái phiếu" } else { $null }
        CoThuPhiLuuKy  = if ($apDungRuleChung) { "Có" } else { $null }
        MenhGia        = if ($MenhGiaVSD) { $MenhGiaVSD } else { "10000" }
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
