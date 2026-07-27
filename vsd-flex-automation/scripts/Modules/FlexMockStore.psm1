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

Export-ModuleMember -Function Get-FlexStore, Save-FlexStore, Set-FlexStatus
