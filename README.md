# VSD-Flex Automation

Phần mềm tự động khai/duyệt mã chứng khoán mới từ VSD (vsd.vn) vào Flex, theo CR
*"Tạo phần mềm nhập tự động nhập mã chứng khoán mới vào Flex"*.

Gồm 4 nghiệp vụ:

1. **Khai mã chứng khoán mới** — mã mới trên VSD chưa có trên Flex, duyệt 2 bước (TT chung → Chứng khoán)
2. **Xử lý mã chuyển sàn** — phát hiện mã đã có trên Flex nhưng đổi Nơi GD
3. **Bổ sung mã cũ tồn đọng** — xử lý theo lô (batch), có cấu hình số lượng mỗi lần
4. **Cập nhật tên TCPH / tên giao dịch** — duyệt 1 bước

> ⚠️ Đây là bản có kèm **Flex giả lập (mock)** — vì chưa có quyền truy cập API/DB Flex
> thật. Toàn bộ business logic (mapping field, luồng duyệt...) đã đúng theo yêu cầu CR,
> chỉ cần thay phần đọc/ghi dữ liệu (`Modules/FlexMockStore.psm1`) bằng API/DB Flex thật
> khi có quyền truy cập.

## 1. Yêu cầu

- Windows PowerShell 5.1 (có sẵn trên Windows, không cần cài thêm)
- Kết nối Internet (để quét vsd.vn)
- Tài khoản Gmail dùng để gửi mail thông báo (khuyên dùng **App Password**, không dùng mật khẩu Gmail chính)

## 2. Cài đặt lần đầu

```powershell
git clone https://github.com/KaFi68/Flex.git
cd Flex
```

Tạo lại 2 file config **nhạy cảm** (không có trong Git — bị `.gitignore` chặn vì chứa
mật khẩu, phải tự tạo trên từng máy):

```powershell
cd vsd-flex-automation\config
copy email.config.json.example email.config.json
copy auth.config.json.example auth.config.json
```

Rồi mở 2 file vừa tạo, điền thông tin thật:

- **`email.config.json`**: `From` = Gmail dùng để gửi, `Password` = App Password 16 ký tự
  (lấy tại [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)),
  `To` = email nhận thông báo duyệt. Chi tiết đầy đủ: xem
  `vsd-flex-automation/config/README-email-setup.txt`.
- **`auth.config.json`**: tài khoản đăng nhập vào giao diện Flex mock (`Username`/`Password`).

**Không commit 2 file này lên Git, không gửi qua chat** — đưa trực tiếp hoặc qua kênh nội bộ an toàn.

## 3. Chạy thử xem demo (không cần config email/Flex thật)

```powershell
cd vsd-flex-automation\scripts
.\Seed-FlexMock.ps1              # tạo dữ liệu Flex giả lập ban đầu (chỉ chạy 1 lần, hoặc khi muốn làm mới sạch)
.\Start-FlexMockUI.ps1           # mở giao diện Flex mock tại http://localhost:8080
```

Đăng nhập bằng tài khoản trong `auth.config.json` (mặc định demo: `admin` / `1234`).

> ⚠️ `Seed-FlexMock.ps1` sẽ **xóa sạch** dữ liệu đang chờ duyệt hiện có — chỉ chạy khi
> muốn làm mới hoàn toàn, không chạy giữa lúc đang test dở.

## 4. Chạy luồng tự động thật (quét VSD → điền Flex)

```powershell
.\Run-VsdFlexAutomation.ps1
```

Script sẽ: quét VSD → so sánh với Flex (mock) → tự động điền tab TT chung cho mã mới/mã
cũ tồn đọng/mã chuyển sàn/đổi tên → gửi mail tổng hợp. Nhân viên **tự bấm Duyệt trên
giao diện** (`Start-FlexMockUI.ps1`) ở từng bước — script **không tự duyệt** để đảm bảo
luôn có người kiểm soát, đúng theo yêu cầu CR.

Tham số hữu ích:

| Tham số | Ý nghĩa |
|---|---|
| `-FetchMaxPages 30` | Giới hạn số trang quét VSD (dùng cho demo cho nhanh). **Sản xuất thật: bỏ tham số này** (mặc định `0` = quét toàn bộ) |
| `-SkipFetch` | Bỏ qua bước quét VSD, dùng lại snapshot gần nhất đã có |
| `-SkipEmail` | Không gửi mail thật, chỉ in ra console (test nhanh) |
| `-AutoApprove` | **CHỈ dùng để demo** — tự động duyệt hết 1 lượt. **Không dùng khi vận hành thật** |

## 5. Thiết lập chạy tự động theo lịch (Task Scheduler)

Chạy PowerShell với quyền Administrator:

```powershell
$action  = New-ScheduledTaskAction -Execute "powershell.exe" `
  -Argument '-ExecutionPolicy Bypass -NoProfile -File "C:\task1\vsd-flex-automation\scripts\Run-VsdFlexAutomation.ps1"'
$trigger = @(
  New-ScheduledTaskTrigger -Daily -At 8:00am
  New-ScheduledTaskTrigger -Daily -At 12:00pm
  New-ScheduledTaskTrigger -Daily -At 4:00pm
)
Register-ScheduledTask -TaskName "VSD-Flex-Automation" -Action $action -Trigger $trigger `
  -Description "Tu dong quet VSD va dien tab TT chung tren Flex, 3 lan/ngay"
```

> Sửa lại đường dẫn `C:\task1\...` cho đúng nơi bạn đặt project trên máy đích.
> Không thêm `-AutoApprove` — để nhân viên tự duyệt trên UI.

## 6. Cấu trúc project

```
vsd-flex-automation/
├── config/                      # file cấu hình (email, auth) — KHÔNG commit bản thật
├── data/                        # dữ liệu sinh ra lúc chạy (snapshot VSD, Flex mock store...)
├── logs/
└── scripts/
    ├── Modules/
    │   ├── FlexMockStore.psm1   # RULE MAPPING nghiệp vụ (VSD -> Flex) theo từng loại chứng khoán
    │   └── VsdDetail.psm1       # scrape trang chi tiết 1 mã trên vsd.vn
    ├── Fetch-VsdSecurities.ps1  # quét danh sách mã chứng khoán trên VSD
    ├── Process-NewSecurities.ps1    # mục 1 CR - mã mới
    ├── Process-MarketTransfer.ps1   # mục 2 CR - mã chuyển sàn
    ├── Process-OldSecurities.ps1    # mục 3 CR - mã cũ tồn đọng
    ├── Process-NameChanges.ps1      # mục 4 CR - đổi tên
    ├── Run-VsdFlexAutomation.ps1    # script tổng, chạy toàn bộ luồng
    ├── Start-FlexMockUI.ps1         # giao diện Flex giả lập (demo)
    └── Seed-FlexMock.ps1            # tạo dữ liệu Flex giả lập ban đầu
```

## 7. Lưu ý bảo mật

- File `config/email.config.json` và `config/auth.config.json` chứa mật khẩu thật —
  **không commit lên Git** (đã có trong `.gitignore`), **không gửi qua chat**.
- Nếu nghi ngờ App Password Gmail bị lộ, thu hồi và tạo lại ngay tại
  [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords).
