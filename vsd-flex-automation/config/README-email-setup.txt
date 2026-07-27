CACH THIET LAP GUI MAIL TEST
============================

1. Copy file "email.config.json.example" thanh "email.config.json" (cung thu muc nay).

2. Mo file "email.config.json" va dien:
   - SmtpServer / Port: neu dung Gmail lam tai khoan GUI thi de nguyen
     smtp.gmail.com / 587. Neu cong ty co SMTP relay rieng, dien thong tin do vao.
   - From: dia chi email dung de GUI di (co the la Gmail cua ban, hoac mail cong ty).
   - Password: KHONG phai mat khau dang nhap Gmail thuong.
       -> Neu dung Gmail: phai bat "2-Step Verification" cho tai khoan Google,
          sau do vao https://myaccount.google.com/apppasswords de tao "App Password"
          (chuoi 16 ky tu), dan chuoi do vao day.
       -> Neu dung mail cong ty: hoi IT ve SMTP relay noi bo (thuong khong can
          password, chi can nam trong mang noi bo / whitelist IP).
   - To: dia chi nhan (mac dinh da dien san nguyencaophi8969@gmail.com de test).

3. LUU Y BAO MAT: file "email.config.json" (sau khi dien that) chua mat khau/app
   password that. KHONG gui file nay qua chat, KHONG dua len git/public repo o dang
   plain text. File nay chi nam tren may local de script doc.

4. Chay thu:
   powershell -ExecutionPolicy Bypass -File "..\scripts\Send-VsdDiffReport.ps1"

   Script se tu dong lay bao cao diff moi nhat trong data/reports va gui mail.
