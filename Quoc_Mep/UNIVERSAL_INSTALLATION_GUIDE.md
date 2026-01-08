# Quoc MEP Universal Installation Guide

## 🎯 Giải pháp: 1 Package cho TẤT CẢ phiên bản Revit!

Thay vì phải build và cài riêng cho từng phiên bản Revit, bây giờ bạn chỉ cần:
- ✅ 1 lần download
- ✅ 1 lần cài đặt
- ✅ Chạy cho TẤT CẢ Revit 2020-2026

## 📦 Cách hoạt động

```
Quoc_MEP_Universal.addin (file chính)
    ↓
Quoc_MEP_Loader.dll (auto-detect Revit version)
    ↓
Revit2020/Quoc_MEP.dll  ← Load nếu Revit 2020
Revit2021/Quoc_MEP.dll  ← Load nếu Revit 2021
Revit2022/Quoc_MEP.dll  ← Load nếu Revit 2022
...và tất cả versions khác
```

## 🚀 Cài đặt (3 bước đơn giản)

### Bước 1: Download package
Từ GitHub Actions Artifacts:
- Vào tab **Actions** trên GitHub
- Chọn workflow run mới nhất
- Download **Quoc_MEP_Universal_Package.zip**

### Bước 2: Giải nén và copy
```
1. Giải nén file zip
2. Copy TOÀN BỘ thư mục vào:
   %APPDATA%\Autodesk\Revit\Addins\{Version}\

   Ví dụ cho Revit 2024:
   %APPDATA%\Autodesk\Revit\Addins\2024\
```

### Bước 3: Restart Revit
Khởi động lại Revit và enjoy!

## 📁 Cấu trúc sau khi cài

```
%APPDATA%\Autodesk\Revit\Addins\2024\
├── Quoc_MEP_Universal.addin    ← File addin chính
├── Quoc_MEP_Loader.dll         ← Loader tự động
├── Revit2020\                  ← DLL cho Revit 2020
│   ├── Quoc_MEP.dll
│   └── Resources\
├── Revit2021\                  ← DLL cho Revit 2021
├── Revit2022\                  ← DLL cho Revit 2022
├── Revit2023\                  ← DLL cho Revit 2023
├── Revit2024\                  ← DLL cho Revit 2024
├── Revit2025\                  ← DLL cho Revit 2025
└── Revit2026\                  ← DLL cho Revit 2026
```

## 🔧 Hoặc dùng Auto Install Script

Chạy file **INSTALL.bat**:
1. Double-click INSTALL.bat
2. Chọn phiên bản Revit của bạn
3. Script tự động copy files vào đúng vị trí
4. Done!

## ✨ Ưu điểm

### So với cách cũ (build riêng từng version):
- ❌ Phải build 7 lần (mỗi version 1 lần)
- ❌ Phải cài 7 lần (mỗi version 1 lần)
- ❌ Update phải làm lại 7 lần

### Với Universal Package (cách mới):
- ✅ Build 1 lần cho tất cả
- ✅ Cài 1 lần cho tất cả
- ✅ Update chỉ cần replace files

## 🆘 Troubleshooting

### Lỗi: "Không tìm thấy DLL cho Revit XXXX"
**Nguyên nhân:** Thiếu thư mục RevitXXXX hoặc file Quoc_MEP.dll

**Giải pháp:**
1. Kiểm tra thư mục RevitXXXX tồn tại
2. Kiểm tra file Quoc_MEP.dll trong thư mục đó
3. Download lại package và cài lại

### Lỗi: "Không thể tạo instance của Ribbon class"
**Nguyên nhân:** File DLL bị corrupt hoặc không tương thích

**Giải pháp:**
1. Download lại package
2. Unblock file zip trước khi giải nén (Right-click → Properties → Unblock)
3. Cài lại

### Add-in không xuất hiện trên Ribbon
**Nguyên nhân:** File .addin không ở đúng vị trí

**Giải pháp:**
1. Kiểm tra file Quoc_MEP_Universal.addin ở:
   `%APPDATA%\Autodesk\Revit\Addins\{Version}\`
2. Restart Revit
3. Kiểm tra Revit add-in manager (R → Options → Add-ins)

## 📞 Support

Nếu gặp vấn đề:
1. Kiểm tra file log tại: `%TEMP%\QuocMEP_*.log`
2. Mở Revit Journal file: `%APPDATA%\Autodesk\Revit\{Version}\Journals\`
3. Tạo issue trên GitHub với thông tin lỗi

## 🎯 Features mới (Move Connect Align)

- Button **Move Connect Align** trong panel **Modify**
- Chức năng: Di chuyển và kết nối MEP families
- Cách dùng:
  1. Click button
  2. Chọn MEP family đích (destination)
  3. Chọn MEP family nguồn (sẽ di chuyển)
  4. Tự động align và connect!
