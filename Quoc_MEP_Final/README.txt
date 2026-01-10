# Quoc MEP Add-in - Universal Installation

Built on: 2026-01-08 09:33:05

## 🚀 Cài đặt đơn giản - Chỉ 3 bước!

### Bước 1: Copy toàn bộ thư mục
Copy toàn bộ thư mục này vào:

    %APPDATA%\Autodesk\Revit\Addins\

### Bước 2: Đổi tên thư mục
Đổi tên thư mục thành tên phiên bản Revit của bạn:
- Revit 2020 → 2020
- Revit 2021 → 2021
- Revit 2022 → 2022
- Revit 2023 → 2023
- Revit 2024 → 2024
- Revit 2025 → 2025
- Revit 2026 → 2026

### Bước 3: Restart Revit
Khởi động lại Revit và tận hưởng!

## 📁 Cấu trúc thư mục sau khi cài:

%APPDATA%\Autodesk\Revit\Addins\2024\
├── Quoc_MEP_Universal.addin    ← File addin chính
├── Quoc_MEP_Loader.dll         ← Loader tự động
└── Revit2024\                  ← DLL cho version này
    ├── Quoc_MEP.dll
    ├── Resources\
    └── (các dependencies khác)

## 🎯 Cách hoạt động:

1. File .addin load Quoc_MEP_Loader.dll
2. Loader tự động detect phiên bản Revit
3. Load đúng DLL từ thư mục Revit{Version}
4. Add-in chạy bình thường!

## ✨ Tính năng:

- ✅ Tự động detect phiên bản Revit
- ✅ 1 lần cài đặt cho tất cả versions
- ✅ Dễ dàng update (chỉ replace files)
- ✅ Hỗ trợ Revit 2020-2026

## 📦 Included Versions:
- Revit 2020 (.NET 4.8)
- Revit 2021 (.NET 4.8)
- Revit 2022 (.NET 4.8)
- Revit 2023 (.NET 4.8)
- Revit 2024 (.NET 4.8)
- Revit 2025 (.NET 8.0)
- Revit 2026 (.NET 8.0)

## 🆘 Troubleshooting:

Nếu gặp lỗi, kiểm tra:
1. File Quoc_MEP_Universal.addin ở đúng vị trí
2. Thư mục Revit{Version} tồn tại
3. File Quoc_MEP.dll trong thư mục version

