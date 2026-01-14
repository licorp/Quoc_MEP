# Quoc MEP Add-in - Universal Installation

Built on: 2026-01-14

## ⚠️ QUAN TRỌNG: Unblock DLL Files trước khi cài đặt!

Nếu bạn download file ZIP từ internet, Windows sẽ block các DLL file.
Chạy **UNBLOCK_FILES.bat** trước hoặc cài đặt sẽ tự động unblock.

Hoặc unblock thủ công:
- Chuột phải vào file ZIP → Properties → Check "Unblock" → OK
- Sau đó extract lại

## 🚀 Cài đặt đơn giản - Chỉ 2 bước!

### Bước 1: Chạy INSTALL.bat
- Double-click file **INSTALL.bat**
- Chọn phiên bản Revit của bạn (1-7) hoặc chọn 8 để cài tất cả
- Script sẽ tự động unblock và copy files vào đúng thư mục

### Bước 2: Restart Revit
Khởi động lại Revit và tận hưởng!

## 📁 Cấu trúc sau khi cài (ví dụ Revit 2024):

%APPDATA%\Autodesk\Revit\Addins\2024\
├── Quoc_MEP_Universal.addin    ← File addin chính
├── Quoc_MEP_Loader.dll         ← Universal loader (165KB)
└── Revit2024\                  ← DLL cho version này
    ├── Quoc_MEP.dll
    └── (16 dependencies DLLs)

## 🎯 Cách hoạt động:

1. File .addin load Quoc_MEP_Loader.dll
2. Loader tự động detect phiên bản Revit đang chạy
3. Load đúng DLL từ thư mục Revit{Version}
4. Add-in chạy bình thường!

## ✨ Tính năng:

- ✅ Tự động detect phiên bản Revit
- ✅ Chỉ copy version bạn cần (tiết kiệm dung lượng)
- ✅ Hoặc cài tất cả version cùng lúc (option 8)
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
2. File Quoc_MEP_Loader.dll tồn tại
3. Thư mục Revit{Version} chứa Quoc_MEP.dll
4. Khởi động lại Revit sau khi cài đặt
