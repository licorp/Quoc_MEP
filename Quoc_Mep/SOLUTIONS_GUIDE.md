# Cấu trúc Solutions - Dự án Revit API MEP

## 📂 CÁC SOLUTION CHÍNH

### 1️⃣ **Quoc_MEP_Main.sln** (SOLUTION CHÍNH)
**Mục đích**: Chứa tất cả các tính năng của add-in Quoc MEP

**Project**: `Quoc_MEP.csproj` (old-style .NET Framework 4.8)

**Chứa các tính năng**:
- Export (DWG, DXF, PDF, NWC)
- Connect (nối các đối tượng MEP)
- DrawPipe (vẽ ống)
- Place Support (đặt giá đỡ)
- Rotate, Split Duct, Trans Data Para
- Sheet from Excel
- Selection Filter
- UpDownTool
- ... và tất cả features khác

**Build**: Sử dụng file `Quoc_MEP.csproj` (không cần build script riêng)

---

### 2️⃣ **ScheduleManager.sln** (SOLUTION RIÊNG - MỚI)
**Mục đích**: Schedule Manager - tính năng chỉnh sửa schedule an toàn

**Project**: `ScheduleManager/ScheduleManager.csproj` (SDK-style .NET 4.8)

**Chứa các tính năng**:
- Đọc schedule data bất đồng bộ (không crash Revit)
- Chỉnh sửa schedule trong DataGrid
- Highlight elements trong view
- Export to Excel
- MVVM pattern với async/await

**Build**: 
```powershell
.\build-schedule-manager.ps1 -Versions @("2023")
```

**Ưu điểm**:
- ✅ Build độc lập (không cần rebuild toàn bộ Quoc_MEP)
- ✅ Dùng SDK-style project (Nice3point packages hoạt động tốt)
- ✅ Kiến trúc an toàn với RevitAsyncHelper
- ✅ Test riêng được

---

## 🎯 RIBBON HOST (Infrastructure chung)

### **Quoc_MEP.RibbonHost/** (PROJECT - không phải solution)
**Mục đích**: Tạo ribbon panel chung cho tất cả các add-ins

**Cách hoạt động**:
1. Revit load `Quoc_MEP.RibbonHost.dll` (từ manifest file)
2. RibbonHost tạo tab "Quoc MEP" và panel "MEP Tools"
3. RibbonHost tự động tìm tất cả DLL trong cùng folder
4. RibbonHost đọc `CommandInfo` classes và tạo buttons
5. Tất cả commands từ cả 2 solutions xuất hiện trên 1 panel

**Build**:
```powershell
.\build-ribbonhost.ps1 -Versions @("2023")
```

**Manifest file**: `Quoc_MEP_SharedRibbon.addin`

---

## 📊 SO SÁNH

| Tiêu chí | Quoc_MEP_Main.sln | ScheduleManager.sln |
|----------|-------------------|---------------------|
| **Loại project** | Old-style (.NET Framework) | SDK-style (modern) |
| **Nice3point packages** | ❌ Không hoạt động tốt | ✅ Hoạt động hoàn hảo |
| **Build time** | Lâu (nhiều features) | Nhanh (~0.6s) |
| **Build độc lập** | N/A (solution chính) | ✅ Có |
| **Dùng khi nào** | Phát triển features chính | Phát triển Schedule Manager |

---

## 🚀 CÁCH BUILD

### Build tất cả cho Revit 2023:
```powershell
# 1. Build RibbonHost (bắt buộc)
.\build-ribbonhost.ps1 -Versions @("2023")

# 2. Build ScheduleManager (tính năng riêng)
.\build-schedule-manager.ps1 -Versions @("2023")

# 3. Build Quoc_MEP (solution chính) - nếu cần
# Mở Visual Studio → Build Quoc_MEP_Main.sln
# HOẶC dùng MSBuild
```

### Build cho nhiều versions:
```powershell
# Build RibbonHost cho tất cả versions
.\build-ribbonhost.ps1 -Versions @("2020", "2021", "2022", "2023", "2024")

# Build ScheduleManager cho tất cả versions  
.\build-schedule-manager.ps1 -Versions @("2020", "2021", "2022", "2023", "2024")
```

---

## 📁 OUTPUT

Tất cả DLLs build vào cùng 1 folder:
```
bin\Release\Revit2023\
├── Quoc_MEP.RibbonHost.dll      ← Ribbon host (bắt buộc)
├── ScheduleManager.dll           ← Schedule Manager feature
├── Quoc_MEP.dll                  ← Main add-in (nếu build)
└── Wpf.Ui.dll                    ← Dependencies
```

---

## 🔧 DEPLOY TO REVIT

### Bước 1: Copy manifest file
```powershell
Copy-Item "Quoc_MEP_SharedRibbon.addin" `
    -Destination "$env:APPDATA\Autodesk\Revit\Addins\2023\"
```

### Bước 2: DLLs đã ở đúng chỗ
Tất cả DLLs đã build vào `bin\Release\Revit2023\`

### Bước 3: Khởi động Revit
- Mở Revit 2023
- Tìm tab "Quoc MEP"
- Thấy panel "MEP Tools" với các buttons

---

## 🎯 KHI NÀO DÙNG CÁI NÀO?

### Dùng **Quoc_MEP_Main.sln** khi:
- ✅ Thêm/sửa features chính (Export, Connect, DrawPipe...)
- ✅ Thay đổi Ribbon.cs
- ✅ Thêm resources, images
- ✅ Update App.cs (IExternalApplication)

### Dùng **ScheduleManager.sln** khi:
- ✅ Làm việc với Schedule Manager
- ✅ Sửa bug Schedule Manager
- ✅ Thêm tính năng cho Schedule Manager
- ✅ Test Schedule Manager riêng

### Build **RibbonHost** khi:
- ✅ Thay đổi cách load commands
- ✅ Thay đổi cách tạo ribbon
- ✅ Thêm logic discovery mới

---

## ⚠️ LƯU Ý QUAN TRỌNG

1. **RibbonHost phải build trước**
   - Nếu không có RibbonHost.dll, Revit sẽ không load được add-in

2. **Không xóa folder "Schedule Manager"** (có dấu cách)
   - Đây là source code gốc của Schedule Manager
   - Folder "ScheduleManager" (không dấu cách) là project mới

3. **Output folder chung**
   - Tất cả projects đều build vào `bin\Release\Revit{Version}\`
   - RibbonHost sẽ tự động tìm và load tất cả DLLs

4. **Manifest file**
   - Chỉ load `Quoc_MEP.RibbonHost.dll`
   - Các DLL khác được RibbonHost discover tự động

---

## 📚 TÀI LIỆU THAM KHẢO

- `SHARED_RIBBON_ARCHITECTURE.md` - Kiến trúc chi tiết
- `QUICK_START.md` - Hướng dẫn nhanh
- `NICE3POINT_README.md` - Về Nice3point packages

---

## ❓ FAQ

**Q: Tại sao có 2 solutions?**  
A: Để Schedule Manager build nhanh và độc lập, không cần rebuild toàn bộ project

**Q: Có thể xóa solution cũ không?**  
A: Quoc_MEP_Main.sln là solution chính, KHÔNG xóa. Các solution trùng lặp đã được xóa.

**Q: Build solution nào trước?**  
A: Build RibbonHost trước, sau đó build ScheduleManager hoặc Quoc_MEP theo nhu cầu

**Q: Có thể add thêm solution mới?**  
A: Có! Tạo solution mới, build vào `bin\Release\Revit{Version}\`, RibbonHost sẽ tự động load

---

**Cập nhật**: 02/11/2025  
**Người tạo**: AI Assistant  
**Trạng thái**: ✅ Đã test build thành công cho Revit 2023
