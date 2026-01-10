# Revit API MEP - Multi-Solution Architecture

Dự án Revit Add-in với kiến trúc multi-solution, hỗ trợ Revit 2020-2024.

## 🎯 Tổng quan

Dự án này sử dụng kiến trúc **2 solutions độc lập + 1 RibbonHost chung**:

1. **Quoc_MEP_Main.sln** - Solution chính (Export, Connect, DrawPipe, etc.)
2. **ScheduleManager.sln** - Solution riêng cho Schedule Manager
3. **RibbonHost** - Infrastructure gom commands từ cả 2 solutions

## 📂 Cấu trúc Solutions

```
📁 RevitAPIMEP/
│
├── 🔵 Quoc_MEP_Main.sln          ← SOLUTION CHÍNH
│   └── Quoc_MEP.csproj           (Old-style, .NET Framework 4.8)
│       ├── Export/               - Export DWG, DXF, PDF, NWC
│       ├── Connect/              - Nối đối tượng MEP
│       ├── DrawPipe/             - Vẽ ống
│       ├── Place Support/        - Đặt giá đỡ
│       └── ... (nhiều features khác)
│
├── 🟢 ScheduleManager.sln        ← SOLUTION RIÊNG
│   └── ScheduleManager/
│       └── ScheduleManager.csproj (SDK-style, .NET Framework 4.8)
│           ├── ScheduleManagerCommand.cs
│           ├── ScheduleManagerViewModel.cs
│           ├── ScheduleManagerWindow.xaml
│           └── ... (10 files)
│
└── 🟡 Quoc_MEP.RibbonHost/       ← RIBBON HOST (Project)
    └── Quoc_MEP.RibbonHost.csproj (SDK-style)
        └── Application.cs         - Tạo ribbon chung + discovery
```

## ⚡ Quick Start

### Build cho Revit 2023:
```powershell
# 1. Build RibbonHost (bắt buộc)
.\build-ribbonhost.ps1 -Versions @("2023")

# 2. Build ScheduleManager
.\build-schedule-manager.ps1 -Versions @("2023")

# 3. Deploy
Copy-Item "Quoc_MEP_SharedRibbon.addin" -Destination "$env:APPDATA\Autodesk\Revit\Addins\2023\"
```

### Build cho tất cả versions:
```powershell
.\build-ribbonhost.ps1 -Versions @("2020","2021","2022","2023","2024")
.\build-schedule-manager.ps1 -Versions @("2020","2021","2022","2023","2024")
```

## 📖 Tài liệu

| File | Mô tả |
|------|-------|
| **[SOLUTIONS_GUIDE.md](SOLUTIONS_GUIDE.md)** | 📘 Hướng dẫn chi tiết về các solutions |
| **[ARCHITECTURE_DIAGRAM.md](ARCHITECTURE_DIAGRAM.md)** | 📊 Sơ đồ kiến trúc với diagrams |
| **[SHARED_RIBBON_ARCHITECTURE.md](SHARED_RIBBON_ARCHITECTURE.md)** | 🎯 Kiến trúc Ribbon Host pattern |
| **[QUICK_START.md](QUICK_START.md)** | ⚡ Hướng dẫn nhanh build & deploy |

## 🔧 Khi nào dùng gì?

### Dùng **Quoc_MEP_Main.sln** khi:
- ✅ Thêm/sửa features chính (Export, Connect, DrawPipe...)
- ✅ Thay đổi Ribbon.cs
- ✅ Update App.cs hoặc resources

### Dùng **ScheduleManager.sln** khi:
- ✅ Làm việc với Schedule Manager
- ✅ Test Schedule Manager riêng
- ✅ Sửa bug Schedule Manager

### Build **RibbonHost** khi:
- ✅ Thay đổi cách load commands
- ✅ Thay đổi logic discovery
- ✅ Update ribbon UI

## 💡 Tại sao có 2 Solutions?

### Trước (1 Solution):
```
❌ Build time: 15-30 giây
❌ Sửa 1 file → Rebuild TẤT CẢ
❌ Nice3point packages không hoạt động tốt
```

### Sau (2 Solutions):
```
✅ Schedule Manager build: 0.6 giây
✅ Build độc lập, không ảnh hưởng nhau
✅ SDK-style project → Nice3point packages OK
✅ Dễ test và maintain
```

## 🏗️ Output Structure

```
bin\Release\Revit2023\
├── Quoc_MEP.RibbonHost.dll     ← Revit load file này
├── ScheduleManager.dll          ← RibbonHost tự động discover
├── Quoc_MEP.dll                 ← (Optional - nếu build solution chính)
└── Wpf.Ui.dll                   ← Dependencies
```

## 🎯 Workflow Example

### Scenario: Sửa bug Schedule Manager
```powershell
# 1. Mở solution riêng
code ScheduleManager.sln

# 2. Sửa code
# Edit: ScheduleManager/ScheduleManagerViewModel.cs

# 3. Build nhanh
.\build-schedule-manager.ps1 -Versions @("2023")
# ⚡ Chỉ 0.6 giây!

# 4. Test trong Revit
# Khởi động Revit → Tab "Quoc MEP" → Click "Schedule Manager"
```

## 📊 Build Status

| Version | RibbonHost | ScheduleManager | Status |
|---------|-----------|----------------|--------|
| 2023 | ✅ 1.3s | ✅ 0.6s | Tested |
| 2020-2024 | 🔶 Ready | 🔶 Ready | Not tested |

## 🔍 Troubleshooting

### Commands không hiện trong ribbon?
```powershell
# Kiểm tra DLL có trong folder không
ls "bin\Release\Revit2023\*.dll"

# Kiểm tra CommandInfo class có đúng không
# Phải có: Name, Text, Tooltip, CommandClass (static properties)
```

### Build errors?
```powershell
# Kiểm tra .NET Framework 4.8 SDK
dotnet --version

# Clean và rebuild
Remove-Item "bin\Release" -Recurse -Force
.\build-ribbonhost.ps1 -Versions @("2023")
```

## 📝 Files đã dọn dẹp

Các files sau đã được xóa (trùng lặp/không dùng):
- ❌ RevitAPIMEP.sln (trùng với Quoc_MEP_Main.sln)
- ❌ ScheduleManager_2020.csproj (dùng ScheduleManager.sln)
- ❌ ScheduleManager_2023.csproj (dùng ScheduleManager.sln)
- ❌ Quoc_MEP_2020.csproj (dùng Quoc_MEP.csproj với /p:RevitVersion)
- ❌ Quoc_MEP_2023.csproj (dùng Quoc_MEP.csproj với /p:RevitVersion)

## 🚀 Next Steps

1. **Test trong Revit 2023**
   ```powershell
   # Copy manifest
   Copy-Item "Quoc_MEP_SharedRibbon.addin" -Destination "$env:APPDATA\Autodesk\Revit\Addins\2023\"
   
   # Launch Revit → Check ribbon
   ```

2. **Build cho versions khác**
   ```powershell
   .\build-ribbonhost.ps1 -Versions @("2020","2021","2022","2024")
   .\build-schedule-manager.ps1 -Versions @("2020","2021","2022","2024")
   ```

3. **Thêm features mới**
   - Tạo solution mới (e.g., MyFeature.sln)
   - Build to bin\Release\Revit{Version}\
   - RibbonHost tự động discover!

## 🤝 Contributing

Khi thêm command mới:
1. Tạo CommandInfo class với static properties
2. Implement IExternalCommand
3. Build to shared bin folder
4. Done! Button tự động xuất hiện

## 📄 License

[Your License Here]

---

**Status**: ✅ Build successfully for Revit 2023  
**Last Updated**: 02/11/2025  
**Architecture**: Multi-solution with Shared Ribbon Host  
