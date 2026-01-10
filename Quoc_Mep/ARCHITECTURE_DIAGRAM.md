# Sơ đồ Kiến trúc - Revit API MEP

## 📊 Cấu trúc Solutions

```
┌─────────────────────────────────────────────────────────────────┐
│                    WORKSPACE: RevitAPIMEP                        │
└─────────────────────────────────────────────────────────────────┘
                              │
                ┌─────────────┴─────────────┐
                │                           │
         ┌──────▼──────┐           ┌───────▼────────┐
         │ SOLUTION 1   │           │  SOLUTION 2     │
         │ Quoc_MEP     │           │ Schedule        │
         │ _Main.sln    │           │ Manager.sln     │
         └──────┬──────┘           └───────┬────────┘
                │                           │
         ┌──────▼──────┐           ┌───────▼────────┐
         │ PROJECT      │           │  PROJECT        │
         │ Quoc_MEP     │           │ Schedule        │
         │ .csproj      │           │ Manager.csproj  │
         │              │           │                 │
         │ Old-style    │           │ SDK-style       │
         │ .NET 4.8     │           │ .NET 4.8        │
         └──────┬──────┘           └───────┬────────┘
                │                           │
                │    ┌─────────────┐       │
                └────► RIBBON HOST ◄───────┘
                     │  (PROJECT)  │
                     │ RibbonHost  │
                     │   .csproj   │
                     └──────┬──────┘
                            │
                     ┌──────▼──────┐
                     │   OUTPUT     │
                     │ bin\Release\ │
                     │ Revit2023\   │
                     └──────┬──────┘
                            │
              ┌─────────────┼─────────────┐
              │             │             │
    ┌─────────▼──┐   ┌─────▼──────┐  ┌──▼─────────┐
    │ RibbonHost │   │ Schedule   │  │ Quoc_MEP   │
    │   .dll     │   │ Manager    │  │   .dll     │
    │            │   │   .dll     │  │ (optional) │
    └─────────┬──┘   └─────┬──────┘  └──┬─────────┘
              │            │            │
              └────────────┼────────────┘
                           │
                    ┌──────▼──────┐
                    │   REVIT     │
                    │ Tab: Quoc   │
                    │     MEP     │
                    │             │
                    │ Panel: MEP  │
                    │    Tools    │
                    └─────────────┘
```

## 🔄 Luồng Build

```
DEVELOPER
    │
    ├─── Làm việc với Schedule Manager
    │    │
    │    └─► Build ScheduleManager.sln
    │        │
    │        ├─► build-schedule-manager.ps1
    │        │
    │        └─► Output: ScheduleManager.dll
    │
    ├─── Làm việc với Main Features
    │    │
    │    └─► Build Quoc_MEP_Main.sln
    │        │
    │        └─► Output: Quoc_MEP.dll
    │
    └─── Update Ribbon Host
         │
         └─► Build RibbonHost project
             │
             ├─► build-ribbonhost.ps1
             │
             └─► Output: Quoc_MEP.RibbonHost.dll
```

## 🎯 Luồng Runtime trong Revit

```
REVIT STARTS
    │
    └─► Đọc manifest: Quoc_MEP_SharedRibbon.addin
        │
        └─► Load: Quoc_MEP.RibbonHost.dll
            │
            ├─► Application.OnStartup()
            │   │
            │   ├─► CreateRibbonTab("Quoc MEP")
            │   │
            │   ├─► CreateRibbonPanel("MEP Tools")
            │   │
            │   └─► LoadCommands()
            │       │
            │       ├─► Scan folder: bin\Release\Revit2023\
            │       │
            │       ├─► Find: ScheduleManager.dll
            │       │   │
            │       │   └─► Load: ScheduleManagerCommandInfo
            │       │       │
            │       │       └─► Create button: "Schedule Manager"
            │       │
            │       ├─► Find: Quoc_MEP.dll (nếu có)
            │       │   │
            │       │   └─► Load: ExportCommandInfo, etc.
            │       │       │
            │       │       └─► Create buttons: "Export", "Connect"...
            │       │
            │       └─► Panel hiển thị tất cả buttons
            │
            └─► USER CLICK BUTTON
                │
                └─► Execute command từ DLL tương ứng
```

## 🏗️ Cấu trúc Folders

```
RevitAPIMEP/
│
├─── 📁 Solutions
│    ├─── Quoc_MEP_Main.sln          ← CHÍNH
│    └─── ScheduleManager.sln        ← RIÊNG
│
├─── 📁 Projects
│    ├─── Quoc_MEP.csproj            ← Old-style
│    ├─── Quoc_MEP.RibbonHost/       ← Infrastructure
│    │    └─── Quoc_MEP.RibbonHost.csproj
│    └─── ScheduleManager/           ← SDK-style
│         └─── ScheduleManager.csproj
│
├─── 📁 Source Code (Solution chính)
│    ├─── App/                        ← Application startup
│    ├─── Export/                     ← Export features
│    ├─── Connect/                    ← Connect tools
│    ├─── DrawPipe/                   ← Draw pipe
│    ├─── Place Support/              ← Support placement
│    ├─── Lib/                        ← Shared utilities
│    └─── ... (nhiều folders khác)
│
├─── 📁 Source Code (Schedule Manager)
│    └─── ScheduleManager/
│         ├─── ScheduleManagerCommand.cs
│         ├─── ScheduleManagerViewModel.cs
│         ├─── ScheduleManagerWindow.xaml
│         ├─── AsyncScheduleReader.cs
│         └─── ... (10 files)
│
├─── 📁 Build Scripts
│    ├─── build-ribbonhost.ps1
│    └─── build-schedule-manager.ps1
│
├─── 📁 Output
│    └─── bin/Release/
│         ├─── Revit2020/
│         ├─── Revit2021/
│         ├─── Revit2022/
│         ├─── Revit2023/            ← CURRENT
│         │    ├─── Quoc_MEP.RibbonHost.dll
│         │    ├─── ScheduleManager.dll
│         │    └─── Wpf.Ui.dll
│         └─── Revit2024/
│
└─── 📁 Deployment
     └─── Quoc_MEP_SharedRibbon.addin → Copy to Revit addins
```

## 💡 Tại sao có 2 Solutions?

### ❌ TRƯỚC (1 Solution lớn):
```
Quoc_MEP.sln
    │
    └─── Quoc_MEP.csproj
         ├─── Export/         (100+ files)
         ├─── Connect/        (50+ files)
         ├─── DrawPipe/       (30+ files)
         ├─── ScheduleManager/ (10 files)
         └─── ... (hàng ngàn files)

Build time: 15-30 giây
Sửa 1 file Schedule → Phải rebuild TẤT CẢ
Package issues với old-style project
```

### ✅ SAU (2 Solutions riêng):
```
Quoc_MEP_Main.sln              ScheduleManager.sln
    │                               │
    └─── Quoc_MEP.csproj           └─── ScheduleManager.csproj
         ├─── Export/                    ├─── ScheduleManagerCommand.cs
         ├─── Connect/                   ├─── ScheduleManagerViewModel.cs
         ├─── DrawPipe/                  └─── ... (chỉ 10 files)
         └─── ...
                                   Build time: 0.6 giây ⚡
Build time: 10-20 giây           Sửa Schedule → Build RIÊNG
Chỉ build khi cần                SDK-style → Nice3point OK ✅
```

## 🎯 Ví dụ Workflow thực tế

### Scenario 1: Sửa bug Schedule Manager
```
1. Mở ScheduleManager.sln (không cần mở solution chính)
2. Sửa file ScheduleManagerViewModel.cs
3. Build: .\build-schedule-manager.ps1 -Versions @("2023")
4. Test trong Revit
   ✅ Nhanh: 0.6s build time
   ✅ Không ảnh hưởng code khác
```

### Scenario 2: Thêm feature Export mới
```
1. Mở Quoc_MEP_Main.sln
2. Thêm code trong Export/
3. Build solution trong Visual Studio
4. Test trong Revit
   ✅ Schedule Manager không bị rebuild
```

### Scenario 3: Deploy cả 2
```
1. Build RibbonHost: .\build-ribbonhost.ps1
2. Build ScheduleManager: .\build-schedule-manager.ps1
3. Build Quoc_MEP_Main.sln (Visual Studio)
4. Copy manifest file
5. Khởi động Revit
   ✅ Tất cả features đều có trên 1 panel
```

## 📈 So sánh Performance

| Thao tác | Trước (1 Solution) | Sau (2 Solutions) |
|----------|-------------------|-------------------|
| Build Schedule Manager | 15-30s | 0.6s ⚡ |
| Build toàn bộ | 15-30s | 10-20s (chỉ khi cần) |
| Sửa 1 file Schedule | Rebuild TẤT CẢ | Rebuild Schedule only |
| Nice3point packages | ❌ Không hoạt động | ✅ Hoạt động tốt |
| Test độc lập | ❌ Khó | ✅ Dễ dàng |

---

**Kết luận**: 
- **Quoc_MEP_Main.sln** = Solution CHÍNH cho tất cả features
- **ScheduleManager.sln** = Solution RIÊNG cho Schedule Manager
- **RibbonHost** = Gom cả 2 lại thành 1 ribbon panel

✅ Build nhanh hơn  
✅ Dễ maintain  
✅ Package management tốt hơn  
✅ Test độc lập được  
