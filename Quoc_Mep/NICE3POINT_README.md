# ⚡ Nice3point.Revit.Api - Quick Start

## 🎯 Mục đích
Cho phép build Revit Add-in cho **nhiều phiên bản Revit** (2020-2026) chỉ bằng cách thay đổi 1 biến `RevitVersion`.

## 📦 Packages cần cài

```
Nice3point.Revit.Api.RevitAPI
Nice3point.Revit.Api.RevitAPIUI  
Nice3point.Revit.Api.AdWindows
```

## 🚀 Cách sử dụng nhanh

### Option 1: Visual Studio (Khuyến nghị)

1. **Right-click** project → **Manage NuGet Packages**
2. **Search** và cài 3 packages trên
3. **Remove** references cũ (RevitAPI, RevitAPIUI, AdWindows từ `C:\Program Files\Autodesk\Revit...`)
4. **Rebuild**

### Option 2: Command Line

```powershell
# Cài packages
.\install-nice3point-packages.ps1 -RevitVersion 2023

# Build cho Revit 2023
msbuild Quoc_MEP.csproj /p:RevitVersion=2023 /p:Configuration=Release

# Build cho tất cả versions (2020-2026)
.\build-all-versions.ps1
```

### Option 3: Chỉnh .csproj thủ công

Thêm vào `.csproj`:

```xml
<PropertyGroup>
  <RevitVersion>2023</RevitVersion>
</PropertyGroup>

<ItemGroup>
  <PackageReference Include="Nice3point.Revit.Api.RevitAPI" Version="$(RevitVersion).*" />
  <PackageReference Include="Nice3point.Revit.Api.RevitAPIUI" Version="$(RevitVersion).*" />
  <PackageReference Include="Nice3point.Revit.Api.AdWindows" Version="$(RevitVersion).*" />
</ItemGroup>
```

Xóa các `<Reference Include="RevitAPI">` cũ.

## 📁 Files đã tạo

| File | Mô tả |
|------|-------|
| `NICE3POINT_SETUP_GUIDE.md` | Hướng dẫn chi tiết đầy đủ |
| `install-nice3point-packages.ps1` | Script cài packages |
| `build-all-versions.ps1` | Script build tất cả versions |
| `packages.config` | ✅ Đã cập nhật với Nice3point packages |

## 🎨 Build cho phiên bản khác

```powershell
# Revit 2024
msbuild Quoc_MEP.csproj /p:RevitVersion=2024

# Revit 2025
msbuild Quoc_MEP.csproj /p:RevitVersion=2025

# Revit 2026
msbuild Quoc_MEP.csproj /p:RevitVersion=2026
```

## ✨ Lợi ích

| Trước | Sau |
|-------|-----|
| ❌ Hardcode path: `C:\Program Files\Autodesk\Revit 2023\...` | ✅ NuGet package tự động |
| ❌ Phải cài Revit để build | ✅ Build mà không cần Revit |
| ❌ Khó switch giữa các versions | ✅ Chỉ thay 1 biến |
| ❌ Mỗi version 1 project | ✅ 1 project cho tất cả |

## 📖 Xem thêm

- **Chi tiết đầy đủ**: `NICE3POINT_SETUP_GUIDE.md`
- **GitHub**: https://github.com/Nice3point/RevitApi
- **NuGet**: https://www.nuget.org/packages/Nice3point.Revit.Api.RevitAPI

## ⏭️ Next Steps

1. ✅ Đọc `NICE3POINT_SETUP_GUIDE.md` 
2. ⚡ Cài packages (Option 1, 2 hoặc 3)
3. 🧹 Xóa references cũ
4. 🔨 Build và test
5. 🎉 Enjoy!

---

**Updated**: November 2, 2025  
**Status**: ✅ Ready to use  
**Support**: Revit 2020-2026
