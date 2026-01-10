# 🔧 Hướng dẫn cập nhật Nice3point.Revit.Api - Hỗ trợ đa phiên bản Revit

## 📌 Tại sao dùng Nice3point.Revit.Api?

✅ **Hỗ trợ đa phiên bản Revit** (2020-2026) - Chỉ cần thay đổi `RevitVersion`  
✅ **NuGet Package** - Không cần cài Revit để build  
✅ **Tự động cập nhật** - Dùng wildcard version  
✅ **Clean references** - Không hardcode đường dẫn  

## 🚀 Cách cập nhật (Khuyến nghị - Dùng Visual Studio)

### Bước 1: Mở NuGet Package Manager

1. **Mở Visual Studio**
2. **Right-click** vào project `Quoc_MEP` trong Solution Explorer
3. Chọn **"Manage NuGet Packages..."**

### Bước 2: Gỡ references cũ (Nếu cần)

Trong Solution Explorer, mở **References** > Xóa:
- ❌ `RevitAPI` (reference trực tiếp)
- ❌ `RevitAPIUI` (reference trực tiếp)
- ❌ `AdWindows` (reference trực tiếp)

### Bước 3: Cài đặt Nice3point.Revit.Api packages

Trong NuGet Package Manager:

#### 📦 Package 1: RevitAPI
```
Search: Nice3point.Revit.Api.RevitAPI
Version: Chọn theo RevitVersion (VD: 2023.x.x cho Revit 2023)
Click: Install
```

#### 📦 Package 2: RevitAPIUI
```
Search: Nice3point.Revit.Api.RevitAPIUI
Version: Chọn theo RevitVersion (VD: 2023.x.x)
Click: Install
```

#### 📦 Package 3: AdWindows
```
Search: Nice3point.Revit.Api.AdWindows
Version: Chọn theo RevitVersion (VD: 2023.x.x)
Click: Install
```

### Bước 4: Chọn phiên bản packages

| Revit Version | Package Version Pattern |
|--------------|------------------------|
| Revit 2020   | `2020.*.*`            |
| Revit 2021   | `2021.*.*`            |
| Revit 2022   | `2022.*.*`            |
| Revit 2023   | `2023.*.*`            |
| Revit 2024   | `2024.*.*`            |
| Revit 2025   | `2025.*.*`            |
| Revit 2026   | `2026.*.*`            |

---

## 🔧 Cách cập nhật (Command Line với NuGet.exe)

### Tại thư mục project:

```powershell
cd "d:\RevitAPI_tu viet\RevitAPIMEP"

# Cài cho Revit 2023
.\nuget.exe install packages.config -OutputDirectory packages

# Hoặc cài từng package riêng
.\nuget.exe install Nice3point.Revit.Api.RevitAPI -OutputDirectory packages
.\nuget.exe install Nice3point.Revit.Api.RevitAPIUI -OutputDirectory packages
.\nuget.exe install Nice3point.Revit.Api.AdWindows -OutputDirectory packages
```

---

## ⚡ Cách cập nhật (Tự động - Chỉnh sửa .csproj)

### Cách tốt nhất: Dùng PackageReference (Modern)

Thêm vào file `.csproj` (trong `<ItemGroup>`):

```xml
<PropertyGroup>
  <!-- Thay đổi giá trị này để build cho phiên bản khác -->
  <RevitVersion>2023</RevitVersion>
</PropertyGroup>

<ItemGroup>
  <!-- Nice3point Revit API - Tự động chọn version theo RevitVersion -->
  <PackageReference Include="Nice3point.Revit.Api.RevitAPI" Version="$(RevitVersion).*" />
  <PackageReference Include="Nice3point.Revit.Api.RevitAPIUI" Version="$(RevitVersion).*" />
  <PackageReference Include="Nice3point.Revit.Api.AdWindows" Version="$(RevitVersion).*" />
</ItemGroup>
```

### Sau đó XÓA các Reference cũ:

Tìm và xóa các dòng này trong `.csproj`:

```xml
<!-- XÓA CÁC DÒNG NÀY -->
<Reference Include="AdWindows">
  <HintPath>C:\Program Files\Autodesk\Revit 2023\AdWindows.dll</HintPath>
  <Private>False</Private>
</Reference>
<Reference Include="RevitAPI">
  <HintPath>C:\Program Files\Autodesk\Revit 2023\RevitAPI.dll</HintPath>
  <Private>False</Private>
</Reference>
<Reference Include="RevitAPIUI">
  <HintPath>C:\Program Files\Autodesk\Revit 2023\RevitAPIUI.dll</HintPath>
  <Private>False</Private>
</Reference>
```

---

## 🎯 Sau khi cập nhật

### 1. Restore packages
```powershell
.\nuget.exe restore
```

### 2. Clean và Rebuild
```powershell
# Clean
msbuild Quoc_MEP.csproj /t:Clean

# Rebuild
msbuild Quoc_MEP.csproj /t:Rebuild /p:Configuration=Release
```

### 3. Kiểm tra references
Mở Visual Studio > Solution Explorer > References
- ✅ Phải thấy: `RevitAPI`, `RevitAPIUI`, `AdWindows` từ NuGet
- ❌ Không còn: HintPath tới `C:\Program Files\Autodesk\Revit 2023\`

---

## 🔄 Build cho nhiều phiên bản Revit

### Cách 1: Thay đổi RevitVersion trong .csproj

```xml
<RevitVersion>2024</RevitVersion>  <!-- Thay 2023 -> 2024 -->
```

Sau đó rebuild.

### Cách 2: Dùng MSBuild parameters

```powershell
# Build cho Revit 2023
msbuild Quoc_MEP.csproj /p:RevitVersion=2023 /p:Configuration=Release

# Build cho Revit 2024
msbuild Quoc_MEP.csproj /p:RevitVersion=2024 /p:Configuration=Release

# Build cho Revit 2025
msbuild Quoc_MEP.csproj /p:RevitVersion=2025 /p:Configuration=Release
```

### Cách 3: Script tự động build tất cả versions

Tạo file `build-all-versions.ps1`:

```powershell
$versions = @("2020", "2021", "2022", "2023", "2024", "2025", "2026")

foreach ($version in $versions) {
    Write-Host "Building for Revit $version..." -ForegroundColor Green
    
    msbuild Quoc_MEP.csproj `
        /p:RevitVersion=$version `
        /p:Configuration=Release `
        /t:Rebuild
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Revit $version build successful" -ForegroundColor Green
    } else {
        Write-Host "✗ Revit $version build failed" -ForegroundColor Red
    }
}
```

Chạy:
```powershell
.\build-all-versions.ps1
```

---

## 📋 Packages đã cập nhật trong packages.config

```xml
<!-- Nice3point Revit API Packages - Hỗ trợ đa phiên bản Revit -->
<package id="Nice3point.Revit.Api.AdWindows" version="2023.0.0" targetFramework="net48" />
<package id="Nice3point.Revit.Api.RevitAPI" version="2023.0.0" targetFramework="net48" />
<package id="Nice3point.Revit.Api.RevitAPIUI" version="2023.0.0" targetFramework="net48" />
```

---

## ✨ Lợi ích

### Trước khi dùng Nice3point.Revit.Api:
```xml
<!-- Phải thay đổi path khi build cho phiên bản khác -->
<Reference Include="RevitAPI">
  <HintPath>C:\Program Files\Autodesk\Revit 2023\RevitAPI.dll</HintPath>
</Reference>
```
❌ Hardcode path  
❌ Phải cài Revit để build  
❌ Khó maintain nhiều versions  

### Sau khi dùng Nice3point.Revit.Api:
```xml
<PackageReference Include="Nice3point.Revit.Api.RevitAPI" Version="$(RevitVersion).*" />
```
✅ Tự động từ NuGet  
✅ Build mà không cần cài Revit  
✅ Chỉ cần thay 1 biến `RevitVersion`  

---

## 🔗 Links tham khảo

- [Nice3point/RevitApi GitHub](https://github.com/Nice3point/RevitApi)
- [NuGet Package - RevitAPI](https://www.nuget.org/packages/Nice3point.Revit.Api.RevitAPI)
- [NuGet Package - RevitAPIUI](https://www.nuget.org/packages/Nice3point.Revit.Api.RevitAPIUI)
- [NuGet Package - AdWindows](https://www.nuget.org/packages/Nice3point.Revit.Api.AdWindows)

---

## ❓ Troubleshooting

### Lỗi: "Package not found"
```
Package 'Nice3point.Revit.Api.RevitAPI 2023.0.0' is not found
```

**Giải pháp**: Dùng version chính xác. Ví dụ:
- `2023.1.0` thay vì `2023.0.0`
- Hoặc dùng wildcard: `2023.*`

### Lỗi: "Could not resolve reference"
```
Could not resolve this reference. Could not locate the assembly "RevitAPI"
```

**Giải pháp**:
1. Restore packages: `nuget restore`
2. Restart Visual Studio
3. Clean và Rebuild solution

### Build chậm hoặc lỗi NuGet
**Giải pháp**: Xóa folder `packages` và restore lại:
```powershell
Remove-Item -Path "packages" -Recurse -Force
.\nuget.exe restore
```

---

## 📝 Checklist

- [ ] Đã xóa Reference cũ (RevitAPI, RevitAPIUI, AdWindows)
- [ ] Đã cài Nice3point.Revit.Api packages
- [ ] Đã update packages.config
- [ ] Đã test build thành công
- [ ] Đã test với plugin trong Revit
- [ ] Đã commit changes vào Git

---

**Cập nhật**: November 2, 2025  
**Hỗ trợ**: Revit 2020-2026  
**Packages**: Nice3point.Revit.Api v2026.3.0 (mới nhất)
