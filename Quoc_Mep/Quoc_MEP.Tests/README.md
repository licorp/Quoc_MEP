# Quoc_MEP.Tests - Test Project

## 📋 Giới thiệu

Project này chứa các unit tests cho Revit Add-in **Quoc_MEP**. Được thiết lập với NUnit framework và sẵn sàng để tích hợp **RevitTestLibrary**.

## 🚀 Đã cài đặt

✅ **NUnit 3.14.0** - Framework để viết và chạy tests  
✅ **NUnit3TestAdapter 4.5.0** - Để chạy tests trong Visual Studio Test Explorer  
✅ **Test Project Structure** - Cấu trúc project test hoàn chỉnh

## 📦 Cài đặt RevitTestLibrary (Bước tiếp theo)

Để sử dụng đầy đủ khả năng mock Revit objects, bạn cần cài thêm **RevitTestLibrary**:

### Cách 1: Dùng NuGet Package Manager Console
```powershell
Install-Package RevitTestLibrary -ProjectName Quoc_MEP.Tests
```

### Cách 2: Dùng .NET CLI
```bash
cd "d:\RevitAPI_tu viet\RevitAPIMEP\Quoc_MEP.Tests"
dotnet add package RevitTestLibrary
```

### Cách 3: Thêm thủ công vào packages.config
Thêm dòng này vào file `packages.config`:
```xml
<package id="RevitTestLibrary" version="1.0.0" targetFramework="net48" />
```
Sau đó chạy: `nuget restore`

## 📂 Cấu trúc Project

```
Quoc_MEP.Tests/
├── Properties/
│   └── AssemblyInfo.cs
├── SampleTests/
│   └── BasicRevitTests.cs          # Tests mẫu đã sẵn sàng
├── packages.config                  # NuGet packages
├── Quoc_MEP.Tests.csproj           # Project file
└── README.md                        # File này
```

## 🧪 Chạy Tests

### Trong Visual Studio:
1. Mở **Test Explorer** (Test > Test Explorer hoặc Ctrl+E, T)
2. Click "Run All" để chạy tất cả tests
3. Xem kết quả trong Test Explorer

### Dùng Command Line:
```powershell
# Từ thư mục solution
dotnet test
```

### Chạy tests cụ thể:
```powershell
dotnet test --filter "TestCategory=Basic"
dotnet test --filter "FullyQualifiedName~BasicRevitTests"
```

## 📝 Tests mẫu có sẵn

File `SampleTests/BasicRevitTests.cs` chứa:

### ✅ BasicRevitTests
- `TestXYZCreation` - Test tạo điểm XYZ
- `TestDistanceCalculation` - Test tính khoảng cách
- `TestVectorAddition` - Test cộng vector
- `TestFeetToMillimeters` - Test chuyển đổi đơn vị
- `TestMillimetersToFeet` - Test chuyển đổi đơn vị ngược

### ✅ UtilityTests
- `TestNullValidation` - Test kiểm tra null
- `TestCollectionOperations` - Test với collections

## 🎯 Viết Test mới

### Template cơ bản:
```csharp
[Test]
[Category("YourCategory")]
[Description("Mô tả test của bạn")]
public void TestYourFeature()
{
    // Arrange - Chuẩn bị data
    var input = "test data";
    
    // Act - Thực hiện action cần test
    var result = YourMethod(input);
    
    // Assert - Kiểm tra kết quả
    Assert.AreEqual("expected", result);
}
```

### Với RevitTestLibrary (sau khi cài):
```csharp
[Test]
public void TestWithMockDocument()
{
    // Arrange
    var mockDoc = new MockDocument();
    var mockApp = new MockApplication();
    
    // Act
    var result = YourRevitCommand.Execute(mockDoc);
    
    // Assert
    Assert.IsTrue(result.Success);
}
```

## 🔧 Debugging Tests

1. Đặt breakpoint trong test code
2. Right-click test trong Test Explorer
3. Chọn "Debug Selected Tests"
4. Debug như code bình thường

## 📊 Test Categories

Tests được phân loại theo categories:
- **Basic** - Tests cơ bản cho Revit API types
- **Conversion** - Tests chuyển đổi đơn vị
- **Utility** - Tests cho utility functions
- **Mock** - Tests dùng mock objects (cần RevitTestLibrary)

Chạy theo category:
```powershell
dotnet test --filter "TestCategory=Basic"
```

## 🎨 Best Practices

1. **Đặt tên test rõ ràng**: `Test_MethodName_Scenario_ExpectedResult`
2. **Một test một mục đích**: Mỗi test chỉ test một điều
3. **AAA Pattern**: Arrange, Act, Assert
4. **Sử dụng Categories**: Phân loại tests để dễ quản lý
5. **Viết Description**: Giải thích mục đích của test

## 🐛 Troubleshooting

### Tests không hiện trong Test Explorer?
- Build lại solution (Ctrl+Shift+B)
- Restart Visual Studio
- Check NUnit3TestAdapter đã được cài đúng

### Build errors liên quan đến Revit DLLs?
- Đảm bảo Revit 2023 đã được cài đặt
- Check đường dẫn RevitAPI.dll trong project file

### Tests fail với "Document not available"?
- Cần dùng mock objects từ RevitTestLibrary
- Uncomment code template trong BasicRevitTests.cs

## 📚 Tài liệu tham khảo

- [NUnit Documentation](https://docs.nunit.org/)
- [RevitTestLibrary GitHub](https://github.com/NeVeSpl/RevitTestLibrary)
- [Revit API Developer Guide](https://www.revitapidocs.com/)

## 🎓 Next Steps

1. ✅ Cài đặt **RevitTestLibrary** package
2. ✅ Uncomment mock test templates trong `BasicRevitTests.cs`
3. ✅ Viết tests cho features hiện tại của bạn
4. ✅ Chạy tests trong CI/CD pipeline
5. ✅ Tích hợp code coverage reports

---

**Lưu ý**: Project này sử dụng .NET Framework 4.8 để tương thích với Revit 2023.
