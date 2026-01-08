using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace Quoc_MEP
{
    /// <summary>
    /// Universal Loader - Tự động load đúng DLL cho từng phiên bản Revit
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class UniversalLoader : IExternalApplication
    {
        private static Assembly _loadedAssembly;
        private static IExternalApplication _actualApp;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Lấy version của Revit đang chạy
                string revitVersion = application.ControlledApplication.VersionNumber;
                
                // Lấy thư mục chứa loader
                string loaderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                
                // Xác định thư mục DLL theo version
                string dllFolder = Path.Combine(loaderPath, $"Revit{revitVersion}");
                string dllPath = Path.Combine(dllFolder, "Quoc_MEP.dll");
                
                // Kiểm tra file tồn tại
                if (!File.Exists(dllPath))
                {
                    TaskDialog.Show("Quoc MEP Error", 
                        $"Không tìm thấy DLL cho Revit {revitVersion}\n" +
                        $"Đường dẫn: {dllPath}\n\n" +
                        $"Vui lòng kiểm tra lại cấu trúc thư mục.");
                    return Result.Failed;
                }
                
                // Load assembly
                _loadedAssembly = Assembly.LoadFrom(dllPath);
                
                // Tìm class Ribbon (IExternalApplication)
                Type ribbonType = _loadedAssembly.GetType("Quoc_MEP.Ribbon");
                if (ribbonType == null)
                {
                    TaskDialog.Show("Quoc MEP Error", 
                        $"Không tìm thấy class 'Ribbon' trong DLL\n" +
                        $"File: {dllPath}");
                    return Result.Failed;
                }
                
                // Tạo instance và gọi OnStartup
                _actualApp = Activator.CreateInstance(ribbonType) as IExternalApplication;
                if (_actualApp == null)
                {
                    TaskDialog.Show("Quoc MEP Error", 
                        "Không thể tạo instance của Ribbon class");
                    return Result.Failed;
                }
                
                // Gọi OnStartup của Ribbon thật
                Result result = _actualApp.OnStartup(application);
                
                if (result == Result.Succeeded)
                {
                    // Optional: Hiển thị thông báo thành công (có thể bỏ comment nếu muốn)
                    // TaskDialog.Show("Quoc MEP", 
                    //     $"Loaded successfully for Revit {revitVersion}");
                }
                
                return result;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Quoc MEP Error", 
                    $"Lỗi khi load add-in:\n\n" +
                    $"{ex.Message}\n\n" +
                    $"StackTrace:\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            try
            {
                if (_actualApp != null)
                {
                    return _actualApp.OnShutdown(application);
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Quoc MEP Error", 
                    $"Lỗi khi shutdown:\n{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
