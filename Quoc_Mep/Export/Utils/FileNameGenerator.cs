using Autodesk.Revit.DB;
using System.IO;
using System.Text.RegularExpressions;
using Quoc_MEP.Export.Models;
using Quoc_MEP.Export.Utils;

namespace Quoc_MEP.Export.Utils
{
    public static class FileNameGenerator
    {
        public static string GenerateFileName(ViewSheet sheet, Document doc, string template, string extension)
        {
            string fileName = template;
            
            // Thay thế các placeholder cơ bản
            fileName = fileName.Replace("{SheetNumber}", sheet.SheetNumber ?? "");
            fileName = fileName.Replace("{SheetName}", sheet.Name ?? "");
            
            // Project information
            ProjectInfo projectInfo = doc.ProjectInformation;
            fileName = fileName.Replace("{ProjectNumber}", 
                ParameterUtils.GetParameterValue(projectInfo, BuiltInParameter.PROJECT_NUMBER));
            fileName = fileName.Replace("{ProjectName}", 
                ParameterUtils.GetParameterValue(projectInfo, BuiltInParameter.PROJECT_NAME));
            fileName = fileName.Replace("{ClientName}", 
                ParameterUtils.GetParameterValue(projectInfo, "Client Name"));
            fileName = fileName.Replace("{Author}", 
                ParameterUtils.GetParameterValue(projectInfo, BuiltInParameter.PROJECT_AUTHOR));
            
            // Revision information
            Parameter revisionParam = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION);
            string revision = revisionParam?.AsString() ?? "";
            fileName = fileName.Replace("{Revision}", revision);
            fileName = fileName.Replace("{Rev}", revision);
            
            // Date and time
            fileName = fileName.Replace("{Date}", System.DateTime.Now.ToString("yyyy-MM-dd"));
            fileName = fileName.Replace("{Time}", System.DateTime.Now.ToString("HH-mm-ss"));
            fileName = fileName.Replace("{DateTime}", System.DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss"));
            
            // User and computer
            fileName = fileName.Replace("{User}", System.Environment.UserName);
            fileName = fileName.Replace("{Computer}", System.Environment.MachineName);
            
            // Custom parameters - tìm tất cả {ParamName} trong template
            var matches = Regex.Matches(template, @"\{([^}]+)\}");
            foreach (Match match in matches)
            {
                string paramName = match.Groups[1].Value;
                
                // Bỏ qua các placeholder đã xử lý
                if (IsStandardPlaceholder(paramName))
                    continue;
                
                // Tìm custom parameter
                Parameter customParam = sheet.LookupParameter(paramName);
                if (customParam != null && customParam.HasValue)
                {
                    string paramValue = ParameterUtils.GetParameterValueAsString(customParam);
                    fileName = fileName.Replace($"{{{paramName}}}", paramValue);
                }
                else
                {
                    // Nếu không tìm thấy parameter, thay thế bằng chuỗi rỗng
                    fileName = fileName.Replace($"{{{paramName}}}", "");
                }
            }
            
            // Làm sạch tên file
            fileName = SanitizeFileName(fileName);
            
            // Thêm extension
            if (!string.IsNullOrEmpty(extension))
            {
                fileName = $"{fileName}.{extension}";
            }
            
            return fileName;
        }
        
        private static bool IsStandardPlaceholder(string placeholder)
        {
            string[] standardPlaceholders = {
                "SheetNumber", "SheetName", "ProjectNumber", "ProjectName", 
                "ClientName", "Author", "Revision", "Rev", "Date", "Time", 
                "DateTime", "User", "Computer"
            };
            
            return System.Array.Exists(standardPlaceholders, p => p.Equals(placeholder, System.StringComparison.OrdinalIgnoreCase));
        }
        
        public static string SanitizeFileName(string fileName)
        {
            // Loại bỏ ký tự không hợp lệ trong tên file
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                fileName = fileName.Replace(c, '_');
            }
            
            // Loại bỏ ký tự đặc biệt khác
            fileName = fileName.Replace(":", "_");
            fileName = fileName.Replace(";", "_");
            fileName = fileName.Replace(",", "_");
            
            // Thay thế nhiều dấu gạch dưới liên tiếp bằng một dấu
            fileName = Regex.Replace(fileName, "_+", "_");
            
            // Loại bỏ dấu gạch dưới ở đầu và cuối
            fileName = fileName.Trim('_');
            
            // Thay thế khoảng trắng bằng dấu gạch dưới nếu cần
            fileName = fileName.Replace(" ", "_");
            
            return fileName;
        }
        
        public static string GenerateSubfolderPath(ViewSheet sheet, Document doc, Quoc_MEP.Export.Models.PSExportSettings settings)
        {
            if (!settings.CreateSubfolders || string.IsNullOrEmpty(settings.SubfolderTemplate))
            {
                return settings.OutputFolder;
            }
            
            string subfolderName = settings.SubfolderTemplate;
            
            // Các placeholder cho subfolder
            subfolderName = subfolderName.Replace("{DrawingType}", GetDrawingType(sheet));
            subfolderName = subfolderName.Replace("{Discipline}", GetDiscipline(sheet));
            subfolderName = subfolderName.Replace("{Level}", GetLevel(sheet));
            subfolderName = subfolderName.Replace("{Phase}", GetPhase(sheet));
            
            // Thay thế các placeholder cơ bản
            subfolderName = GenerateFileName(sheet, doc, subfolderName, "");
            
            // Làm sạch tên subfolder
            subfolderName = SanitizeFileName(subfolderName);
            
            return Path.Combine(settings.OutputFolder, subfolderName);
        }
        
        private static string GetDrawingType(ViewSheet sheet)
        {
            // Phân loại dựa trên sheet number prefix
            string sheetNumber = sheet.SheetNumber ?? "";
            
            if (sheetNumber.StartsWith("A"))
                return "Architectural";
            else if (sheetNumber.StartsWith("S"))
                return "Structural";
            else if (sheetNumber.StartsWith("M"))
                return "Mechanical";
            else if (sheetNumber.StartsWith("E"))
                return "Electrical";
            else if (sheetNumber.StartsWith("P"))
                return "Plumbing";
            else if (sheetNumber.StartsWith("C"))
                return "Civil";
            else if (sheetNumber.StartsWith("L"))
                return "Landscape";
            else
                return "General";
        }
        
        private static string GetDiscipline(ViewSheet sheet)
        {
            // Lấy từ parameter discipline nếu có
            Parameter disciplineParam = sheet.LookupParameter("Discipline");
            if (disciplineParam != null && disciplineParam.HasValue)
            {
                return disciplineParam.AsString();
            }
            
            return GetDrawingType(sheet);
        }
        
        private static string GetLevel(ViewSheet sheet)
        {
            // Tìm parameter level trong sheet
            Parameter levelParam = sheet.LookupParameter("Level") ?? 
                                  sheet.LookupParameter("Floor") ??
                                  sheet.LookupParameter("Storey");
            
            if (levelParam != null && levelParam.HasValue)
            {
                return levelParam.AsString();
            }
            
            return "All_Levels";
        }
        
        private static string GetPhase(ViewSheet sheet)
        {
            // Tìm parameter phase trong sheet
            Parameter phaseParam = sheet.LookupParameter("Phase") ??
                                  sheet.LookupParameter("Construction Phase");
            
            if (phaseParam != null && phaseParam.HasValue)
            {
                return phaseParam.AsString();
            }
            
            return "Current";
        }
    }
}