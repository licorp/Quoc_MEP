using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Quoc_MEP.Export.Models;
using Quoc_MEP.Export.Managers;

namespace Quoc_MEP.Export.Commands
{
    /// <summary>
    /// External Event Handler for Export operations
    /// </summary>
    public class ExportHandler : IExternalEventHandler
    {
        public Document Document { get; set; }
        public List<SheetItem> SheetsToExport { get; set; }
        public List<string> Formats { get; set; }
        public string OutputFolder { get; set; }
        public ExportSettings ExportSettings { get; set; }
        public Action<int, int, string, bool> ProgressCallback { get; set; }
        public Action<bool, string> CompletionCallback { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                if (Document == null || SheetsToExport == null || Formats == null)
                {
                    CompletionCallback?.Invoke(false, "Invalid export parameters");
                    return;
                }

                bool overallSuccess = true;
                string errorMessage = "";

                foreach (var format in Formats)
                {
                    try
                    {
                        if (format.ToUpper() == "PDF")
                        {
#if REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                            var pdfManager = new PDFExportManager(Document);
                            bool result = pdfManager.ExportSheetsWithCustomNames(
                                SheetsToExport,
                                OutputFolder,
                                ExportSettings,
                                ProgressCallback);

                            if (!result)
#else
                            // PDFExportManager not available in Revit 2020-2022
                            errorMessage += $"PDF export not supported in Revit {Application.VersionNumber}\n";
                            if (false)
#endif
                            {
                                overallSuccess = false;
                                errorMessage += $"PDF export failed. ";
                            }
                        }
                        else if (format.ToUpper() == "DWG")
                        {
                            // TODO: Implement DWG export
                            errorMessage += $"DWG export not yet implemented. ";
                        }
                        else if (format.ToUpper() == "IFC")
                        {
                            // TODO: Implement IFC export
                            errorMessage += $"IFC export not yet implemented. ";
                        }
                        else
                        {
                            errorMessage += $"Format {format} not supported. ";
                        }
                    }
                    catch (Exception ex)
                    {
                        overallSuccess = false;
                        errorMessage += $"Error exporting {format}: {ex.Message}. ";
                    }
                }

                CompletionCallback?.Invoke(overallSuccess, errorMessage);
            }
            catch (Exception ex)
            {
                CompletionCallback?.Invoke(false, $"Critical error: {ex.Message}");
            }
        }

        public string GetName()
        {
            return "Export + Handler";
        }
    }
}
