#if REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Quoc_MEP.Export.Models;
using Quoc_MEP.Export.Managers;

namespace Quoc_MEP.Export.Events
{
    /// <summary>
    /// External Event Handler for PDF Export
    /// This allows PDF export to run in proper Revit API context with Transaction support
    /// ⚠️ ONLY available in Revit 2023+
    /// </summary>
    public class PDFExportEventHandler : IExternalEventHandler
    {
        // Export parameters set by UI
        public Document Document { get; set; }
        public List<SheetItem> SheetItems { get; set; }
        public string OutputFolder { get; set; }
        public ExportSettings Settings { get; set; }
        public Action<int, int, string, bool> ProgressCallback { get; set; }
        
        // Export result
        public bool ExportResult { get; private set; }
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Execute method - called when event is raised
        /// Runs in proper Revit API context with Transaction support
        /// </summary>
        public void Execute(UIApplication app)
        {
            try
            {
                WriteDebugLog("=== PDF Export External Event STARTED ===");
                WriteDebugLog($"Running in API context - Thread ID: {System.Threading.Thread.CurrentThread.ManagedThreadId}");

                if (Document == null)
                {
                    ErrorMessage = "Document is null";
                    WriteDebugLog($"ERROR: {ErrorMessage}");
                    ExportResult = false;
                    return;
                }

                if (SheetItems == null || SheetItems.Count == 0)
                {
                    ErrorMessage = "No sheets to export";
                    WriteDebugLog($"ERROR: {ErrorMessage}");
                    ExportResult = false;
                    return;
                }

                // Create PDF Export Manager
                var pdfManager = new PDFExportManager(Document);

                // Execute export (now with Transaction support)
                WriteDebugLog("Calling ExportSheetsWithCustomNames...");
                ExportResult = pdfManager.ExportSheetsWithCustomNames(
                    SheetItems,
                    OutputFolder,
                    Settings,
                    ProgressCallback
                );

                WriteDebugLog($"Export completed with result: {ExportResult}");
                
                if (!ExportResult)
                {
                    ErrorMessage = "Export failed - check debug log for details";
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Exception in PDF export: {ex.Message}";
                WriteDebugLog($"EXCEPTION: {ErrorMessage}");
                WriteDebugLog($"Stack trace: {ex.StackTrace}");
                ExportResult = false;
            }
            finally
            {
                WriteDebugLog("=== PDF Export External Event COMPLETED ===");
            }
        }

        /// <summary>
        /// Name of the external event
        /// </summary>
        public string GetName()
        {
            return "ExportPlus PDF Export Event";
        }

        /// <summary>
        /// Write debug log (using same mechanism as other classes)
        /// </summary>
        private void WriteDebugLog(string message)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                string fullMessage = $"[Export +] {timestamp} - {message}";
                System.Diagnostics.Debug.WriteLine(fullMessage);
            }
            catch { /* Ignore logging errors */ }
        }
    }
}
#endif
