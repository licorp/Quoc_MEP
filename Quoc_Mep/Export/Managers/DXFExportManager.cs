using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Quoc_MEP.Export.Models;
using ricaun.Revit.UI.StatusBar; // ✅ ADD: StatusBar package for progress display

namespace Quoc_MEP.Export.Managers
{
    /// <summary>
    /// DXF Export Manager based on DA4R-DxfExporter
    /// Supports exporting 3D, Plan, Section, and Sheet views to DXF format
    /// </summary>
    public class DXFExportManager
    {
        private readonly Document _document;

        public DXFExportManager(Document document)
        {
            _document = document;
        }

        /// <summary>
        /// Export views to DXF format (DA4R Pattern)
        /// </summary>
        /// <param name="outputFolder">Output folder path</param>
        /// <param name="settings">DXF export settings</param>
        /// <param name="progressCallback">Progress callback (current, total, viewName, success)</param>
        /// <returns>True if export succeeded</returns>
        public bool ExportViewsToDXF(string outputFolder, PSDXFExportSettings settings, Action<int, int, string, bool> progressCallback = null)
        {
            try
            {
                LogTrace("===== DA4R-DxfExporter Pattern: Starting DXF Export =====");
                LogTrace($"Output folder: {outputFolder}");

                // Ensure output directory exists
                Directory.CreateDirectory(outputFolder);

                // Collect views based on settings
                var viewIds = CollectViewsForExport(settings);

                if (viewIds == null || viewIds.Count == 0)
                {
                    LogTrace("⚠ No views found for export");
                    return false;
                }

                LogTrace($"✓ Collected {viewIds.Count} views for export");

                // Create DXF export options
                var exportOptions = new DXFExportOptions();

                // Determine file prefix
                string filePrefix = settings.UseDocumentTitle ? _document.Title : 
                                    (!string.IsNullOrEmpty(settings.CustomFilePrefix) ? settings.CustomFilePrefix : "Export");
                
                // Sanitize file prefix
                filePrefix = SanitizeFileName(filePrefix);

                LogTrace($"File prefix: {filePrefix}");
                LogTrace("Starting export...");

                // ✅ ADD: StatusBar progress for DXF export (batch export)
                bool exportSuccess = false;
                RevitProgressBarUtils.Run($"Exporting {viewIds.Count} views to DXF", viewIds.Count, (i) =>
                {
                    // DXF exports all views at once (batch), so this simulates progress
                    if (i == viewIds.Count - 1) // Last iteration
                    {
                        // Export all views with single API call (DA4R pattern)
                        _document.Export(outputFolder, filePrefix, viewIds, exportOptions);
                        exportSuccess = true;
                    }
                    System.Threading.Thread.Sleep(10); // Simulate progress for visual feedback
                });

                LogTrace($"✅ DXF Export completed successfully");
                LogTrace($"Output: {outputFolder}\\{filePrefix}*.dxf");
                
                // Report success
                progressCallback?.Invoke(viewIds.Count, viewIds.Count, "All views", exportSuccess);

                return exportSuccess;
            }
            catch (Autodesk.Revit.Exceptions.InvalidPathArgumentException ex)
            {
                LogTrace($"❌ Invalid path: {ex.Message}");
                return false;
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException ex)
            {
                LogTrace($"❌ Invalid argument: {ex.Message}");
                return false;
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
            {
                LogTrace($"❌ Invalid operation: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                LogTrace($"❌ Export failed: {ex.Message}");
                LogTrace($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Collect views for export based on DA4R-DxfExporter pattern
        /// </summary>
        private List<ElementId> CollectViewsForExport(PSDXFExportSettings settings)
        {
            LogTrace("Collecting views for export...");

            var viewElemIds = new List<ElementId>();

            try
            {
                if (settings.ExportAllViews)
                {
                    LogTrace("Mode: Export All Views (with filters)");

                    // Collect 3D Views
                    if (settings.Export3DViews)
                    {
                        using (var collector = new FilteredElementCollector(_document))
                        {
                            var view3dIds = collector
                                .WhereElementIsNotElementType()
                                .OfClass(typeof(View3D))
                                .Cast<View>()
                                .Where(v => !v.IsTemplate || !settings.ExcludeTemplateViews)
                                .Select(v => v.Id);

                            viewElemIds.AddRange(view3dIds);
                            LogTrace($"  + 3D Views: {view3dIds.Count()}");
                        }
                    }

                    // Collect Plan Views
                    if (settings.ExportPlanViews)
                    {
                        using (var collector = new FilteredElementCollector(_document))
                        {
                            var planIds = collector
                                .WhereElementIsNotElementType()
                                .OfClass(typeof(ViewPlan))
                                .Cast<View>()
                                .Where(v => !v.IsTemplate || !settings.ExcludeTemplateViews)
                                .Select(v => v.Id);

                            viewElemIds.AddRange(planIds);
                            LogTrace($"  + Plan Views: {planIds.Count()}");
                        }
                    }

                    // Collect Section Views
                    if (settings.ExportSectionViews)
                    {
                        using (var collector = new FilteredElementCollector(_document))
                        {
                            var sectionIds = collector
                                .WhereElementIsNotElementType()
                                .OfClass(typeof(ViewSection))
                                .Cast<View>()
                                .Where(v => !v.IsTemplate || !settings.ExcludeTemplateViews)
                                .Select(v => v.Id);

                            viewElemIds.AddRange(sectionIds);
                            LogTrace($"  + Section Views: {sectionIds.Count()}");
                        }
                    }

                    // Collect Sheet Views
                    if (settings.ExportSheetViews)
                    {
                        using (var collector = new FilteredElementCollector(_document))
                        {
                            var sheetIds = collector
                                .WhereElementIsNotElementType()
                                .OfClass(typeof(ViewSheet))
                                .Cast<View>()
                                .Where(v => !v.IsTemplate || !settings.ExcludeTemplateViews)
                                .Select(v => v.Id);

                            viewElemIds.AddRange(sheetIds);
                            LogTrace($"  + Sheet Views: {sheetIds.Count()}");
                        }
                    }
                }
                else
                {
                    LogTrace("Mode: Export Selected Views");
                    // In this mode, views would be provided by caller
                    // For now, return empty to indicate selection needed
                    LogTrace("⚠ No views selected for export");
                }

                LogTrace($"Total views collected: {viewElemIds.Count}");
                return viewElemIds;
            }
            catch (Exception ex)
            {
                LogTrace($"Error collecting views: {ex.Message}");
                return new List<ElementId>();
            }
        }

        /// <summary>
        /// Export specific views to DXF
        /// </summary>
        public bool ExportSpecificViews(List<ElementId> viewIds, string outputFolder, string filePrefix, Action<int, int, string, bool> progressCallback = null)
        {
            try
            {
                if (viewIds == null || viewIds.Count == 0)
                {
                    LogTrace("No views provided for export");
                    return false;
                }

                LogTrace($"===== Exporting {viewIds.Count} specific views to DXF =====");

                // Ensure output directory exists
                Directory.CreateDirectory(outputFolder);

                // Create DXF export options
                var exportOptions = new DXFExportOptions();

                // Sanitize file prefix
                filePrefix = SanitizeFileName(filePrefix);

                LogTrace($"Output: {outputFolder}\\{filePrefix}*.dxf");

                // Export
                _document.Export(outputFolder, filePrefix, viewIds, exportOptions);

                LogTrace($"✅ Export completed");
                progressCallback?.Invoke(viewIds.Count, viewIds.Count, "Selected views", true);

                return true;
            }
            catch (Exception ex)
            {
                LogTrace($"❌ Export failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Sanitize filename by replacing invalid characters
        /// </summary>
        private string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return "Export";

            var invalidChars = Path.GetInvalidFileNameChars();
            foreach (var c in invalidChars)
            {
                fileName = fileName.Replace(c, '_');
            }

            // Replace common problematic characters
            fileName = fileName.Replace("{", "_")
                               .Replace("}", "_")
                               .Replace("[", "_")
                               .Replace("]", "_")
                               .Replace(":", "_")
                               .Replace(";", "_")
                               .Replace(",", "_");

            // Remove multiple underscores
            while (fileName.Contains("__"))
            {
                fileName = fileName.Replace("__", "_");
            }

            return fileName.Trim('_');
        }

        /// <summary>
        /// Log trace message
        /// </summary>
        private void LogTrace(string message)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                string fullMessage = $"[DXF Export] {timestamp} - {message}";
                System.Diagnostics.Debug.WriteLine(fullMessage);
            }
            catch { }
        }
    }
}
