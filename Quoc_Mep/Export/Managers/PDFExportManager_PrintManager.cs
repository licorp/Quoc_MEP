using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Autodesk.Revit.DB;
using Quoc_MEP.Export.Models;

namespace Quoc_MEP.Export.Managers
{
    /// <summary>
    /// PDF Export Manager using PrintManager API (minimalist approach)
    /// This version uses PrintManager to support advanced PDF settings (Color, Zoom, Raster Quality)
    /// </summary>
    public class PDFExportManagerPrintManager
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
        private static extern void OutputDebugStringA(string lpOutputString);

        private readonly Document _document;

        public PDFExportManagerPrintManager(Document document)
        {
            _document = document;
        }

        /// <summary>
        /// Export sheets using PrintManager API (supports Color, Zoom, Raster Quality settings)
        /// Uses minimalist approach to avoid C++ runtime errors
        /// </summary>
        public bool ExportSheetsWithPrintManager(List<SheetItem> sheetItems, string outputFolder, ExportSettings settings, Action<int, int, string> progressCallback = null)
        {
            try
            {
                WriteDebugLog($"[PrintManager] Starting PDF export for {sheetItems.Count} sheets");
                WriteDebugLog($"[PrintManager] Output folder: {outputFolder}");

                // Ensure output directory exists
                Directory.CreateDirectory(outputFolder);

                int successCount = 0;
                int failCount = 0;
                int total = sheetItems.Count;

                // Export each sheet ONE AT A TIME (minimalist approach)
                for (int i = 0; i < total; i++)
                {
                    var sheetItem = sheetItems[i];
                    
                    try
                    {
                        // Report progress
                        progressCallback?.Invoke(i + 1, total, sheetItem.Number);

                        // Get ViewSheet from document
                        ViewSheet sheet = _document.GetElement(sheetItem.Id) as ViewSheet;
                        if (sheet == null)
                        {
                            WriteDebugLog($"[PrintManager] ERROR: Cannot find sheet {sheetItem.Number}");
                            failCount++;
                            continue;
                        }

                        WriteDebugLog($"[PrintManager] Exporting sheet {i + 1}/{total}: {sheet.SheetNumber} - {sheet.Name}");

                        // Get custom filename (use same naming as Document.Export method)
                        string customFileName = sheetItem.CustomFileName;
                        if (string.IsNullOrEmpty(customFileName))
                        {
                            customFileName = $"{sheet.SheetNumber}_{sheet.Name}";
                        }

                        // Export THIS sheet using PrintManager
                        bool exportSuccess = ExportSingleSheetWithPrintManager(sheet, outputFolder, customFileName, settings);

                        if (exportSuccess)
                        {
                            successCount++;
                            WriteDebugLog($"[PrintManager] ✓ Exported: {customFileName}.pdf");
                        }
                        else
                        {
                            failCount++;
                            WriteDebugLog($"[PrintManager] ✗ Failed: {customFileName}.pdf");
                        }
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        WriteDebugLog($"[PrintManager] ERROR exporting {sheetItem.Number}: {ex.Message}");
                    }
                }

                WriteDebugLog($"[PrintManager] Export complete: {successCount} succeeded, {failCount} failed");
                return successCount > 0;
            }
            catch (Exception ex)
            {
                WriteDebugLog($"[PrintManager] FATAL ERROR: {ex.Message}");
                WriteDebugLog($"[PrintManager] Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Export a single sheet using PrintManager (minimalist approach to avoid errors)
        /// </summary>
        private bool ExportSingleSheetWithPrintManager(ViewSheet sheet, string outputFolder, string fileName, ExportSettings settings)
        {
            PrintManager pm = _document.PrintManager;
            
            try
            {
                WriteDebugLog($"[PrintManager] Configuring PrintManager for sheet: {sheet.SheetNumber}");

                // ===================================================================
                // STEP 1: Configure PrintManager to print THIS sheet
                // Use Transaction to create a temporary saved ViewSheetSet
                // ===================================================================
                try
                {
                    using (Transaction trans = new Transaction(_document, "Configure Print ViewSet"))
                    {
                        trans.Start();
                        
                        // Create a temporary ViewSheetSet for this export
                        ViewSheetSetting vss = pm.ViewSheetSetting;
                        
                        // Save current views to restore later (if needed)
                        vss.SaveAs($"_TempPrintSet_{sheet.Id}");
                        
                        // Get the ViewSet and manipulate it
                        ViewSet viewSet = vss.CurrentViewSheetSet.Views;
                        viewSet.Clear();
                        viewSet.Insert(sheet);
                        
                        trans.Commit();
                    }
                    
                    WriteDebugLog($"[PrintManager] ✓ ViewSet configured for sheet: {sheet.SheetNumber}");
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"[PrintManager] ERROR configuring ViewSet: {ex.Message}");
                    // Try fallback: use PrintRange.Current
                    WriteDebugLog($"[PrintManager] Falling back to PrintRange.Current");
                    pm.PrintRange = PrintRange.Current;
                    pm.PrintToFile = true;
                    pm.CombinedFile = true;
                    WriteDebugLog($"[PrintManager] PrintRange: Current");
                    goto skip_viewset;
                }
                
                pm.PrintRange = PrintRange.Select; // Print selected views  
                pm.PrintToFile = true;
                pm.CombinedFile = false; // Export individual file
                WriteDebugLog($"[PrintManager] PrintRange: Select (single sheet)");
                
                skip_viewset:
                
                WriteDebugLog($"[PrintManager] PrintRange: {pm.PrintRange}");
                WriteDebugLog($"[PrintManager] PrintToFile: {pm.PrintToFile}");
                WriteDebugLog($"[PrintManager] CombinedFile: {pm.CombinedFile}");

                // ===================================================================
                // STEP 3: Apply PDF settings (Color, Zoom, Raster Quality, etc.)
                // This is what Document.Export() CAN'T do
                // ===================================================================
                
                // Apply view options FIRST (hide categories, crop boxes)
                try
                {
                    using (Transaction trans = new Transaction(_document, "Apply View Options"))
                    {
                        trans.Start();
                        Services.PDFOptionsApplier.ApplyViewOptionsToSheetNoTransaction(_document, sheet, settings);
                        trans.Commit();
                    }
                    WriteDebugLog($"[PrintManager] ✓ View options applied");
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"[PrintManager] WARNING: Could not apply view options: {ex.Message}");
                }

                // Apply advanced PrintManager settings (Color, Zoom, etc.)
                try
                {
                    Services.PDFOptionsApplier.ApplyPrintManagerSettings(pm, settings);
                    WriteDebugLog($"[PrintManager] ✓ PrintManager settings applied");
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"[PrintManager] WARNING: Could not apply PrintManager settings: {ex.Message}");
                }

                // ===================================================================
                // STEP 4: Apply settings and prepare for export
                // ===================================================================
                try
                {
                    pm.Apply();
                    WriteDebugLog($"[PrintManager] ✓ PrintManager.Apply() succeeded");
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"[PrintManager] ERROR in Apply(): {ex.Message}");
                    return false;
                }

                // ===================================================================
                // STEP 5: Set output file path
                // ===================================================================
                string outputPath = Path.Combine(outputFolder, fileName + ".pdf");
                
                // Delete existing file if present
                if (File.Exists(outputPath))
                {
                    try
                    {
                        File.Delete(outputPath);
                        WriteDebugLog($"[PrintManager] Deleted existing file: {fileName}.pdf");
                    }
                    catch (Exception ex)
                    {
                        WriteDebugLog($"[PrintManager] WARNING: Could not delete existing file: {ex.Message}");
                    }
                }

                pm.PrintToFileName = outputPath;
                WriteDebugLog($"[PrintManager] Output path: {outputPath}");

                // ===================================================================
                // STEP 6: Submit print job
                // ===================================================================
                WriteDebugLog($"[PrintManager] Calling SubmitPrint()...");
                
                bool submitResult = pm.SubmitPrint();
                
                WriteDebugLog($"[PrintManager] SubmitPrint() returned: {submitResult}");

                // Wait for file to be written
                System.Threading.Thread.Sleep(1000);

                // Verify file was created
                if (File.Exists(outputPath))
                {
                    FileInfo fi = new FileInfo(outputPath);
                    WriteDebugLog($"[PrintManager] ✓ File created: {fi.Name} ({fi.Length} bytes)");
                    return true;
                }
                else
                {
                    WriteDebugLog($"[PrintManager] ✗ File NOT created: {outputPath}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"[PrintManager] ERROR: {ex.Message}");
                WriteDebugLog($"[PrintManager] Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        private void WriteDebugLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string logMessage = $"[ExportPlus PrintMgr] {timestamp} - {message}";
            OutputDebugStringA(logMessage);
        }
    }
}
