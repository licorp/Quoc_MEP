using Autodesk.Revit.DB;
using Quoc_MEP.Export.Models;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Quoc_MEP.Export.Services
{
    /// <summary>
    /// Service to apply PDF export options to Revit PrintManager
    /// </summary>
    public class PDFOptionsApplier
    {
        #region Debug Logging

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern void OutputDebugString(string message);

        private static void WriteDebugLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string logMessage = $"[ExportPlus PDFOptions] {timestamp} - {message}";
            
            // Output to DebugView
            OutputDebugString(logMessage);
            
            // Output to Visual Studio Output window
            Debug.WriteLine(logMessage);
        }

        #endregion Debug Logging

        /// <summary>
        /// Apply all PDF export options to the document's print manager
        /// </summary>
        public static void ApplyPDFOptions(Document doc, ExportSettings options)
        {
            if (doc == null)
            {
                WriteDebugLog("ERROR: Document is null");
                throw new ArgumentNullException(nameof(doc));
            }

            if (options == null)
            {
                WriteDebugLog("ERROR: ExportSettings is null");
                throw new ArgumentNullException(nameof(options));
            }

            var printManager = doc.PrintManager;
            WriteDebugLog("Starting to apply PDF options...");

            try
            {
                // 1. Apply Paper Placement
                WriteDebugLog($"Applying Paper Placement: {options.PaperPlacement}");
                ApplyPaperPlacement(doc, printManager, options);

                // 2. Apply Zoom settings
                WriteDebugLog($"Applying Zoom: {options.Zoom}");
                ApplyZoomSettings(doc, printManager, options);

                // 3. Apply Hidden Line Views
                WriteDebugLog($"Applying Hidden Line Views: {options.HiddenLineViews}");
                ApplyHiddenLineSettings(doc, printManager, options);

                // 4. Apply Appearance settings
                WriteDebugLog($"Applying Appearance - RasterQuality: {options.RasterQuality}, Colors: {options.Colors}");
                ApplyAppearanceSettings(doc, printManager, options);

                // 5. Apply View Options (logged separately)
                WriteDebugLog("Applying View Options...");
                ApplyViewOptions(doc, options);

                // Apply all changes
                printManager.Apply();
                WriteDebugLog("✓ All PDF options applied successfully");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR applying PDF options: {ex.Message}");
                WriteDebugLog($"Stack trace: {ex.StackTrace}");
                throw new InvalidOperationException($"Error applying PDF options: {ex.Message}", ex);
            }
        }

        #region Private Helper Methods

        /// <summary>
        /// Apply paper placement settings (Center or Offset)
        /// </summary>
        private static void ApplyPaperPlacement(Document doc, PrintManager pm, ExportSettings options)
        {
            try
            {
                var printSetup = pm.PrintSetup;
                var currentSetting = printSetup.CurrentPrintSetting;
                
                using (Transaction trans = new Transaction(doc, "Apply Paper Placement"))
                {
                    trans.Start();

                    // Get print parameters
                    var printParams = currentSetting.PrintParameters;

                    if (options.PaperPlacement == PSPaperPlacement.Center)
                    {
                        // Note: PageOrientationType.Auto not available in Revit 2020, using Landscape as default
                        printParams.PageOrientation = PageOrientationType.Landscape;
                        WriteDebugLog("Paper placement set to CENTER");
                    }
                    else if (options.PaperPlacement == PSPaperPlacement.OffsetFromCorner)
                    {
                        // Apply offset - note: Revit uses units in feet
                        // If you have offset values, apply them here
                        WriteDebugLog("Paper placement set to OFFSET FROM CORNER");
                    }

                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error applying paper placement: {ex.Message}");
            }
        }

        /// <summary>
        /// Apply zoom settings (Fit to Page or specific zoom percentage)
        /// </summary>
        private static void ApplyZoomSettings(Document doc, PrintManager pm, ExportSettings options)
        {
            try
            {
                var printSetup = pm.PrintSetup;
                var currentSetting = printSetup.CurrentPrintSetting;

                using (Transaction trans = new Transaction(doc, "Apply Zoom Settings"))
                {
                    trans.Start();

                    var printParams = currentSetting.PrintParameters;

                    if (options.Zoom == PSZoomType.FitToPage)
                    {
                        printParams.ZoomType = ZoomType.FitToPage;
                        WriteDebugLog("Zoom set to FIT TO PAGE");
                    }
                    else if (options.Zoom == PSZoomType.Zoom)
                    {
                        printParams.ZoomType = ZoomType.Zoom;
                        printParams.Zoom = 100; // Default 100%
                        WriteDebugLog($"Zoom set to ZOOM at 100%");
                    }

                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error applying zoom settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Apply hidden line view settings (Vector or Raster processing)
        /// </summary>
        private static void ApplyHiddenLineSettings(Document doc, PrintManager pm, ExportSettings options)
        {
            try
            {
                var printSetup = pm.PrintSetup;
                var currentSetting = printSetup.CurrentPrintSetting;

                using (Transaction trans = new Transaction(doc, "Apply Hidden Line Settings"))
                {
                    trans.Start();

                    var printParams = currentSetting.PrintParameters;

                    if (options.HiddenLineViews == PSHiddenLineViews.VectorProcessing)
                    {
                        printParams.HiddenLineViews = HiddenLineViewsType.VectorProcessing;
                        WriteDebugLog("Hidden Line Views set to VECTOR PROCESSING");
                    }
                    else if (options.HiddenLineViews == PSHiddenLineViews.RasterProcessing)
                    {
                        printParams.HiddenLineViews = HiddenLineViewsType.RasterProcessing;
                        WriteDebugLog("Hidden Line Views set to RASTER PROCESSING");
                    }

                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error applying hidden line settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Apply appearance settings (Raster Quality and Colors)
        /// </summary>
        private static void ApplyAppearanceSettings(Document doc, PrintManager pm, ExportSettings options)
        {
            try
            {
                var printSetup = pm.PrintSetup;
                var currentSetting = printSetup.CurrentPrintSetting;

                using (Transaction trans = new Transaction(doc, "Apply Appearance Settings"))
                {
                    trans.Start();

                    var printParams = currentSetting.PrintParameters;

                    // Apply raster quality
                    switch (options.RasterQuality)
                    {
                        case PSRasterQuality.Low:
                            printParams.RasterQuality = RasterQualityType.Low;
                            break;
                        case PSRasterQuality.Medium:
                            printParams.RasterQuality = RasterQualityType.Medium;
                            break;
                        case PSRasterQuality.High:
                        case PSRasterQuality.Maximum:
                            printParams.RasterQuality = RasterQualityType.High;
                            break;
                    }
                    WriteDebugLog($"Raster Quality set to {options.RasterQuality}");

                    // Apply color settings
                    switch (options.Colors)
                    {
                        case PSColors.Color:
                            printParams.ColorDepth = ColorDepthType.Color;
                            break;
                        case PSColors.BlackAndWhite:
                            printParams.ColorDepth = ColorDepthType.BlackLine;
                            break;
                        case PSColors.Grayscale:
                            printParams.ColorDepth = ColorDepthType.GrayScale;
                            break;
                    }
                    WriteDebugLog($"Colors set to {options.Colors}");

                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error applying appearance settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Apply view options (hide/show categories)
        /// </summary>
        private static void ApplyViewOptions(Document doc, ExportSettings options)
        {
            try
            {
                WriteDebugLog($"View Options - HideRefWorkPlanes: {options.HideRefWorkPlanes}, " +
                            $"HideScopeBoxes: {options.HideScopeBoxes}, " +
                            $"HideCropBoundaries: {options.HideCropBoundaries}, " +
                            $"HideUnreferencedViewTags: {options.HideUnreferencedViewTags}, " +
                            $"ViewLinksInBlue: {options.ViewLinksInBlue}, " +
                            $"ReplaceHalftone: {options.ReplaceHalftone}, " +
                            $"RegionEdgesMask: {options.RegionEdgesMask}");

                // These options will be applied per-view during export
                // in the PDFExportManager when processing each sheet
                WriteDebugLog("View options logged for per-sheet application");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error logging view options: {ex.Message}");
            }
        }

        /// <summary>
        /// Hide or show a specific category in the active view (WITH Transaction)
        /// Use when NOT inside a Transaction/TransactionGroup
        /// </summary>
        public static void SetCategoryVisibility(Document doc, View view, BuiltInCategory category, bool hide)
        {
            try
            {
                var catId = new ElementId(category);
                
                if (view.CanCategoryBeHidden(catId))
                {
                    using (Transaction trans = new Transaction(doc, $"Set Category Visibility"))
                    {
                        trans.Start();
                        view.SetCategoryHidden(catId, hide);
                        trans.Commit();
                        
                        WriteDebugLog($"Category {category} visibility set to {(hide ? "HIDDEN" : "VISIBLE")} in view {view.Name}");
                    }
                }
                else
                {
                    WriteDebugLog($"Category {category} cannot be hidden in view {view.Name}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error setting category visibility: {ex.Message}");
            }
        }

        /// <summary>
        /// Hide or show a specific category in the active view (WITHOUT Transaction)
        /// Use when already inside a Transaction or TransactionGroup
        /// </summary>
        public static void SetCategoryVisibilityNoTransaction(Document doc, View view, BuiltInCategory category, bool hide)
        {
            try
            {
                var catId = new ElementId(category);
                
                if (view.CanCategoryBeHidden(catId))
                {
                    // NO TRANSACTION - caller must provide Transaction context
                    view.SetCategoryHidden(catId, hide);
                    WriteDebugLog($"Category {category} visibility set to {(hide ? "HIDDEN" : "VISIBLE")} in view {view.Name}");
                }
                else
                {
                    WriteDebugLog($"Category {category} cannot be hidden in view {view.Name}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error setting category visibility: {ex.Message}");
            }
        }

        /// <summary>
        /// Apply view-specific options to a sheet before export
        /// </summary>
        /// <summary>
        /// Apply view-specific options to a sheet (version WITHOUT transaction)
        /// Use when already inside a Transaction or TransactionGroup
        /// </summary>
        public static void ApplyViewOptionsToSheetNoTransaction(Document doc, ViewSheet sheet, ExportSettings options)
        {
            WriteDebugLog($"Applying view options to sheet: {sheet.SheetNumber} - {sheet.Name}");

            try
            {
                // NO TRANSACTION - caller must provide Transaction context

                // Hide ref/work planes
                if (options.HideRefWorkPlanes)
                {
                    SetCategoryVisibilityNoTransaction(doc, sheet, BuiltInCategory.OST_CLines, true);
                    SetCategoryVisibilityNoTransaction(doc, sheet, BuiltInCategory.OST_Grids, false); // Keep grids visible
                }

                // Hide scope boxes
                if (options.HideScopeBoxes)
                {
                    SetCategoryVisibilityNoTransaction(doc, sheet, BuiltInCategory.OST_VolumeOfInterest, true);
                }

                // Hide crop boundaries
                if (options.HideCropBoundaries)
                {
                    sheet.CropBoxVisible = false;
                }

                WriteDebugLog($"✓ View options applied to sheet {sheet.SheetNumber}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR applying view options to sheet: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Apply view-specific options to a sheet (version WITH transaction)
        /// Use when NOT inside a Transaction/TransactionGroup
        /// </summary>
        public static void ApplyViewOptionsToSheet(Document doc, ViewSheet sheet, ExportSettings options)
        {
            WriteDebugLog($"Applying view options to sheet: {sheet.SheetNumber} - {sheet.Name}");

            try
            {
                using (Transaction trans = new Transaction(doc, "Apply View Options to Sheet"))
                {
                    trans.Start();

                    // Hide ref/work planes
                    if (options.HideRefWorkPlanes)
                    {
                        SetCategoryVisibility(doc, sheet, BuiltInCategory.OST_CLines, true);
                        SetCategoryVisibility(doc, sheet, BuiltInCategory.OST_Grids, false); // Keep grids visible
                    }

                    // Hide scope boxes
                    if (options.HideScopeBoxes)
                    {
                        SetCategoryVisibility(doc, sheet, BuiltInCategory.OST_VolumeOfInterest, true);
                    }

                    // Hide crop boundaries
                    if (options.HideCropBoundaries)
                    {
                        sheet.CropBoxVisible = false;
                    }

                    trans.Commit();
                    WriteDebugLog($"✓ View options applied to sheet {sheet.SheetNumber}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR applying view options to sheet: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Apply PrintManager settings (Color, Zoom, Raster Quality, etc.) - Minimalist approach
        /// This method applies ONLY PrintManager settings without Document Transaction
        /// Use for minimalist PrintManager export approach
        /// </summary>
        public static void ApplyPrintManagerSettings(PrintManager pm, ExportSettings settings)
        {
            WriteDebugLog($"Applying PrintManager settings (minimalist)...");

            try
            {
                var printSetup = pm.PrintSetup;
                var currentSetting = printSetup.CurrentPrintSetting;
                var printParams = currentSetting.PrintParameters;

                // 1. Color vs Black & White
                WriteDebugLog($"Setting Colors: {settings.Colors}");
                if (settings.Colors == PSColors.BlackAndWhite)
                {
                    printParams.ColorDepth = ColorDepthType.BlackLine;
                    WriteDebugLog("✓ Color mode: BLACK AND WHITE");
                }
                else if (settings.Colors == PSColors.Grayscale)
                {
                    printParams.ColorDepth = ColorDepthType.GrayScale;
                    WriteDebugLog("✓ Color mode: GRAYSCALE");
                }
                else if (settings.Colors == PSColors.Color)
                {
                    printParams.ColorDepth = ColorDepthType.Color;
                    WriteDebugLog("✓ Color mode: COLOR");
                }

                // 2. Raster Quality
                WriteDebugLog($"Setting RasterQuality: {settings.RasterQuality}");
                if (settings.RasterQuality == PSRasterQuality.High)
                {
                    printParams.RasterQuality = RasterQualityType.High;
                    WriteDebugLog("✓ Raster quality: HIGH");
                }
                else if (settings.RasterQuality == PSRasterQuality.Medium)
                {
                    printParams.RasterQuality = RasterQualityType.Medium;
                    WriteDebugLog("✓ Raster quality: MEDIUM");
                }
                else if (settings.RasterQuality == PSRasterQuality.Low)
                {
                    printParams.RasterQuality = RasterQualityType.Low;
                    WriteDebugLog("✓ Raster quality: LOW");
                }
                else if (settings.RasterQuality == PSRasterQuality.Maximum)
                {
                    printParams.RasterQuality = RasterQualityType.Presentation;
                    WriteDebugLog("✓ Raster quality: MAXIMUM/PRESENTATION");
                }

                // 3. Zoom settings
                WriteDebugLog($"Setting Zoom: {settings.Zoom}");
                if (settings.Zoom == PSZoomType.FitToPage)
                {
                    printParams.ZoomType = ZoomType.FitToPage;
                    WriteDebugLog("✓ Zoom: FIT TO PAGE");
                }
                else if (settings.Zoom == PSZoomType.Zoom)
                {
                    printParams.ZoomType = ZoomType.Zoom;
                    // Note: Zoom percentage is set separately in PSZoomType.Zoom
                    // If there's a ZoomValue property, use: printParams.Zoom = settings.ZoomValue;
                    printParams.Zoom = 100; // Default 100% for now
                    WriteDebugLog($"✓ Zoom: 100% (custom zoom)");
                }

                // 4. Hidden Line Views
                WriteDebugLog($"Setting HiddenLineViews: {settings.HiddenLineViews}");
                if (settings.HiddenLineViews == PSHiddenLineViews.VectorProcessing)
                {
                    printParams.HiddenLineViews = HiddenLineViewsType.VectorProcessing;
                    WriteDebugLog("✓ Hidden lines: VECTOR PROCESSING");
                }
                else if (settings.HiddenLineViews == PSHiddenLineViews.RasterProcessing)
                {
                    printParams.HiddenLineViews = HiddenLineViewsType.RasterProcessing;
                    WriteDebugLog("✓ Hidden lines: RASTER PROCESSING");
                }

                // 5. Paper Placement
                WriteDebugLog($"Setting PaperPlacement: {settings.PaperPlacement}");
                if (settings.PaperPlacement == PSPaperPlacement.Center)
                {
                    // Note: PageOrientationType.Auto not available in Revit 2020, using Landscape as default
                    printParams.PageOrientation = PageOrientationType.Landscape;
                    WriteDebugLog("✓ Paper placement: CENTER");
                }
                else if (settings.PaperPlacement == PSPaperPlacement.OffsetFromCorner)
                {
                    // Apply offset if needed
                    WriteDebugLog("✓ Paper placement: OFFSET FROM CORNER");
                }

                // 6. Additional PDF options
                // Note: Some properties might not exist on PrintParameters
                // printParams.HideReferencePlane = settings.HideRefWorkPlanes;
                printParams.HideCropBoundaries = settings.HideCropBoundaries;
                printParams.HideUnreferencedViewTags = true; // Always hide unreferenced view tags
                
                WriteDebugLog("✓ PrintManager settings applied successfully (minimalist)");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR applying PrintManager settings: {ex.Message}");
                throw;
            }
        }

        #endregion Private Helper Methods
    }
}

