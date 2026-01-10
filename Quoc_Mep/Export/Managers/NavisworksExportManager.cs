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
    /// Manager for exporting to Navisworks format (NWC)
    /// </summary>
    public class NavisworksExportManager
    {
        private readonly Document _document;

        public NavisworksExportManager(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// Export 3D views to Navisworks Cache (NWC) format with custom settings
        /// </summary>
        public bool ExportToNavisworks(List<ViewItem> selectedViews, NWCExportSettings settings, string outputFolder, string fileNamePrefix = "", Action<string, bool> progressCallback = null)
        {
            try
            {
                if (selectedViews == null || !selectedViews.Any())
                {
                    System.Diagnostics.Debug.WriteLine("No views selected for Navisworks export");
                    return false;
                }

                if (string.IsNullOrEmpty(outputFolder))
                {
                    System.Diagnostics.Debug.WriteLine("Invalid output folder");
                    return false;
                }

                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Filter only 3D views for Navisworks export
                var threeDViews = selectedViews.Where(v => 
                    v.ViewType != null && 
                    (v.ViewType.Contains("ThreeD") || v.ViewType.Contains("3D"))).ToList();
                
                if (!threeDViews.Any())
                {
                    System.Diagnostics.Debug.WriteLine("No 3D views found. Creating default 3D export.");
                    return ExportModelToNavisworks(settings, outputFolder, fileNamePrefix);
                }

                int exportedCount = 0;

                // ✅ ADD: StatusBar progress integration (giống StatusBar Demo)
                RevitProgressBarUtils.Run($"Exporting {threeDViews.Count} views to NWC", threeDViews, (viewItem) =>
                {
                    try
                    {
                        var view = _document.GetElement(viewItem.RevitViewId) as View3D;
                        if (view != null)
                        {
                            string fileName = !string.IsNullOrEmpty(viewItem.CustomFileName) 
                                ? viewItem.CustomFileName 
                                : $"{fileNamePrefix}{view.Name}";
                            
                            // Clean filename
                            fileName = CleanFileName(fileName);
                            string fullPath = Path.Combine(outputFolder, $"{fileName}.nwc");

                            // Create export options with settings
                            var options = CreateNavisworksExportOptions(settings);
                            options.ExportScope = NavisworksExportScope.View;
                            options.ViewId = view.Id;

                            // Export to Navisworks
                            _document.Export(outputFolder, fileName, options);
                            exportedCount++;
                            
                            // Notify progress callback
                            progressCallback?.Invoke(viewItem.ViewName, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log error but continue with other views
                        System.Diagnostics.Debug.WriteLine($"Error exporting view {viewItem.ViewName}: {ex.Message}");
                        progressCallback?.Invoke(viewItem.ViewName, false);
                    }
                }); // ✅ Close RevitProgressBarUtils.Run()

                return exportedCount > 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Navisworks export error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Export entire model to Navisworks with settings
        /// </summary>
        private bool ExportModelToNavisworks(NWCExportSettings settings, string outputFolder, string fileNamePrefix)
        {
            try
            {
                string fileName = !string.IsNullOrEmpty(fileNamePrefix) 
                    ? $"{fileNamePrefix}_Model" 
                    : $"{_document.Title}_Model";
                
                fileName = CleanFileName(fileName);
                
                var options = CreateNavisworksExportOptions(settings);
                options.ExportScope = NavisworksExportScope.Model;

                _document.Export(outputFolder, fileName, options);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Model export to Navisworks error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Create NavisworksExportOptions from NWCExportSettings
        /// </summary>
        private NavisworksExportOptions CreateNavisworksExportOptions(NWCExportSettings settings)
        {
            var options = new NavisworksExportOptions();

            try
            {
                // Set basic properties that are available in NavisworksExportOptions
                options.ExportRoomGeometry = settings.ExportRoomGeometry;
                options.DivideFileIntoLevels = settings.DivideFileIntoLevels;
                options.ExportRoomAsAttribute = settings.ConvertRoomAsAttribute;
                options.ExportLinks = settings.ConvertLinkedFiles;
                options.ExportUrls = settings.ConvertURLs;
                options.FindMissingMaterials = settings.TryAndFindMissingMaterials;
                options.ConvertElementProperties = settings.ConvertElementProperties;

                // Convert element parameters enum (None, Elements, All)
                switch (settings.ConvertElementParameters)
                {
                    case "None":
                        options.Parameters = NavisworksParameters.None;
                        break;
                    case "Elements":
                        options.Parameters = NavisworksParameters.Elements;
                        break;
                    case "All":
                    default:
                        options.Parameters = NavisworksParameters.All;
                        break;
                }

                // Convert coordinates enum (Shared, Project Internal)
                switch (settings.Coordinates)
                {
                    case "Shared":
                        options.Coordinates = NavisworksCoordinates.Shared;
                        break;
                    case "Project Internal":
                    case "Internal":
                    default:
                        options.Coordinates = NavisworksCoordinates.Internal;
                        break;
                }

                // Apply Revit 2020+ settings using reflection (for version compatibility)
                try
                {
                    var type = options.GetType();
                    
                    // Convert lights (Revit 2020+)
                    var convertLightsProperty = type.GetProperty("ConvertLights");
                    if (convertLightsProperty != null && convertLightsProperty.CanWrite)
                        convertLightsProperty.SetValue(options, settings.ConvertLights);
                    
                    // Convert linked CAD formats (Revit 2020+)
                    var convertLinkedCADProperty = type.GetProperty("ConvertLinkedCADFormats");
                    if (convertLinkedCADProperty != null && convertLinkedCADProperty.CanWrite)
                        convertLinkedCADProperty.SetValue(options, settings.ConvertLinkedCADFormats);
                    
                    // Faceting factor (Revit 2020+)
                    var facetingFactorProperty = type.GetProperty("FacetingFactor");
                    if (facetingFactorProperty != null && facetingFactorProperty.CanWrite)
                        facetingFactorProperty.SetValue(options, settings.FacetingFactor);
                    
                    // Convert element IDs
                    var convertElementIdsProperty = type.GetProperty("ConvertElementId");
                    if (convertElementIdsProperty != null && convertElementIdsProperty.CanWrite)
                        convertElementIdsProperty.SetValue(options, settings.ConvertElementIds);
                }
                catch
                {
                    // Ignore if properties don't exist in this Revit version
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating Navisworks export options: {ex.Message}");
            }

            return options;
        }

        /// <summary>
        /// Export sheets as reference for coordination (limited functionality)
        /// </summary>
        public bool ExportSheetsReference(List<SheetItem> selectedSheets, string outputFolder, string fileNamePrefix = "")
        {
            try
            {
                // Navisworks typically works with 3D models, not 2D sheets
                // This creates a reference file with sheet information
                
                string fileName = !string.IsNullOrEmpty(fileNamePrefix) 
                    ? $"{fileNamePrefix}_Sheets_Reference" 
                    : "Sheets_Reference";
                
                fileName = CleanFileName(fileName);
                string fullPath = Path.Combine(outputFolder, $"{fileName}.txt");

                var content = new List<string>
                {
                    "Revit Sheets Reference for Navisworks",
                    $"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                    $"Document: {_document.Title}",
                    "",
                    "Selected Sheets:"
                };

                foreach (var sheet in selectedSheets)
                {
                    content.Add($"- {sheet.SheetNumber}: {sheet.SheetName}");
                }

                File.WriteAllLines(fullPath, content);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Sheets reference export error: {ex.Message}");
                return false;
            }
        }

        public bool ExportSheetsReference(List<SheetItem> sheets, string outputFolder)
        {
            try
            {
                if (sheets?.Any() != true)
                    return false;

                string fileName = "Sheets_Reference";
                string filePath = Path.Combine(outputFolder, $"{fileName}.nwc");

                // Create 3D view for reference
                var collector = new FilteredElementCollector(_document);
                var view3D = collector.OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate);

                if (view3D == null)
                {
                    return false;
                }

                var exportOptions = new NavisworksExportOptions();
                exportOptions.ExportScope = NavisworksExportScope.View;
                exportOptions.ViewId = view3D.Id;

                try
                {
                    _document.Export(outputFolder, fileName, exportOptions);
                    return true;
                }
                catch
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private string CleanFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return "Untitled";
            
            // Remove invalid characters
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');
            }
            
            // Replace spaces and special characters
            fileName = fileName.Replace(' ', '_')
                             .Replace('/', '_')
                             .Replace('\\', '_')
                             .Replace(':', '_')
                             .Replace('*', '_')
                             .Replace('?', '_')
                             .Replace('"', '_')
                             .Replace('<', '_')
                             .Replace('>', '_')
                             .Replace('|', '_');
            
            return fileName;
        }
    }
}