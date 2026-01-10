using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Quoc_MEP.Export.Models;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.IFC;
using Autodesk.Revit.DB.ExtensibleStorage;
using Newtonsoft.Json;
using ricaun.Revit.UI.StatusBar; // ✅ ADD: StatusBar package for progress display

namespace Quoc_MEP.Export.Managers
{
    /// <summary>
    /// IFC Export Manager - Compatible with Revit API
    /// </summary>
    public class IFCExportManager
    {
        private Document _document;

        public IFCExportManager(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// Export sheets/views to IFC format
        /// </summary>
        public bool ExportToIFC(List<ViewSheet> sheets, IFCExportSettings settings, string outputPath, Action<string> logCallback = null)
        {
            try
            {
                logCallback?.Invoke($"Starting IFC export with {sheets.Count} sheets");
                
                // Create IFC export options from settings
                var ifcOptions = CreateIFCExportOptions(settings, logCallback);
                
                // Create output directory if needed
                if (!Directory.Exists(outputPath))
                {
                    Directory.CreateDirectory(outputPath);
                    logCallback?.Invoke($"Created output directory: {outputPath}");
                }

                using (Transaction trans = new Transaction(_document, "Export IFC"))
                {
                    trans.Start();

                    try
                    {
                        foreach (var sheet in sheets)
                        {
                            // Note: ViewSheet from Revit API doesn't have CustomFileName
                            // CustomFileName is a property of SheetItem model
                            string fileName = SanitizeFileName(sheet.SheetNumber + "_" + sheet.Name);
                            
                            logCallback?.Invoke($"Exporting sheet: {sheet.SheetNumber} - {sheet.Name}");

                            // Export using Document.Export method for IFC
                            string fullPath = Path.Combine(outputPath, fileName + ".ifc");
                            
                            // Use IFC export with options
                            using (Transaction t = new Transaction(_document, "IFC Export"))
                            {
                                t.Start();
                                
                                // Note: Revit IFC export API is complex
                                // We'll use a simplified approach
                                bool success = ExportSingleSheet(sheet, fullPath, ifcOptions, logCallback);
                                
                                if (success)
                                {
                                    logCallback?.Invoke($"✓ Exported: {fileName}.ifc");
                                }
                                else
                                {
                                    logCallback?.Invoke($"✗ Failed to export: {fileName}");
                                }
                                
                                t.RollBack(); // Don't commit changes
                            }
                        }

                        trans.RollBack(); // Don't commit outer transaction
                        logCallback?.Invoke($"IFC export completed: {sheets.Count} sheets processed");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        logCallback?.Invoke($"ERROR during export: {ex.Message}");
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                logCallback?.Invoke($"IFC Export failed: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"IFC Export Error: {ex}");
                return false;
            }
        }

        /// <summary>
        /// Export single sheet
        /// </summary>
        private bool ExportSingleSheet(ViewSheet sheet, string filePath, IFCExportOptions options, Action<string> logCallback)
        {
            try
            {
                // Create a set of views to export
                var viewIds = new List<ElementId> { sheet.Id };
                
                // Export using the IFC exporter
                // Note: This is a simplified implementation
                // Full implementation would use IFCExportConfigurationsMap
                
                _document.Export(Path.GetDirectoryName(filePath), 
                                Path.GetFileNameWithoutExtension(filePath), 
                                options);
                
                return File.Exists(filePath);
            }
            catch (Exception ex)
            {
                logCallback?.Invoke($"Error exporting sheet: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Create IFCExportOptions from settings (Enhanced with APS features)
        /// </summary>
        private IFCExportOptions CreateIFCExportOptions(IFCExportSettings settings, Action<string> logCallback = null)
        {
            var options = new IFCExportOptions();

            try
            {
                logCallback?.Invoke("=== IFC Export Configuration (APS Enhanced) ===");
                
                // IFC Version (CORE - Always available)
                options.FileVersion = ConvertIFCVersion(settings.IFCVersion);
                logCallback?.Invoke($"✓ IFC Version: {settings.IFCVersion}");

                // Space Boundaries (CORE - Always available)
                options.SpaceBoundaryLevel = ConvertSpaceBoundaries(settings.SpaceBoundaries);
                logCallback?.Invoke($"✓ Space Boundaries: {settings.SpaceBoundaries} (Level {options.SpaceBoundaryLevel})");

                // Property Sets - ExportBaseQuantities (Always available)
                options.ExportBaseQuantities = settings.ExportBaseQuantities;
                logCallback?.Invoke($"✓ Export Base Quantities: {settings.ExportBaseQuantities}");
                
                // ========== USER DEFINED PROPERTY SETS (APS FEATURE) ==========
                // This is critical for custom IFC properties like COBie, etc.
                if (settings.ExportUserDefinedPsets && !string.IsNullOrEmpty(settings.ExportUserDefinedPsetsFileName))
                {
                    string psetsFile = settings.ExportUserDefinedPsetsFileName;
                    
                    // Fix relative paths (like APS does)
                    if (!Path.IsPathRooted(psetsFile))
                    {
                        psetsFile = Path.Combine(Path.GetDirectoryName(_document.PathName) ?? "", psetsFile);
                    }
                    
                    if (File.Exists(psetsFile))
                    {
                        try
                        {
                            // Apply property sets file
                            options.AddOption("ExportUserDefinedPsets", "true");
                            options.AddOption("ExportUserDefinedPsetsFileName", psetsFile);
                            
                            logCallback?.Invoke($"✓ User Defined Property Sets: ENABLED");
                            logCallback?.Invoke($"  └─ File: {Path.GetFileName(psetsFile)}");
                        }
                        catch (Exception ex)
                        {
                            logCallback?.Invoke($"⚠ Warning: Could not set property sets file: {ex.Message}");
                        }
                    }
                    else
                    {
                        logCallback?.Invoke($"⚠ Warning: Property sets file not found: {psetsFile}");
                    }
                }
                
                // ========== PARAMETER MAPPING (APS FEATURE) ==========
                // Maps Revit parameters to IFC properties
                if (settings.ExportParameterMapping && !string.IsNullOrEmpty(settings.ExportParameterMappingFileName))
                {
                    string mappingFile = settings.ExportParameterMappingFileName;
                    
                    // Fix relative paths
                    if (!Path.IsPathRooted(mappingFile))
                    {
                        mappingFile = Path.Combine(Path.GetDirectoryName(_document.PathName) ?? "", mappingFile);
                    }
                    
                    if (File.Exists(mappingFile))
                    {
                        try
                        {
                            options.AddOption("ExportUserDefinedParameterMapping", "true");
                            options.AddOption("ExportUserDefinedParameterMappingFileName", mappingFile);
                            
                            logCallback?.Invoke($"✓ Parameter Mapping: ENABLED");
                            logCallback?.Invoke($"  └─ File: {Path.GetFileName(mappingFile)}");
                        }
                        catch (Exception ex)
                        {
                            logCallback?.Invoke($"⚠ Warning: Could not set parameter mapping: {ex.Message}");
                        }
                    }
                    else
                    {
                        logCallback?.Invoke($"⚠ Warning: Parameter mapping file not found: {mappingFile}");
                    }
                }
                
                // ========== ADVANCED OPTIONS (APS FEATURES) ==========
                
                // Export only elements visible in active view
                if (settings.VisibleElementsOfCurrentView)
                {
                    try
                    {
                        options.AddOption("VisibleElementsOfCurrentView", "true");
                        logCallback?.Invoke($"✓ Visible Elements Only: ENABLED");
                    }
                    catch { }
                }
                
                // Use active view geometry
                if (settings.UseActiveViewGeometry)
                {
                    try
                    {
                        options.AddOption("UseActiveViewGeometry", "true");
                        logCallback?.Invoke($"✓ Use Active View Geometry: ENABLED");
                    }
                    catch { }
                }
                
                // Wall and Column Splitting
                try 
                { 
                    options.WallAndColumnSplitting = settings.SplitWallsByLevel; 
                    logCallback?.Invoke($"✓ Split Walls/Columns by Level: {settings.SplitWallsByLevel}");
                } 
                catch { }
                
                // Export linked files
                try
                {
                    options.AddOption("ExportLinkedFiles", settings.ExportLinkedFiles.ToString());
                    logCallback?.Invoke($"✓ Export Linked Files: {settings.ExportLinkedFiles}");
                }
                catch { }
                
                // Store IFC GUID
                try
                {
                    options.AddOption("StoreIFCGUID", settings.StoreIFCGUID.ToString());
                    if (settings.StoreIFCGUID)
                    {
                        logCallback?.Invoke($"✓ Store IFC GUID: ENABLED (GUIDs will be saved to Revit model)");
                    }
                }
                catch { }

                logCallback?.Invoke("=== IFC Export Options Applied Successfully ===");
            }
            catch (Exception ex)
            {
                logCallback?.Invoke($"⚠ Warning: Some IFC options not set: {ex.Message}");
            }

            return options;
        }

        /// <summary>
        /// Convert IFC Version string to enum
        /// </summary>
        private Autodesk.Revit.DB.IFCVersion ConvertIFCVersion(string version)
        {
            switch (version)
            {
                case "IFC 2x3 Coordination View 2.0":
                case "IFC 2x3 Coordination View":
                    return Autodesk.Revit.DB.IFCVersion.IFC2x3CV2;
                    
                case "IFC 4 Reference View":
                    return Autodesk.Revit.DB.IFCVersion.IFC4RV;
                    
                case "IFC 4 Design Transfer View":
                    return Autodesk.Revit.DB.IFCVersion.IFC4DTV;
                    
                case "IFC 2x2":
                    return Autodesk.Revit.DB.IFCVersion.IFC2x2;
                    
                case "IFC 4":
                    return Autodesk.Revit.DB.IFCVersion.IFC4;
                    
                default:
                    return Autodesk.Revit.DB.IFCVersion.IFC2x3CV2; // Default
            }
        }

        /// <summary>
        /// Convert Space Boundaries string to level
        /// </summary>
        private int ConvertSpaceBoundaries(string spaceBoundaries)
        {
            switch (spaceBoundaries)
            {
                case "None":
                    return 0;
                case "1st Level":
                    return 1;
                case "2nd Level":
                    return 2;
                default:
                    return 0;
            }
        }

        /// <summary>
        /// Sanitize file name to remove invalid characters
        /// </summary>
        private string SanitizeFileName(string fileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                fileName = fileName.Replace(c, '_');
            }
            
            // Also replace some additional problematic characters
            fileName = fileName.Replace(':', '-');
            fileName = fileName.Replace('/', '-');
            fileName = fileName.Replace('\\', '-');
            
            return fileName;
        }

        /// <summary>
        /// Get all 3D views from document (for View-based export)
        /// </summary>
        public List<View3D> Get3DViews()
        {
            var views3D = new FilteredElementCollector(_document)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate)
                .ToList();

            return views3D;
        }

        /// <summary>
        /// Export 3D views to IFC
        /// </summary>
        public bool Export3DViewsToIFC(List<View3D> views, IFCExportSettings settings, string outputPath, Action<string> logCallback = null, Action<string, bool> progressCallback = null)
        {
            try
            {
                logCallback?.Invoke($"Starting IFC export with {views.Count} 3D views");
                
                var ifcOptions = CreateIFCExportOptions(settings, logCallback);
                
                if (!Directory.Exists(outputPath))
                {
                    Directory.CreateDirectory(outputPath);
                }

                int successCount = 0;
                int failCount = 0;

                // CRITICAL: Wrap entire export in Transaction
                // This ONLY works when called from ExternalEvent/ExternalCommand (Revit API context)
                using (Transaction trans = new Transaction(_document, "Export IFC"))
                {
                    trans.Start();
                    
                    try
                    {
                        // ✅ ADD: StatusBar progress integration (giống StatusBar Demo)
                        RevitProgressBarUtils.Run($"Exporting {views.Count} views to IFC", views, (view) =>
                        {
                            string fileName = SanitizeFileName(view.Name);
                            string fullPath = Path.Combine(outputPath, fileName + ".ifc");
                            
                            logCallback?.Invoke($"Exporting 3D view: {view.Name} (ID: {view.Id.IntegerValue})");

                            try
                            {
                                // Create IFC options for this view
                                var viewSpecificOptions = CreateIFCExportOptions(settings, null); // Don't repeat logs
                                
                                // Set FilterViewId to limit export to this specific 3D view
                                viewSpecificOptions.FilterViewId = view.Id;
                                logCallback?.Invoke($"  Set FilterViewId: {view.Id.IntegerValue}");

                                // Export to IFC
                                _document.Export(Path.GetDirectoryName(fullPath), 
                                                Path.GetFileNameWithoutExtension(fullPath), 
                                                viewSpecificOptions);
                                
                                // Verify file was created
                                if (File.Exists(fullPath))
                                {
                                    var fileInfo = new FileInfo(fullPath);
                                    logCallback?.Invoke($"✓ Exported: {fileName}.ifc ({fileInfo.Length / 1024} KB)");
                                    successCount++;
                                    
                                    // Notify progress callback - this file completed successfully
                                    progressCallback?.Invoke(view.Name, true);
                                }
                                else
                                {
                                    logCallback?.Invoke($"✗ Export failed: File not created for {view.Name}");
                                    failCount++;
                                    
                                    // Notify progress callback - this file failed
                                    progressCallback?.Invoke(view.Name, false);
                                }
                            }
                            catch (Exception ex)
                            {
                                logCallback?.Invoke($"✗ Failed to export {view.Name}: {ex.Message}");
                                logCallback?.Invoke($"   Exception Type: {ex.GetType().Name}");
                                if (ex.InnerException != null)
                                {
                                    logCallback?.Invoke($"   Inner Exception: {ex.InnerException.Message}");
                                }
                                failCount++;
                                
                                // Notify progress callback - this file failed
                                progressCallback?.Invoke(view.Name, false);
                            }
                        }); // ✅ Close RevitProgressBarUtils.Run()
                        
                        // Commit transaction after all exports
                        trans.Commit();
                        logCallback?.Invoke($"Transaction committed successfully");
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        logCallback?.Invoke($"Transaction rolled back due to error: {ex.Message}");
                        throw;
                    }
                }

                logCallback?.Invoke($"IFC export completed: {successCount} succeeded, {failCount} failed");
                return successCount > 0; // Return true only if at least one view exported successfully
            }
            catch (Exception ex)
            {
                logCallback?.Invoke($"IFC Export failed: {ex.Message}");
                logCallback?.Invoke($"Stack Trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Get list of IFC Export Configuration names from Revit
        /// This includes both built-in setups and user-created custom setups
        /// READS FROM: Document's ExtensibleStorage (same as Autodesk IFC Exporter)
        /// </summary>
        public static List<string> GetAvailableIFCSetups(Document document)
        {
            var setupNames = new List<string>();

            // DEBUG: Write to file for verification
            string debugLogPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), 
                "ExportPlus_IFC_Debug.txt");
            System.IO.File.AppendAllText(debugLogPath, $"\n\n========== {DateTime.Now} ==========\n");
            System.IO.File.AppendAllText(debugLogPath, "GetAvailableIFCSetups() CALLED - Using ExtensibleStorage\n");

            try
            {
                // Always add In-Session Setup first (default Revit behavior)
                setupNames.Add("<In-Session Setup>");
                System.IO.File.AppendAllText(debugLogPath, "Added: <In-Session Setup>\n");

                // METHOD: Read from Document's ExtensibleStorage (like Autodesk IFC Exporter does)
                // Schema GUID used by Autodesk IFC Exporter: DCB88B13-594F-44F6-8F5D-AE9477305AC3
                System.IO.File.AppendAllText(debugLogPath, "========== READING FROM EXTENSIBLE STORAGE ==========\n");
                
                try
                {
                    // Try to find IFC Configuration schema in document
                    Guid jsonSchemaId = new Guid("C2A3E6FE-CE51-4F35-8FF1-20C34567B687"); // Latest JSON-based schema
                    Guid oldSchemaId = new Guid("DCB88B13-594F-44F6-8F5D-AE9477305AC3");  // Older MapField schema
                    
                    Schema jsonSchema = Schema.Lookup(jsonSchemaId);
                    Schema oldSchema = Schema.Lookup(oldSchemaId);
                    
                    System.IO.File.AppendAllText(debugLogPath, $"JSON Schema found: {jsonSchema != null}\n");
                    System.IO.File.AppendAllText(debugLogPath, $"Old Schema found: {oldSchema != null}\n");
                    
                    int customCount = 0;
                    
                    // Try JSON schema first (Revit 2020+)
                    if (jsonSchema != null)
                    {
                        System.IO.File.AppendAllText(debugLogPath, "Using JSON Schema...\n");
                        
                        // Get all DataStorage elements with this schema
                        FilteredElementCollector collector = new FilteredElementCollector(document);
                        var dataStorages = collector.OfClass(typeof(DataStorage)).Cast<DataStorage>();
                        
                        foreach (DataStorage storage in dataStorages)
                        {
                            Entity entity = storage.GetEntity(jsonSchema);
                            if (entity != null && entity.IsValid())
                            {
                                try
                                {
                                    // Get configuration data from JSON field
                                    string configData = entity.Get<string>("MapField");
                                    if (!string.IsNullOrEmpty(configData))
                                    {
                                        // Parse JSON to get configuration name
                                        var configDict = JsonConvert.DeserializeObject<Dictionary<string, object>>(configData);
                                        if (configDict != null && configDict.ContainsKey("Name"))
                                        {
                                            string configName = configDict["Name"].ToString();
                                            if (!setupNames.Contains(configName))
                                            {
                                                setupNames.Add(configName);
                                                customCount++;
                                                System.IO.File.AppendAllText(debugLogPath, $"  [{customCount}] {configName} (from JSON)\n");
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.IO.File.AppendAllText(debugLogPath, $"  Error parsing JSON config: {ex.Message}\n");
                                }
                            }
                        }
                    }
                    
                    // Try old schema (Revit 2019 and earlier)
                    if (oldSchema != null && customCount == 0)
                    {
                        System.IO.File.AppendAllText(debugLogPath, "Using Old MapField Schema...\n");
                        
                        FilteredElementCollector collector = new FilteredElementCollector(document);
                        var dataStorages = collector.OfClass(typeof(DataStorage)).Cast<DataStorage>();
                        
                        foreach (DataStorage storage in dataStorages)
                        {
                            Entity entity = storage.GetEntity(oldSchema);
                            if (entity != null && entity.IsValid())
                            {
                                try
                                {
                                    // Get configuration map (Dictionary)
                                    var configMap = entity.Get<IDictionary<string, string>>("MapField");
                                    if (configMap != null && configMap.ContainsKey("Name"))
                                    {
                                        string configName = configMap["Name"];
                                        if (!setupNames.Contains(configName))
                                        {
                                            setupNames.Add(configName);
                                            customCount++;
                                            System.IO.File.AppendAllText(debugLogPath, $"  [{customCount}] {configName} (from MapField)\n");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.IO.File.AppendAllText(debugLogPath, $"  Error reading MapField config: {ex.Message}\n");
                                }
                            }
                        }
                    }
                    
                    System.IO.File.AppendAllText(debugLogPath, $"✓ Found {customCount} custom configurations in document\n");
                }
                catch (Exception ex)
                {
                    System.IO.File.AppendAllText(debugLogPath, $"✗ ERROR reading ExtensibleStorage: {ex.Message}\n");
                    System.IO.File.AppendAllText(debugLogPath, $"   Type: {ex.GetType().Name}\n");
                    System.IO.File.AppendAllText(debugLogPath, $"   Stack: {ex.StackTrace}\n");
                }

                // ALWAYS add built-in configurations (even if custom ones found)
                System.IO.File.AppendAllText(debugLogPath, "========== ADDING BUILT-IN CONFIGURATIONS ==========\n");
                
                List<string> builtInSetups = new List<string>
                {
                    "IFC 2x3 Coordination View 2.0",
                    "IFC 2x3 Coordination View",
                    "IFC 2x3 GSA Concept Design BIM 2010",
                    "IFC 2x3 Basic FM Handover View",
                    "IFC 2x2 Coordination View",
                    "IFC 2x2 Singapore BCA e-Plan Check",
                    "IFC 2x3 COBie 2.4 Design Deliverable View",
                    "IFC4 Reference View",
                    "IFC4 Design Transfer View"
                };
                
                foreach (string builtIn in builtInSetups)
                {
                    if (!setupNames.Contains(builtIn))
                    {
                        setupNames.Add(builtIn);
                    }
                }
                
                System.IO.File.AppendAllText(debugLogPath, $"✓ Added {builtInSetups.Count} built-in setups\n");
                System.IO.File.AppendAllText(debugLogPath, $"\n========== FINAL TOTAL: {setupNames.Count} setups ==========\n");
                
                // Log all final setups
                for (int i = 0; i < setupNames.Count; i++)
                {
                    System.IO.File.AppendAllText(debugLogPath, $"  [{i+1}] {setupNames[i]}\n");
                }
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(debugLogPath, $"\n✗ OUTER ERROR: {ex.Message}\n");
                System.IO.File.AppendAllText(debugLogPath, $"   Stack: {ex.StackTrace}\n");
                
                // Ensure at least basic setups
                if (setupNames.Count == 0)
                {
                    setupNames.Add("<In-Session Setup>");
                }
            }

            return setupNames;
        }

        /// <summary>
        /// Load IFC Export Configuration from Revit by name
        /// Returns settings object populated with the configuration values
        /// </summary>
        public static IFCExportSettings LoadIFCSetupFromRevit(Document document, string setupName)
        {
            var settings = new IFCExportSettings();

            try
            {
                // If In-Session Setup, return default settings
                if (setupName == "<In-Session Setup>")
                {
                    return settings;
                }

                // Try to load configuration from Revit
                try
                {
                    var ifcExportConfigType = Type.GetType("BIM.IFC.Export.UI.IFCExportConfigurationsMap, RevitIFCUI");
                    
                    if (ifcExportConfigType != null)
                    {
                        var getMethod = ifcExportConfigType.GetMethod("GetStoredConfigurations",
                            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                        
                        if (getMethod != null)
                        {
                            var configs = getMethod.Invoke(null, new object[] { document }) as IDictionary<string, object>;
                            
                            if (configs != null && configs.ContainsKey(setupName))
                            {
                                var config = configs[setupName];
                                
                                // Use reflection to read configuration properties
                                var configType = config.GetType();
                                
                                // Try to read common properties
                                var ifcVersionProp = configType.GetProperty("IFCVersion");
                                if (ifcVersionProp != null)
                                {
                                    var versionValue = ifcVersionProp.GetValue(config);
                                    settings.IFCVersion = versionValue?.ToString() ?? "IFC 2x3 Coordination View 2.0";
                                }

                                var spaceBoundariesProp = configType.GetProperty("SpaceBoundaries");
                                if (spaceBoundariesProp != null)
                                {
                                    var spaceBoundValue = spaceBoundariesProp.GetValue(config);
                                    settings.SpaceBoundaries = spaceBoundValue?.ToString() ?? "None";
                                }

                                var exportBaseQtyProp = configType.GetProperty("ExportBaseQuantities");
                                if (exportBaseQtyProp != null)
                                {
                                    var baseQtyValue = exportBaseQtyProp.GetValue(config);
                                    settings.ExportBaseQuantities = baseQtyValue is bool ? (bool)baseQtyValue : false;
                                }

                                // Add more property mappings as needed...
                                
                                return settings;
                            }
                        }
                    }
                }
                catch
                {
                    // Failed to load from Revit API, use defaults
                }

                // If we couldn't load from Revit, create sensible defaults based on setup name
                settings = CreateDefaultSetupSettings(setupName);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading IFC setup '{setupName}': {ex.Message}");
            }

            return settings;
        }

        /// <summary>
        /// Create default settings for known setup types
        /// </summary>
        private static IFCExportSettings CreateDefaultSetupSettings(string setupName)
        {
            var settings = new IFCExportSettings();

            // Remove " Setup>" suffix if present for matching
            var cleanName = setupName.Replace(" Setup>", "").Replace("Setup>", "");

            // Match based on clean name or original name
            if (cleanName.Contains("IFC 2x3 Coordination View 2.0") || setupName.Contains("IFC 2x3 Coordination View 2.0"))
            {
                settings.IFCVersion = "IFC 2x3 Coordination View 2.0";
                settings.FileType = "IFC";
                settings.ExportBaseQuantities = false;
                settings.SplitWallsByLevel = true;
            }
            else if (cleanName.Contains("IFC 2x3 GSA") || setupName.Contains("IFC 2x3 GSA"))
            {
                settings.IFCVersion = "IFC 2x3 GSA Concept Design BIM 2010";
                settings.FileType = "IFC";
                settings.ExportBaseQuantities = true;
                settings.ExportBoundingBox = true;
            }
            else if (cleanName.Contains("IFC 2x3 Basic FM") || setupName.Contains("IFC 2x3 Basic FM"))
            {
                settings.IFCVersion = "IFC 2x3 Basic FM Handover View";
                settings.FileType = "IFC";
                settings.ExportBaseQuantities = true;
                settings.SpaceBoundaries = "1st Level";
                settings.ExportRoomsIn3DViews = true;
            }
            else if (cleanName.Contains("IFC 2x3 COBie") || setupName.Contains("IFC 2x3 COBie"))
            {
                settings.IFCVersion = "IFC 2x3 COBie 2.4 Design Deliverable View";
                settings.FileType = "IFC";
                settings.ExportBaseQuantities = true;
                settings.SpaceBoundaries = "2nd Level";
                settings.ExportRoomsIn3DViews = true;
            }
            else if (cleanName.Contains("IFC4 Reference View") || setupName.Contains("IFC4 Reference View"))
            {
                settings.IFCVersion = "IFC4 Reference View";
                settings.FileType = "IFC";
                settings.ExportBaseQuantities = false;
            }
            else if (cleanName.Contains("IFC4 Design Transfer") || setupName.Contains("IFC4 Design Transfer"))
            {
                settings.IFCVersion = "IFC4 Design Transfer View";
                settings.FileType = "IFC";
                settings.ExportBaseQuantities = true;
            }
            else if (cleanName.Contains("IFC 2x3 Coordination View") || setupName.Contains("IFC 2x3 Coordination View"))
            {
                settings.IFCVersion = "IFC 2x3 Coordination View";
                settings.FileType = "IFC";
                settings.ExportBaseQuantities = false;
            }
            else if (cleanName.Contains("IFC 2x2 Coordination View") || setupName.Contains("IFC 2x2 Coordination View"))
            {
                settings.IFCVersion = "IFC 2x2 Coordination View";
                settings.FileType = "IFC";
                settings.ExportBaseQuantities = false;
            }
            else if (cleanName.Contains("IFC 2x2 Singapore") || setupName.Contains("IFC 2x2 Singapore"))
            {
                settings.IFCVersion = "IFC 2x2 Singapore BCA e-Plan Check";
                settings.FileType = "IFC";
                settings.ExportBaseQuantities = true;
            }
            else if (cleanName.Contains("Typical") || setupName.Contains("Typical"))
            {
                settings.IFCVersion = "IFC 2x3 Coordination View 2.0";
                settings.FileType = "IFC";
                settings.DetailLevel = "Medium";
            }
            else
            {
                // Default fallback for unknown/custom setups
                settings.IFCVersion = "IFC 2x3 Coordination View 2.0";
                settings.FileType = "IFC";
            }

            return settings;
        }
    }
}