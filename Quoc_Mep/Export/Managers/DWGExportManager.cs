using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Quoc_MEP.Export.Models;
using Quoc_MEP.Export.Utils;
using RevitDB = Autodesk.Revit.DB;
using ricaun.Revit.UI.StatusBar; // ✅ ADD: StatusBar package for progress display

namespace Quoc_MEP.Export.Managers
{
    public class DWGExportManager
    {
        private readonly RevitDB.Document _document;
        
        public DWGExportManager(RevitDB.Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// Get DWG export options from existing setup or create default
        /// </summary>
        private RevitDB.DWGExportOptions GetDWGExportOptions(string setupName, PSDWGExportSettings settings)
        {
            try
            {
                // Try to find existing DWG export setup in document
                RevitDB.ExportDWGSettings dwgSettings = RevitDB.ExportDWGSettings.FindByName(_document, setupName);
                
                if (dwgSettings != null)
                {
                    System.Diagnostics.Debug.WriteLine($"Using existing DWG setup: {setupName}");
                    RevitDB.DWGExportOptions options = dwgSettings.GetDWGExportOptions();
                    
                    // CRITICAL: Override settings from UI
                    OverrideOptionsFromUI(options, settings);
                    
                    return options;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Setup '{setupName}' not found, creating default options");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading DWG setup: {ex.Message}");
            }
            
            // If setup not found or error, create default options
            return CreateDefaultDWGOptions(settings);
        }

        /// <summary>
        /// Override DWG options with UI settings
        /// </summary>
        private void OverrideOptionsFromUI(RevitDB.DWGExportOptions options, PSDWGExportSettings settings)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DWG Override] ExportViewsOnSheets: {settings.ExportViewsOnSheets}");
                
                // CRITICAL: Prevent multiple XREF files - similar to DIRoots approach
                if (!settings.ExportViewsOnSheets)
                {
                    System.Diagnostics.Debug.WriteLine("[DWG Override] DISABLING all XREF export options");
                    
                    // Key properties to disable XREF export
                    TrySetProperty(options, "ExportingAreas", false);  // Disable exporting views on sheets as separate files
                    TrySetProperty(options, "MergedViews", false);      // Don't merge views
                    TrySetProperty(options, "ExportOfSolids", RevitDB.SolidGeometry.Polymesh); // Export as mesh, not ACIS
                    
                    // IMPORTANT: These properties prevent linked model export as XREFs
                    TrySetProperty(options, "TargetUnit", RevitDB.ExportUnit.Default);
                    
                    // Try to find and set ACAPreference to Geometry (not AEC objects)
                    var acaPrefType = typeof(RevitDB.DWGExportOptions).Assembly
                        .GetTypes()
                        .FirstOrDefault(t => t.Name == "ACAObjectPreference");
                    if (acaPrefType != null)
                    {
                        var geometryValue = Enum.Parse(acaPrefType, "Geometry");
                        TrySetProperty(options, "ACAPreference", geometryValue);
                        System.Diagnostics.Debug.WriteLine("[DWG Override] Set ACAPreference = Geometry");
                    }
                    
                    System.Diagnostics.Debug.WriteLine("[DWG Override] All XREF options disabled - will export single file per sheet");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[DWG Override] Enabling ExportingAreas for XREF export");
                    TrySetProperty(options, "ExportingAreas", true);
                }
                
                // Override other settings
                options.FileVersion = GetDWGVersion(settings.DWGVersion);
                options.SharedCoords = settings.UseSharedCoordinates;
                
                // Advanced settings from DA4R-DwgExporter
                TrySetProperty(options, "HideScopeBox", settings.HideScopeBox);
                TrySetProperty(options, "HideReferencePlane", settings.HideReferencePlane);
                TrySetProperty(options, "HideUnreferenceViewTags", settings.HideUnreferenceViewTags);
                TrySetProperty(options, "PreserveCoincidentLines", settings.PreserveCoincidentLines);
                
                System.Diagnostics.Debug.WriteLine($"[DWG Override] FileVersion: {options.FileVersion}");
                System.Diagnostics.Debug.WriteLine($"[DWG Override] SharedCoords: {options.SharedCoords}");
                System.Diagnostics.Debug.WriteLine($"[DWG Override] HideScopeBox: {settings.HideScopeBox}");
                System.Diagnostics.Debug.WriteLine($"[DWG Override] HideReferencePlane: {settings.HideReferencePlane}");
                System.Diagnostics.Debug.WriteLine($"[DWG Override] PreserveCoincidentLines: {settings.PreserveCoincidentLines}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error overriding DWG options: {ex.Message}");
            }
        }

        /// <summary>
        /// Try to set property using reflection (for compatibility across Revit versions)
        /// </summary>
        private void TrySetProperty(RevitDB.DWGExportOptions options, string propertyName, object value)
        {
            try
            {
                var property = typeof(RevitDB.DWGExportOptions).GetProperty(propertyName);
                if (property != null && property.CanWrite)
                {
                    property.SetValue(options, value);
                    System.Diagnostics.Debug.WriteLine($"[DWG Override] Set {propertyName} = {value}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[DWG Override] Property {propertyName} not found or read-only");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DWG Override] Failed to set {propertyName}: {ex.Message}");
            }
        }

        /// <summary>
        /// Create default DWG export options
        /// </summary>
        private RevitDB.DWGExportOptions CreateDefaultDWGOptions(PSDWGExportSettings settings)
        {
            var options = new RevitDB.DWGExportOptions();
            
            System.Diagnostics.Debug.WriteLine("[DWG Default] Creating default DWG options");
            
            // Basic settings
            options.FileVersion = GetDWGVersion(settings.DWGVersion);
            options.SharedCoords = settings.UseSharedCoordinates;
            
            // CRITICAL: Disable XREF export by default (like DIRoots)
            if (!settings.ExportViewsOnSheets)
            {
                System.Diagnostics.Debug.WriteLine("[DWG Default] Disabling XREF export");
                TrySetProperty(options, "ExportingAreas", false);
                TrySetProperty(options, "MergedViews", false);
                TrySetProperty(options, "ExportOfSolids", RevitDB.SolidGeometry.Polymesh);
                
                // Set ACAPreference to Geometry (not AEC)
                var acaPrefType = typeof(RevitDB.DWGExportOptions).Assembly
                    .GetTypes()
                    .FirstOrDefault(t => t.Name == "ACAObjectPreference");
                if (acaPrefType != null)
                {
                    var geometryValue = Enum.Parse(acaPrefType, "Geometry");
                    TrySetProperty(options, "ACAPreference", geometryValue);
                }
            }
            else
            {
                TrySetProperty(options, "ExportingAreas", true);
            }
            
            // Quality settings
            TrySetProperty(options, "Colors", GetEnumValue("ExportColorMode", "TrueColorPerView"));
            
            System.Diagnostics.Debug.WriteLine("[DWG Default] Default DWG export options created");
            
            return options;
        }

        /// <summary>
        /// Get enum value by name using reflection
        /// </summary>
        private object GetEnumValue(string enumTypeName, string valueName)
        {
            try
            {
                var enumType = typeof(RevitDB.DWGExportOptions).Assembly
                    .GetTypes()
                    .FirstOrDefault(t => t.Name == enumTypeName && t.IsEnum);
                    
                if (enumType != null)
                {
                    return Enum.Parse(enumType, valueName);
                }
            }
            catch { }
            
            return null;
        }

        /// <summary>
        /// Export sheets to DWG format (DiRoots style - SHEET ONLY, NO XREF views)
        /// NOW WITH OPTION: Temporarily unload linked models to create single DWG file (no XREFs)
        /// </summary>
        public bool ExportToDWG(List<RevitDB.ViewSheet> sheets, PSDWGExportSettings settings)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DWG Export] ========================================");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] Starting SHEET-ONLY export (DiRoots style)");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] Sheets count: {sheets.Count}");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] Output folder: {settings.OutputFolder}");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] ExportViewsOnSheets: {settings.ExportViewsOnSheets} (MUST be FALSE)");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] ========================================");
                
                // FORCE settings to prevent XREF export
                settings.ExportViewsOnSheets = false;  
                System.Diagnostics.Debug.WriteLine($"[DWG Export] ExportViewsOnSheets = FALSE");
                
                // DiRoots method: Do NOT unload links - rely on DWGExportOptions to prevent XREF export
                System.Diagnostics.Debug.WriteLine($"[DWG Export] ?? Using DiRoots clean export method - ExportingAreas=FALSE");
                
                // Create DWG export options for SHEET ONLY
                RevitDB.DWGExportOptions dwgOptions = CreateSheetOnlyDWGOptions(settings);
                
                int successCount = 0;
                int failCount = 0;
                
                // ✅ ADD: StatusBar progress integration (giống StatusBar Demo)
                RevitProgressBarUtils.Run($"Exporting {sheets.Count} sheets to DWG", sheets, (sheet) =>
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine($"[DWG Export] ---");
                        System.Diagnostics.Debug.WriteLine($"[DWG Export] Processing: {sheet.SheetNumber} - {sheet.Name}");
                        
                        // Generate clean filename like DiRoots: P102-PLAN - L2 SANITARY
                        string fileName = GenerateDiRootsFileName(sheet);
                        System.Diagnostics.Debug.WriteLine($"[DWG Export] Generated filename: {fileName}");
                        
                        var outputPath = settings.CreateSubfolders 
                            ? FileNameGenerator.GenerateSubfolderPath(sheet, _document, settings)
                            : settings.OutputFolder;
                        
                        System.Diagnostics.Debug.WriteLine($"[DWG Export] Output path: {outputPath}");
                        
                        if (!Directory.Exists(outputPath))
                        {
                            Directory.CreateDirectory(outputPath);
                            System.Diagnostics.Debug.WriteLine($"[DWG Export] Created directory: {outputPath}");
                        }
                        
                        // KEY: Only export the SHEET element ID, NOT views on the sheet
                        ICollection<RevitDB.ElementId> sheetOnly = new List<RevitDB.ElementId> { sheet.Id };
                        
                        System.Diagnostics.Debug.WriteLine($"[DWG Export] ElementIds: [{string.Join(", ", sheetOnly.Select(id => id.IntegerValue))}]");
                        System.Diagnostics.Debug.WriteLine($"[DWG Export] Calling Document.Export()...");
                        
                        bool success = _document.Export(outputPath, fileName, sheetOnly, dwgOptions);
                            
                            if (success)
                            {
                                successCount++;
                                System.Diagnostics.Debug.WriteLine($"[DWG Export] ? SUCCESS - File created: {outputPath}\\{fileName}.dwg");
                                
                                // Verify file exists
                                string fullPath = Path.Combine(outputPath, fileName + ".dwg");
                                if (File.Exists(fullPath))
                                {
                                    FileInfo fi = new FileInfo(fullPath);
                                    System.Diagnostics.Debug.WriteLine($"[DWG Export] ? File verified - Size: {fi.Length / 1024} KB");
                                    
                                    // Check if XREF files were created
                                    bool hasXRefs = DWGCleanupManager.HasXRefReferences(fullPath);
                                    
                                    if (hasXRefs)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[DWG Export] ? XREF files detected - Revit created multiple files");
                                        System.Diagnostics.Debug.WriteLine($"[DWG Export] ?? Attempting cleanup...");
                                        
                                        // Try AutoCAD BIND first (best option - merges geometry)
                                        bool bindSuccess = AutoCADBindManager.BindXRefsInDWG(fullPath, deleteXRefFiles: true);
                                        
                                        if (bindSuccess)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"[DWG Export] ? AutoCAD BIND SUCCESS - XREFs merged into single file!");
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"[DWG Export] ? AutoCAD BIND failed/skipped");
                                            System.Diagnostics.Debug.WriteLine($"[DWG Export] ? WARNING: Multiple XREF files remain - manual cleanup needed");
                                            // DO NOT delete XREFs if BIND failed - file would be empty!
                                        }
                                    }
                                    else
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[DWG Export] ? NO XREF files created - Clean single file export!");
                                    }
                                    
                                    // Check final result
                                    var finalFi = new FileInfo(fullPath);
                                    if (finalFi.Exists)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[DWG Export] ? FINAL: {Path.GetFileName(fullPath)} - {finalFi.Length / 1024} KB");
                                    }
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"[DWG Export] ? WARNING: Export returned success but file not found!");
                                }
                            }
                            else
                            {
                                failCount++;
                                System.Diagnostics.Debug.WriteLine($"[DWG Export] ? FAILED: Document.Export() returned FALSE for {sheet.SheetNumber}");
                                System.Diagnostics.Debug.WriteLine($"[DWG Export] No exception thrown - check Revit's DWG export settings");
                            }
                    } // Close outer try block
                    catch (Exception ex)
                    {
                        failCount++;
                        System.Diagnostics.Debug.WriteLine($"[DWG Export] ? EXCEPTION for {sheet.SheetNumber}:");
                        System.Diagnostics.Debug.WriteLine($"[DWG Export]   Message: {ex.Message}");
                        System.Diagnostics.Debug.WriteLine($"[DWG Export]   Type: {ex.GetType().Name}");
                        if (ex.InnerException != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DWG Export]   Inner: {ex.InnerException.Message}");
                        }
                        System.Diagnostics.Debug.WriteLine($"[DWG Export]   Stack: {ex.StackTrace}");
                    }
                }); // ✅ Close RevitProgressBarUtils.Run()
                
                System.Diagnostics.Debug.WriteLine($"[DWG Export] ========================================");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] EXPORT COMPLETED");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] Success: {successCount}, Failed: {failCount}");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] Expected files in: {settings.OutputFolder}");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] ========================================");
                
                return successCount > 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DWG Export] ??? MANAGER ERROR ???");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] Message: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[DWG Export] Type: {ex.GetType().Name}");
                if (ex.InnerException != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[DWG Export] Inner: {ex.InnerException.Message}");
                }
                System.Diagnostics.Debug.WriteLine($"[DWG Export] Stack: {ex.StackTrace}");
                return false;
            }
        }
        
        /// <summary>
        /// Generate filename like DiRoots: P102-PLAN - L2 SANITARY
        /// </summary>
        private string GenerateDiRootsFileName(RevitDB.ViewSheet sheet)
        {
            string fileName = $"{sheet.SheetNumber}-{sheet.Name}";
            
            // Clean invalid characters
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '-');
            }
            
            return fileName;
        }
        
        /// <summary>
        /// Create DWG export options for SHEET ONLY (DiRoots style - NO views as XREFs)
        /// </summary>
        /// <summary>
        /// Create or modify DWG export setup in Revit to ensure SHEET-ONLY export
        /// This is more reliable than just passing options to Document.Export()
        /// </summary>
        /// <summary>
        /// Create DWG export options for SHEET-ONLY export with FULL GEOMETRY
        /// ULTRA CLEAN method - prevents XREF files AND includes all geometry
        /// </summary>
        private RevitDB.DWGExportOptions CreateSheetOnlyDWGOptions(PSDWGExportSettings settings)
        {
            var options = new RevitDB.DWGExportOptions();
            System.Diagnostics.Debug.WriteLine("[DWG Options] Creating ULTRA CLEAN options");
            
            // =====================================================
            // CRITICAL SETTINGS - PREVENT XREF FILES
            // =====================================================
            
            // 1. MOST IMPORTANT: Do NOT export views/areas as separate files
            TrySetProperty(options, "ExportingAreas", false);
            System.Diagnostics.Debug.WriteLine("[DWG Options] ? ExportingAreas = FALSE");
            
            // 2. Do NOT merge views
            TrySetProperty(options, "MergedViews", false);
            System.Diagnostics.Debug.WriteLine("[DWG Options] ? MergedViews = FALSE");
            
            // 3. TRY FALSE for SharedCoords - might prevent view splitting
            options.SharedCoords = false;  // ? KEY CHANGE!
            System.Diagnostics.Debug.WriteLine("[DWG Options] ? SharedCoords = FALSE (prevent view splitting)");
            
            // 4. Additional blocking
            TrySetProperty(options, "ExportRoomsAndAreas", false);
            TrySetProperty(options, "PropOverrides", false);
            System.Diagnostics.Debug.WriteLine("[DWG Options] ? ExportRoomsAndAreas = FALSE");
            System.Diagnostics.Debug.WriteLine("[DWG Options] ? PropOverrides = FALSE");
            
            // =====================================================
            // GEOMETRY SETTINGS - INCLUDE ALL CONTENT
            // =====================================================
            
            // Export as simple Polymesh
            options.ExportOfSolids = RevitDB.SolidGeometry.Polymesh;
            System.Diagnostics.Debug.WriteLine("[DWG Options] ? ExportOfSolids = Polymesh");
            
            // Export as pure GEOMETRY (not AEC objects)
            var acaPrefType = typeof(RevitDB.DWGExportOptions).Assembly
                .GetTypes()
                .FirstOrDefault(t => t.Name == "ACAObjectPreference");
            if (acaPrefType != null)
            {
                var geometryValue = Enum.Parse(acaPrefType, "Geometry");
                TrySetProperty(options, "ACAPreference", geometryValue);
                System.Diagnostics.Debug.WriteLine("[DWG Options] ? ACAPreference = Geometry");
            }
            
            // Target unit - use Millimeter instead of Default
            try
            {
                var targetUnit = Enum.Parse(typeof(RevitDB.ExportUnit), "Millimeter");
                TrySetProperty(options, "TargetUnit", targetUnit);
                System.Diagnostics.Debug.WriteLine("[DWG Options] ? TargetUnit = Millimeter");
            }
            catch
            {
                TrySetProperty(options, "TargetUnit", RevitDB.ExportUnit.Default);
                System.Diagnostics.Debug.WriteLine("[DWG Options] ? TargetUnit = Default (fallback)");
            }
            
            // =====================================================
            // COLOR & QUALITY SETTINGS
            // =====================================================
            
            // Use IndexColors instead of TrueColor - simpler, less likely to split
            TrySetProperty(options, "Colors", GetEnumValue("ExportColorMode", "IndexColors"));
            System.Diagnostics.Debug.WriteLine("[DWG Options] ? Colors = IndexColors");
            
            TrySetProperty(options, "LineScaling", GetEnumValue("LineScaling", "ViewScale"));
            System.Diagnostics.Debug.WriteLine("[DWG Options] ? LineScaling = ViewScale");
            
            // Hide unnecessary elements
            TrySetProperty(options, "HideReferencePlane", true);
            TrySetProperty(options, "HideScopeBox", true);
            TrySetProperty(options, "HideUnreferenceViewTags", true);
            
            // =====================================================
            // BASIC SETTINGS
            // =====================================================
            
            options.FileVersion = GetDWGVersion(settings.DWGVersion);
            System.Diagnostics.Debug.WriteLine($"[DWG Options] FileVersion = {options.FileVersion}");
            
            System.Diagnostics.Debug.WriteLine("[DWG Options] ================================");
            System.Diagnostics.Debug.WriteLine("[DWG Options] ? ULTRA CLEAN EXPORT configured!");
            System.Diagnostics.Debug.WriteLine("[DWG Options] ? Should export FULL GEOMETRY into 1 file");
            System.Diagnostics.Debug.WriteLine("[DWG Options] ================================");
            
            return options;
        }

        private RevitDB.ACADVersion GetDWGVersion(string version)
        {
            switch (version?.ToLower())
            {
                case "2018":
                    return RevitDB.ACADVersion.R2018;
                case "2013":
                    return RevitDB.ACADVersion.R2013;
                case "2010":
                    return RevitDB.ACADVersion.R2010;
                case "2007":
                    return RevitDB.ACADVersion.R2007;
                default:
                    return RevitDB.ACADVersion.R2018;
            }
        }

        /// <summary>
        /// Unload all linked models temporarily to prevent XREF file creation
        /// This is the DiRoots method for single-file DWG export
        /// </summary>
        private List<RevitDB.ElementId> UnloadAllLinkedModels()
        {
            var unloadedLinks = new List<RevitDB.ElementId>();
            
            try
            {
                // Find all RevitLinkType elements (loaded link definitions)
                var linkTypes = new RevitDB.FilteredElementCollector(_document)
                    .OfClass(typeof(RevitDB.RevitLinkType))
                    .Cast<RevitDB.RevitLinkType>()
                    .Where(lt => lt.GetLinkedFileStatus() == RevitDB.LinkedFileStatus.Loaded)
                    .ToList();

                System.Diagnostics.Debug.WriteLine($"[DWG Export] Found {linkTypes.Count} loaded link types");

                if (linkTypes.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[DWG Export] No linked models to unload");
                    return unloadedLinks;
                }

                // Unload each link using Transaction
                using (RevitDB.Transaction trans = new RevitDB.Transaction(_document, "Unload Links for DWG Export"))
                {
                    trans.Start();
                    
                    foreach (var linkType in linkTypes)
                    {
                        try
                        {
                            var linkName = linkType.Name;
                            System.Diagnostics.Debug.WriteLine($"[DWG Export]   Unloading: {linkName}");
                            
                            // Unload the link
                            linkType.Unload(null);
                            unloadedLinks.Add(linkType.Id);
                            
                            System.Diagnostics.Debug.WriteLine($"[DWG Export]   ? Unloaded: {linkName}");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DWG Export]   ? Failed to unload {linkType.Name}: {ex.Message}");
                        }
                    }
                    
                    trans.Commit();
                }

                System.Diagnostics.Debug.WriteLine($"[DWG Export] Successfully unloaded {unloadedLinks.Count} linked models");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DWG Export] ERROR unloading links: {ex.Message}");
            }

            return unloadedLinks;
        }

        /// <summary>
        /// Reload linked models after export
        /// </summary>
        private int ReloadLinkedModels(List<RevitDB.ElementId> linkIds)
        {
            int reloadedCount = 0;
            
            try
            {
                if (linkIds == null || linkIds.Count == 0)
                {
                    return 0;
                }

                using (RevitDB.Transaction trans = new RevitDB.Transaction(_document, "Reload Links after DWG Export"))
                {
                    trans.Start();
                    
                    foreach (var linkId in linkIds)
                    {
                        try
                        {
                            var linkType = _document.GetElement(linkId) as RevitDB.RevitLinkType;
                            if (linkType != null)
                            {
                                var linkName = linkType.Name;
                                System.Diagnostics.Debug.WriteLine($"[DWG Export]   Reloading: {linkName}");
                                
                                // Reload the link
                                linkType.Reload();
                                reloadedCount++;
                                
                                System.Diagnostics.Debug.WriteLine($"[DWG Export]   ? Reloaded: {linkName}");
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DWG Export]   ? Failed to reload link {linkId}: {ex.Message}");
                        }
                    }
                    
                    trans.Commit();
                }

                System.Diagnostics.Debug.WriteLine($"[DWG Export] Successfully reloaded {reloadedCount} linked models");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DWG Export] ERROR reloading links: {ex.Message}");
            }

            return reloadedCount;
        }
    }
}
