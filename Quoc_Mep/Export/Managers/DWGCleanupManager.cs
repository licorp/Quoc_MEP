using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Quoc_MEP.Export.Managers
{
    /// <summary>
    /// Clean up DWG files after export - Remove XREF references and delete XREF files
    /// This is the ONLY way to get single DWG file with linked model content!
    /// </summary>
    public class DWGCleanupManager
    {
        /// <summary>
        /// Clean up DWG export - Delete all XREF files created by Revit
        /// </summary>
        public static void CleanupDWGExport(string mainDwgPath)
        {
            try
            {
                if (!File.Exists(mainDwgPath))
                {
                    Debug.WriteLine($"[DWG Cleanup] File not found: {mainDwgPath}");
                    return;
                }

                Debug.WriteLine($"[DWG Cleanup] ========================================");
                Debug.WriteLine($"[DWG Cleanup] Starting cleanup for: {Path.GetFileName(mainDwgPath)}");

                var directory = Path.GetDirectoryName(mainDwgPath);
                var mainFileName = Path.GetFileNameWithoutExtension(mainDwgPath);
                
                // Find the MAIN file (usually the smallest or the one without suffix)
                var allDwgFiles = Directory.GetFiles(directory, "*.dwg")
                    .Where(f => Path.GetFileNameWithoutExtension(f).StartsWith(mainFileName, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(f => new FileInfo(f).Length) // Sort by size - main file usually smallest
                    .ToList();

                Debug.WriteLine($"[DWG Cleanup] Found {allDwgFiles.Count} related DWG files");

                if (allDwgFiles.Count <= 1)
                {
                    Debug.WriteLine($"[DWG Cleanup] Only 1 file found - no cleanup needed");
                    return;
                }

                // The MAIN file is the first one (smallest) or exact match
                string realMainFile = allDwgFiles.FirstOrDefault(f => 
                    Path.GetFileNameWithoutExtension(f).Equals(mainFileName, StringComparison.OrdinalIgnoreCase)
                ) ?? allDwgFiles[0];

                Debug.WriteLine($"[DWG Cleanup] Main file identified: {Path.GetFileName(realMainFile)}");

                // Delete all XREF files (files that are NOT the main file)
                int deletedCount = 0;
                foreach (var file in allDwgFiles)
                {
                    if (file.Equals(realMainFile, StringComparison.OrdinalIgnoreCase))
                    {
                        Debug.WriteLine($"[DWG Cleanup]   ✓ KEEP: {Path.GetFileName(file)} (main file)");
                        continue;
                    }

                    try
                    {
                        File.Delete(file);
                        deletedCount++;
                        Debug.WriteLine($"[DWG Cleanup]   ✗ DELETED: {Path.GetFileName(file)}");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[DWG Cleanup]   ⚠ Could not delete {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                Debug.WriteLine($"[DWG Cleanup] ========================================");
                Debug.WriteLine($"[DWG Cleanup] Cleanup completed: Deleted {deletedCount} XREF files");
                Debug.WriteLine($"[DWG Cleanup] ⚠ NOTE: Main DWG still contains XREF REFERENCES (will show 'XREF not found' when opened)");
                Debug.WriteLine($"[DWG Cleanup] ========================================");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DWG Cleanup] ERROR: {ex.Message}");
            }
        }

        /// <summary>
        /// Advanced cleanup: Use netDxf library to remove XREF references from DWG
        /// This requires netDxf NuGet package
        /// </summary>
        public static bool RemoveXRefReferences(string dwgPath)
        {
            try
            {
                Debug.WriteLine($"[DWG Cleanup] Attempting to remove XREF references from DWG...");
                Debug.WriteLine($"[DWG Cleanup] ⚠ This feature requires netDxf library (not implemented yet)");
                
                // TODO: Implement using netDxf or similar library
                // This would:
                // 1. Open DWG file
                // 2. Find all XREF blocks
                // 3. Convert XREFs to regular blocks OR remove them
                // 4. Save file
                
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DWG Cleanup] ERROR removing XREF references: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Check if a DWG file contains XREF references
        /// </summary>
        public static bool HasXRefReferences(string dwgPath)
        {
            try
            {
                // Simple check: Look for companion files
                var directory = Path.GetDirectoryName(dwgPath);
                var baseName = Path.GetFileNameWithoutExtension(dwgPath);
                
                var xrefFiles = Directory.GetFiles(directory, "*.dwg")
                    .Where(f => {
                        var name = Path.GetFileNameWithoutExtension(f);
                        return !name.Equals(baseName, StringComparison.OrdinalIgnoreCase) 
                            && name.StartsWith(baseName, StringComparison.OrdinalIgnoreCase);
                    })
                    .ToList();

                return xrefFiles.Count > 0;
            }
            catch
            {
                return false;
            }
        }
    }
}
