using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Quoc_MEP.Export.Commands
{
    /// <summary>
    /// External Event Handler for IFC Export
    /// This runs in Revit API context, allowing Transaction creation
    /// </summary>
    public class IFCExportHandler : IExternalEventHandler
    {
        // Export parameters (set before raising event)
        public Document Document { get; set; }
        public List<View3D> Views3D { get; set; }
        public Models.IFCExportSettings Settings { get; set; }
        public string OutputFolder { get; set; }
        public Action<string> LogCallback { get; set; }
        
        // Progress callback (called after each file export completes)
        public Action<string, bool> ProgressCallback { get; set; }
        
        // Completion callback (called after export finishes)
        public Action<bool> CompletionCallback { get; set; }
        
        // Export result
        public bool ExportResult { get; private set; }

        public void Execute(UIApplication app)
        {
            try
            {
                if (Document == null || Views3D == null || Views3D.Count == 0)
                {
                    LogCallback?.Invoke("❌ IFC Export: Invalid parameters");
                    ExportResult = false;
                    return;
                }

                LogCallback?.Invoke($"[IFC ExternalEvent] Starting export with {Views3D.Count} views");
                
                // Now we're in Revit API context - Transaction is allowed!
                var ifcManager = new Managers.IFCExportManager(Document);
                ExportResult = ifcManager.Export3DViewsToIFC(Views3D, Settings, OutputFolder, LogCallback, ProgressCallback);
                
                LogCallback?.Invoke($"[IFC ExternalEvent] Export completed: {(ExportResult ? "SUCCESS" : "FAILED")}");
                
                // Notify UI that export is complete
                CompletionCallback?.Invoke(ExportResult);
            }
            catch (Exception ex)
            {
                LogCallback?.Invoke($"❌ IFC ExternalEvent Exception: {ex.Message}");
                LogCallback?.Invoke($"   Stack: {ex.StackTrace}");
                ExportResult = false;
                
                // Notify UI of failure
                CompletionCallback?.Invoke(false);
            }
        }

        public string GetName()
        {
            return "IFC Export Handler";
        }
    }
}
