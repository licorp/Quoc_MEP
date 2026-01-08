using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Nice3point.Revit.Toolkit.Options;

namespace Quoc_MEP
{
    /// <summary>
    /// Command để show/hide MEP Tools Dockable Panel
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ShowDockablePanelCommand : IExternalCommand
    {
        private static DockablePaneId _paneId = new DockablePaneId(MEPToolsPanel.PanelGuid);

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                
                Logger.StartOperation("ShowDockablePanelCommand.Execute");
                Logger.Info($"UIApplication available: {uiapp != null}");
                
                // ============================================================
                // BƯỚC 1: Set UIApplication cho Panel
                // Panel đã được tạo trong Ribbon.OnApplicationInitialized
                // Bây giờ chỉ cần cung cấp UIApplication để Panel hoạt động
                // ============================================================
                MEPToolsPanel.SetUIApplication(uiapp);
                RevitContext.UIApplication = uiapp;
                Logger.Info($"✅ UIApplication set - RevitContext.IsInitialized={RevitContext.IsInitialized}");
                
                // ============================================================
                // BƯỚC 2: Get DockablePane (đã được register trong OnStartup)
                // ============================================================
                DockablePane dockablePane = uiapp.GetDockablePane(_paneId);
                
                if (dockablePane == null)
                {
                    Logger.Error("DockablePane not found - may not be registered", null);
                    TaskDialog.Show("Error", 
                        "MEP Tools Panel not found!\n\n" +
                        "This usually means the panel was not registered during startup.\n" +
                        "Please restart Revit.");
                    return Result.Failed;
                }
                
                Logger.Info($"DockablePane found - IsShown={dockablePane.IsShown()}");
                
                // ============================================================
                // BƯỚC 3: Toggle Show/Hide với Nice3point Toolkit
                // ============================================================
                if (dockablePane != null)
                {
                    // ✨ DÙNG Nice3point.Revit.Toolkit patterns
                    bool isCurrentlyShown = dockablePane.IsShown();
                    
                    if (isCurrentlyShown)
                    {
                        dockablePane.Hide();
                        Logger.Info("📦 Panel hidden");
                    }
                    else
                    {
                        dockablePane.Show();
                        Logger.Info("📦 Panel shown - UIApplication available");
                    }
                }
                else
                {
                    Logger.Error("DockablePane is null after retrieval", null);
                    
                    // ✨ DÙNG Nice3point TaskDialog helpers (nếu có)
                    TaskDialog.Show("Error", 
                        "MEP Tools Panel not found!\n\n" +
                        "Please restart Revit and try again.");
                    return Result.Failed;
                }                Logger.EndOperation("ShowDockablePanelCommand.Execute");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Logger.Error("ShowDockablePanelCommand failed", ex);
                TaskDialog.Show("Error", 
                    $"Cannot show/hide dockable panel:\n\n" +
                    $"{ex.Message}\n\n" +
                    $"Stack Trace:\n{ex.StackTrace}");
                return Result.Failed;
            }
        }
    }
}
