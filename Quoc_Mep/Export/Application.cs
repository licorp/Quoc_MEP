using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Reflection;
using System.Diagnostics;

namespace Quoc_MEP
{
    [Transaction(TransactionMode.Manual)]
    public class Application : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            Debug.WriteLine("[ExportPlus] Application OnStartup started");
            
            // Tạo ribbon tab "Licorp"
            string tabName = "Licorp";
            try
            {
                application.CreateRibbonTab(tabName);
                Debug.WriteLine("[ExportPlus] Created ribbon tab: " + tabName);
            }
            catch
            {
                // Tab đã tồn tại, sử dụng tab hiện có
                Debug.WriteLine("[ExportPlus] Ribbon tab already exists: " + tabName);
            }
            
            // Tạo ribbon panel "Data Tools"
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Data Tools");
            
            // Thêm push button "ExportPlus"
            PushButtonData buttonData = new PushButtonData(
                "ExportPlusButton", 
                "Export+", 
                Assembly.GetExecutingAssembly().Location, 
                "Quoc_MEP.Export.SimpleExportCommand");
            
            PushButton pushButton = panel.AddItem(buttonData) as PushButton;
            pushButton.ToolTip = "Export sheets to PDF, DWG, IFC and other formats with advanced settings";
            pushButton.LongDescription = "ExportPlus by Licorp - Professional batch export tool for Revit\n\n" +
                "Features:\n" +
                "• Export to PDF, DWG, IFC, NWC, Images\n" +
                "• Profile management\n" +
                "• Custom file naming\n" +
                "• Batch processing";
            
            Debug.WriteLine("[ExportPlus] Ribbon setup completed successfully");
            return Result.Succeeded;
        }
        
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}