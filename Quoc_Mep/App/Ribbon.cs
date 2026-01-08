using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using System.Reflection;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System;
// using ricaun.Revit.UI; // ⚠️ Comment out for minimal build

namespace Quoc_MEP
{
    [Transaction(TransactionMode.Manual)]
    class Ribbon : IExternalApplication
    {
        private readonly string nameSpace = "Quoc_MEP.";
        private readonly string tabName = "Quoc_MEP";
        private readonly string path = Assembly.GetExecutingAssembly().Location;
        private readonly string resourcePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources");
        
        // Giữ reference để dùng trong ApplicationInitialized event
        private static UIControlledApplication _uiCtrlApp;
        
        // Static UIApplication để Panel có thể access
        public static UIApplication StaticUIApp { get; private set; }

        private BitmapImage LoadImageFromFile(string fileName)
        {
            try
            {
                string fullPath = System.IO.Path.Combine(resourcePath, fileName);
                if (System.IO.File.Exists(fullPath))
                {
                    BitmapImage bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(fullPath);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    return bitmap;
                }
            }
            catch { }
            return null;
        }

        private BitmapImage Convert(Bitmap bimapImage)
        {
            if (bimapImage == null)
                return null;

            try
            {
                MemoryStream memory = new MemoryStream();
                bimapImage.Save(memory, ImageFormat.Png);
                memory.Position = 0;
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
            catch
            {
                return null;
            }
        }



        private void SetPull_Image(PulldownButton pull, Bitmap imageSource)
        {
            if (imageSource != null)
                pull.LargeImage = Convert(imageSource);
        }

        private void SetPush_Image(PushButton push, Bitmap imageSource)
        {
            if (imageSource != null)
                push.LargeImage = Convert(imageSource);
        }

        private void SetPush_ImageFromFile(PushButton push, string fileName)
        {
            var image = LoadImageFromFile(fileName);
            if (image != null)
                push.LargeImage = image;
        }

        private void MyPush(RibbonPanel panel, PushButtonData data, Bitmap bitmap, string description)
        {
            PushButton push = panel.AddItem(data) as PushButton;
            push.ToolTip = description;
            if (bitmap != null)
                SetPush_Image(push, bitmap);
        }

        private void MyPush(RibbonPanel panel, PushButtonData data, string iconFileName, string description)
        {
            PushButton push = panel.AddItem(data) as PushButton;
            push.ToolTip = description;
            SetPush_ImageFromFile(push, iconFileName);
        }


        private void MyPull(RibbonPanel panel, PulldownButtonData data, Bitmap bitmap, List<PushButtonData> list, string des, List<string> listdes)
        {
            PulldownButton pulldown = panel.AddItem(data) as PulldownButton;
            for (int i = 0; i < list.Count; i++)
            {
                PushButtonData pushdata = list[i];
                PushButton bt = pulldown.AddPushButton(pushdata);
                if (listdes.Count > 0)
                {
                    bt.ToolTip = listdes[i];
                }
            }
            if (bitmap != null)
                SetPull_Image(pulldown, bitmap);
            pulldown.ToolTip = des;
        }


        private void Modify_Tool (RibbonPanel panel)
        {
            //move_Connect_Align
            PushButtonData move_Connect_Align = new PushButtonData("move_Connect_Align", "Move Connect" + '\n' + "Align", path, nameSpace + "MoveConnectCommand");
            MyPush(panel, move_Connect_Align, "pipe.png", "Click on a destination MEP family, and then on an MEP family you'd like to move. The second family is moved so that the two closest connectors meet, align and connect.");

            //place_Family
            PushButtonData place_Family = new PushButtonData("place_Family", "Place Family", path, nameSpace + "PlaceFamilyCmd");
            MyPush(panel, place_Family, "place.png", "Place family from location of block in file link CAD.");

            //draw_Pipe
            PushButtonData draw_Pipe = new PushButtonData("draw_Pipe", "Draw Pipe", path, nameSpace + "DrawPipe");
            MyPush(panel, draw_Pipe, "pipe.png", "Create and align multiple pipes.");

            //change_Length
            PushButtonData change_Length = new PushButtonData("change_Length", "Change" + '\n' + "Length", path, nameSpace + "ChangeLengthcmd");
            MyPush(panel, change_Length, "pipe.png", "Change the length of selected pipe or duct.");

            //split_Duct
            PushButtonData split_Duct = new PushButtonData("split_Duct", "Split Duct", path, nameSpace + "SplitDuctCmd");
            MyPush(panel, split_Duct, "split.png", "Divide the selected section of the duct into multiple duct sections of equal length.");

            //mep_UpDown
            PushButtonData mep_UpDown = new PushButtonData("mep_UpDown", "MEP Up Down", path, nameSpace + "MEPUpDownCmd");
            MyPush(panel, mep_UpDown, "up.png", "A tool that supports resolving clashes within the MEP system.");

            //Rotate
            PushButtonData Rotate = new PushButtonData("Rotate", "Rotate Element", path, nameSpace + "RotateElementsCommand");
            MyPush(panel, Rotate, "rotate-32.png", "Rotate Element with Angle");

            // ⚠️ StatusBar Demo - Disabled for minimal build (requires ricaun package)
            // PushButtonData statusBar_Demo = new PushButtonData("statusBar_Demo", "StatusBar" + '\n' + "Demo", path, nameSpace + "StatusBarDemoCmd");
            // MyPush(panel, statusBar_Demo, "excel.png", "Demo ricaun.Revit.UI.StatusBar - Progress bar on Revit StatusBar.");

            //MEP Tools Panel (Dockable)
            PushButtonData mepToolsPanel = new PushButtonData("mepToolsPanel", "MEP Tools" + '\n' + "Panel", path, nameSpace + "ShowDockablePanelCommand");
            MyPush(panel, mepToolsPanel, "pipe.png", "Show/Hide MEP Tools Dockable Panel - Quick access to common MEP tools.");

            //Trans Data Para
            PushButtonData transDataPara = new PushButtonData("transDataPara", "Copy" + '\n' + "Parameters", path, nameSpace + "CopyParametersCommand");
            MyPush(panel, transDataPara, "excel.png", "Copy dữ liệu giữa các Parameters.");

        }

        private void Data_Tool(RibbonPanel panel)
        {
            //create_Sheet
            PushButtonData create_Sheet = new PushButtonData("create_Sheet", "Create " + '\n' + "Sheets", path, nameSpace + "SheetFromExcelCmd");
            MyPush(panel, create_Sheet, "sheet.png", "Create sheets from Excel data.");

            //export_Schedule
            PushButtonData export_Schedule = new PushButtonData("export_Schedule", "Export " + '\n' + "Schedule", path, nameSpace + "SheetFromExcelCmd");
            MyPush(panel, export_Schedule, "excel.png", "Export Schedule to Excel.");
        }

        private void Annotation_Tool(RibbonPanel panel)
        {
            //create_Sheet
            PushButtonData create_Sheet = new PushButtonData("create_Sheet", "Create " + '\n' + "Sheets", path, nameSpace + "SheetFromExcelCmd");
            MyPush(panel, create_Sheet, "sheet.png", "Create sheets from Excel data.");

            //export_Schedule
            PushButtonData export_Schedule = new PushButtonData("export_Schedule", "Export " + '\n' + "Schedule", path, nameSpace + "SheetFromExcelCmd");
            MyPush(panel, export_Schedule, "excel.png", "Export Schedule to Excel.");
        }
        // ============================================
        // ⚠️ EXPORT_TOOL - REMOVED
        // ============================================
        // private void Export_Tool(RibbonPanel panel)
        // {
        //     // Export+ (Main Export Tool with PDF, DWG, IFC, NWC)
        //     PushButtonData exportPlus = new PushButtonData("exportPlus", "Export+", path, nameSpace + "Export.SimpleExportCommand");
        //     MyPush(panel, exportPlus, "sheet.png", "Export+ Manager - Export Sheets to PDF, DWG, IFC, NWC with advanced options including Forge patterns.");
        //
        //     //create_Sheet
        //     PushButtonData create_Sheet = new PushButtonData("create_Sheet", "Create " + '\n' + "Sheets", path, nameSpace + "SheetFromExcelCmd");
        //     MyPush(panel, create_Sheet, "sheet.png", "Create sheets from Excel data.");
        //
        //     //export_Schedule
        //     PushButtonData export_Schedule = new PushButtonData("export_Schedule", "Export " + '\n' + "Schedule", path, nameSpace + "SheetFromExcelCmd");
        //     MyPush(panel, export_Schedule, "excel.png", "Export Schedule to Excel.");
        // }

        public Result OnStartup(UIControlledApplication application)
        {
            // Lưu reference để dùng sau
            _uiCtrlApp = application;

            //tạo tab có tên là "Quoc_MEP"
            application.CreateRibbonTab(tabName);

            //tạo panel
            RibbonPanel ModifyPanel = application.CreateRibbonPanel(tabName, "Modify");
            RibbonPanel dataPanel = application.CreateRibbonPanel(tabName, "Data");
            RibbonPanel AnnotationPanel = application.CreateRibbonPanel(tabName, "Annotation");
            // ⚠️ REMOVED ExportPanel
            // RibbonPanel ExportPanel = application.CreateRibbonPanel(tabName, "Export");

            // Note: CustomPanelTitleBarBackground not available on RibbonPanel in Revit 2020
            // ModifyPanel.CustomPanelTitleBarBackground = System.Windows.Media.Brushes.LightBlue;
            // dataPanel.CustomPanelTitleBarBackground = System.Windows.Media.Brushes.LightGreen;
            // AnnotationPanel.CustomPanelTitleBarBackground = System.Windows.Media.Brushes.LightCoral;
            // ExportPanel.CustomPanelTitleBarBackground = System.Windows.Media.Brushes.LightGoldenrodYellow;

            Modify_Tool(ModifyPanel);
            Data_Tool(dataPanel);
            Annotation_Tool(AnnotationPanel);
            // ⚠️ REMOVED Export_Tool
            // Export_Tool(ExportPanel);

            // ============================================================
            // Register DockablePane trong ApplicationInitialized event
            // Đây là thời điểm ĐÚNG để register theo Revit API
            // ============================================================
            application.ControlledApplication.ApplicationInitialized += OnApplicationInitialized;
            
            Logger.Info("Ribbon OnStartup completed");

            return Result.Succeeded;
        }

        private void OnApplicationInitialized(object sender, Autodesk.Revit.DB.Events.ApplicationInitializedEventArgs e)
        {
            try
            {
                Logger.StartOperation("OnApplicationInitialized");
                
                // GET UIApplication và lưu vào static field
                var app = sender as Autodesk.Revit.ApplicationServices.Application;
                if (app != null)
                {
                    // Create UIApplication from Application
                    StaticUIApp = new UIApplication(app);
                    Logger.Info($"✅ StaticUIApp created: {StaticUIApp != null}");
                    
                    // Set vào RevitContext luôn
                    RevitContext.UIApplication = StaticUIApp;
                    Logger.Info($"✅ RevitContext.IsInitialized: {RevitContext.IsInitialized}");
                }
                
                // Tạo Panel instance với UIApplication
                var panel = new MEPToolsPanel(StaticUIApp);
                
                // Dùng UIControlledApplication đã lưu trong OnStartup
                if (_uiCtrlApp != null)
                {
                    var paneId = new DockablePaneId(MEPToolsPanel.PanelGuid);
                    _uiCtrlApp.RegisterDockablePane(paneId, "MEP Tools Panel", panel);
                    
                    Logger.Info("✅ DockablePane registered in ApplicationInitialized");
                }
                else
                {
                    Logger.Error("UIControlledApplication is null", null);
                }
                
                Logger.EndOperation("OnApplicationInitialized");
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to register DockablePane", ex);
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            Logger.Info("Ribbon OnShutdown completed");
            return Result.Succeeded;
        }
    }
}
