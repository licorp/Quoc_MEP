using System;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Nice3point.Revit.Extensions;

namespace Quoc_MEP
{
    public partial class MEPToolsPanel : Page, IDockablePaneProvider
    {
        public static Guid PanelGuid = new Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890");
        private static UIApplication _staticUIApp; // Static để giữ reference
        private UIApplication _uiapp;
        private static bool _traceListenerAdded = false;
        private static MEPToolsPanel _instance; // Singleton reference

        // ✨ ExternalEvents cho Change Length và Rotate
        private static ExternalEvent _changeLengthEvent;
        private static ChangeLengthEventHandler _changeLengthHandler;
        private static ExternalEvent _rotateEvent;
        private static RotateEventHandler _rotateHandler;

        public MEPToolsPanel()
        {
            InitializeComponent();
            
            // Lưu instance reference
            _instance = this;
            
            if (!_traceListenerAdded)
            {
                Trace.Listeners.Add(new DefaultTraceListener());
                _traceListenerAdded = true;
            }
            
            // Subscribe to Loaded event để lấy UIApplication khi Panel được show
            this.Loaded += OnPanelLoaded;
            
            Logger.Info("MEPToolsPanel created (UIApplication will be set when panel is shown)");
        }
        
        private void OnPanelLoaded(object sender, RoutedEventArgs e)
        {
            Logger.Info("=== Panel Loaded Event ===");
            
            // Panel loaded - check if UIApplication is available
            if (_staticUIApp != null)
            {
                Logger.Info($"✅ Panel loaded - UIApplication already set");
            }
            else if (RevitContext.IsInitialized)
            {
                _staticUIApp = RevitContext.UIApplication;
                _uiapp = RevitContext.UIApplication;
                Logger.Info("✅ Panel loaded - UIApplication retrieved from RevitContext");
            }
            else
            {
                // TRY HARDER: Tìm UIApplication từ Revit context
                Logger.Warning("⚠️ Panel loaded but UIApplication not available");
                Logger.Warning("⚠️ Trying to get UIApplication from Revit CommandManager...");
                
                try
                {
                    // Workaround: Lấy từ Autodesk.Windows.ComponentManager
                    var ribbonControl = Autodesk.Windows.ComponentManager.Ribbon;
                    if (ribbonControl != null)
                    {
                        Logger.Info("✅ Found Ribbon control - will get UIApplication when command executes");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("Cannot access ComponentManager", ex);
                }
            }
            
            Logger.Info($"Panel Loaded Summary: static={_staticUIApp != null}, instance={_uiapp != null}, context={RevitContext.IsInitialized}");
        }

        public MEPToolsPanel(UIApplication uiapp) : this()
        {
            SetUIApplicationInternal(uiapp);
            Logger.Info("MEPToolsPanel created with UIApplication parameter");
        }
        
        /// <summary>
        /// Public static method để set UIApplication từ bên ngoài (ShowDockablePanelCommand)
        /// </summary>
        public static void SetUIApplication(UIApplication uiapp)
        {
            if (_instance != null)
            {
                _instance.SetUIApplicationInternal(uiapp);
            }
            else
            {
                // Instance chưa tạo, chỉ lưu vào static
                _staticUIApp = uiapp;
                Logger.Info("Static UIApplication set (Panel instance not created yet)");
            }
        }
        
        /// <summary>
        /// Internal method để set UIApplication cho instance này
        /// </summary>
        private void SetUIApplicationInternal(UIApplication uiapp)
        {
            _uiapp = uiapp;
            _staticUIApp = uiapp;
            RevitContext.UIApplication = uiapp;
            Logger.Info($"✅ UIApplication set for Panel instance (IsInitialized={RevitContext.IsInitialized})");
            
            // ✨ TẠO ExternalEvents khi có UIApplication
            InitializeExternalEvents(uiapp);
        }
        
        /// <summary>
        /// Khởi tạo ExternalEvents - CHỈ GỌI 1 LẦN
        /// </summary>
        private static void InitializeExternalEvents(UIApplication uiapp)
        {
            if (_changeLengthEvent == null)
            {
                _changeLengthHandler = new ChangeLengthEventHandler();
                _changeLengthEvent = ExternalEvent.Create(_changeLengthHandler);
                Logger.Info("✅ ChangeLengthEvent created");
            }
            
            if (_rotateEvent == null)
            {
                _rotateHandler = new RotateEventHandler();
                _rotateEvent = ExternalEvent.Create(_rotateHandler);
                Logger.Info("✅ RotateEvent created");
            }
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            Logger.StartOperation("SetupDockablePane");
            
            data.FrameworkElement = this;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Tabbed,
                MinimumWidth = 300,
                MinimumHeight = 400
            };
            
            Logger.Info("DockablePane configured");
            Logger.EndOperation("SetupDockablePane");
        }

        private void BtnRunChangeLength_Click(object sender, RoutedEventArgs e)
        {
            Logger.StartOperation("DockPanel.ChangeLength");
            
            try
            {
                // ============================================================
                // VALIDATION 1: Kiểm tra UIApplication - TRY EVERYTHING
                // ============================================================
                // VALIDATION 1: Kiểm tra UIApplication
                // ============================================================
                UIApplication uiapp = _staticUIApp ?? _uiapp ?? RevitContext.UIApplication;
                
                Logger.Info($"UIApp check - static={_staticUIApp != null}, instance={_uiapp != null}, context={RevitContext.IsInitialized}");
                
                if (uiapp == null)
                {
                    Logger.Error("❌ UIApplication is NULL!");
                    txtChangeLengthStatus.Text = "❌ Cannot get UIApplication\nPanel opened incorrectly";
                    txtChangeLengthStatus.Foreground = System.Windows.Media.Brushes.Red;
                    
                    TaskDialog.Show("Critical Error", 
                        "Cannot get UIApplication!\n\n" +
                        "This panel may have been opened incorrectly.\n\n" +
                        "Solution:\n" +
                        "1. RESTART Revit 2023\n" +
                        "2. Click 'Show MEP Tools Panel' button from Ribbon\n" +
                        "3. Try again\n\n" +
                        "If issue persists, contact developer.");
                    return;
                }
                
                Logger.Info("✅ UIApplication OK");
                
                // ============================================================
                // VALIDATION 2: Kiểm tra Active Document
                // ============================================================
                if (uiapp.ActiveUIDocument == null)
                {
                    Logger.Error("No active Revit document");
                    txtChangeLengthStatus.Text = "✗ No active document";
                    txtChangeLengthStatus.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                // ============================================================
                // VALIDATION 3: Parse và validate input
                // ============================================================
                if (!double.TryParse(txtChangeLengthValue.Text, out double lengthMm))
                {
                    Logger.Error($"Invalid length: {txtChangeLengthValue.Text}");
                    txtChangeLengthStatus.Text = "✗ Invalid length value";
                    txtChangeLengthStatus.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                // Validate length range
                if (lengthMm < 1 || lengthMm > 100000)
                {
                    Logger.Error($"Length out of range: {lengthMm}mm");
                    txtChangeLengthStatus.Text = "✗ Length must be 1-100,000mm";
                    txtChangeLengthStatus.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                // ============================================================
                // ✨ DÙNG ExternalEvent - User sẽ PickPoint trong handler
                // ============================================================
                if (_changeLengthEvent == null || _changeLengthHandler == null)
                {
                    Logger.Error("ExternalEvent not initialized!");
                    txtChangeLengthStatus.Text = "✗ Event not initialized";
                    txtChangeLengthStatus.Foreground = System.Windows.Media.Brushes.Red;
                    TaskDialog.Show("Error", "ExternalEvent not initialized!\nPlease reopen the panel.");
                    return;
                }

                // Set data vào handler
                _changeLengthHandler.PendingLengthMm = lengthMm;
                
                // Raise ExternalEvent
                Logger.Info($"✅ Raising ChangeLengthEvent with {lengthMm}mm - user will PickPoint");
                _changeLengthEvent.Raise();
                
                txtChangeLengthStatus.Text = $"⏳ Click on pipe to select direction...";
                txtChangeLengthStatus.Foreground = System.Windows.Media.Brushes.Blue;
                
                Logger.EndOperation("DockPanel.ChangeLength");
            }
            catch (Exception ex)
            {
                Logger.Error("DockPanel: Change Length failed", ex);
                txtChangeLengthStatus.Text = $"✗ Error: {ex.Message}";
                txtChangeLengthStatus.Foreground = System.Windows.Media.Brushes.Red;
            }
        }

        private void BtnRunRotate_Click(object sender, RoutedEventArgs e)
        {
            Logger.StartOperation("DockPanel.Rotate");
            
            try
            {
                UIApplication uiapp = _staticUIApp ?? _uiapp;
                
                if (uiapp == null)
                {
                    Logger.Error("UIApplication is null");
                    txtRotateStatus.Text = "✗ Panel not initialized";
                    txtRotateStatus.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }
                
                if (uiapp.ActiveUIDocument == null)
                {
                    Logger.Error("No active Revit document");
                    txtRotateStatus.Text = "✗ No active document";
                    txtRotateStatus.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                if (!double.TryParse(txtRotateAngleValue.Text, out double angleDegrees))
                {
                    Logger.Error($"Invalid angle: {txtRotateAngleValue.Text}");
                    txtRotateStatus.Text = "✗ Invalid angle value";
                    txtRotateStatus.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                var selectedIds = uiapp.ActiveUIDocument.Selection.GetElementIds();
                if (selectedIds.Count == 0)
                {
                    Logger.Warning("No selection");
                    txtRotateStatus.Text = "⚠ Please select elements first";
                    txtRotateStatus.Foreground = System.Windows.Media.Brushes.Orange;
                    return;
                }

                // ============================================================
                // ✨ DÙNG ExternalEvent cho Rotate
                // ============================================================
                if (_rotateEvent == null || _rotateHandler == null)
                {
                    Logger.Error("RotateEvent not initialized!");
                    txtRotateStatus.Text = "✗ Event not initialized";
                    txtRotateStatus.Foreground = System.Windows.Media.Brushes.Red;
                    TaskDialog.Show("Error", "RotateEvent not initialized!\nPlease reopen the panel.");
                    return;
                }

                // Set data vào handler
                _rotateHandler.PendingAngleDegrees = angleDegrees;
                
                // Raise ExternalEvent
                Logger.Info($"✅ Raising RotateEvent with {angleDegrees}° for {selectedIds.Count} element(s)");
                _rotateEvent.Raise();
                
                txtRotateStatus.Text = $"⏳ Rotating {selectedIds.Count} element(s)...";
                txtRotateStatus.Foreground = System.Windows.Media.Brushes.Blue;
                
                Logger.EndOperation("DockPanel.Rotate");
            }
            catch (Exception ex)
            {
                Logger.Error("DockPanel: Change Length failed", ex);
                txtChangeLengthStatus.Text = $"✗ Error: {ex.Message}";
                txtChangeLengthStatus.Foreground = System.Windows.Media.Brushes.Red;
            }
        }
    }
}

