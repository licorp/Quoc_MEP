using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Nice3point.Revit.Extensions;
using Quoc_MEP.Export.Models;
using Quoc_MEP.Export.Managers;
using Quoc_MEP.Export.Commands;
using Quoc_MEP.Export.Utils;
using MessageBox = System.Windows.MessageBox;
using ComboBox = System.Windows.Controls.ComboBox;
using TextBox = System.Windows.Controls.TextBox;
using FolderBrowserDialog = System.Windows.Forms.FolderBrowserDialog;
using WpfGrid = System.Windows.Controls.Grid;
using WpfColor = System.Windows.Media.Color;
using SelectionChangedEventArgs = System.Windows.Controls.SelectionChangedEventArgs;
using FastWpfGrid;
using Button = System.Windows.Controls.Button;
using TextBlock = System.Windows.Controls.TextBlock;
using MessageBoxButton = System.Windows.MessageBoxButton;
using MessageBoxResult = System.Windows.MessageBoxResult;
using Wpf.Ui.Controls;

namespace Quoc_MEP.Export.Views
{
    /// <summary>
    /// Export + - Professional Style Interface with WPF-UI
    /// </summary>
    public partial class ExportPlusMainWindow : FluentWindow, INotifyPropertyChanged
    {
        private readonly Document _document;
        private readonly UIApplication _uiApp;
        private ObservableRangeCollection<SheetItem> _sheets;
        private ProfileManagerService _profileManager;
        private Models.Profile _selectedProfile;
        private ExternalEvent _exportEvent;
        private ExportHandler _exportHandler;
        
        // PDF Export External Event
        private ExternalEvent _pdfExportEvent;
        private Events.PDFExportEventHandler _pdfExportHandler;
        
        // IFC Export External Event
        private ExternalEvent _ifcExportEvent;
        private Commands.IFCExportHandler _ifcExportHandler;
        
        // View/Sheet Set Manager
        private ViewSheetSetManager _viewSheetSetManager;
        private ObservableCollection<ViewSheetSetInfo> _viewSheetSets;
        
        // Cancellation token for export operations
        private System.Threading.CancellationTokenSource _exportCancellationTokenSource;
        
        // Flag để tracking export completion - để reset khi user chọn lại sheet/view
        private bool _exportJustCompleted = false;
        
        // Flag to prevent infinite loop when bulk updating checkboxes
        private bool _isBulkUpdatingCheckboxes = false;
        
        // Flag to indicate window is closing - used to stop LoadSheets/LoadViews early
        private volatile bool _isClosing = false;
        
        // ⚡ Lazy loading flags - only load sheets/views when user actually needs them
        private bool _sheetsLoaded = false;
        private bool _viewsLoaded = false;
        private bool _fastGridInitialized = false;  // ← Prevent double init
        private bool _windowFullyLoaded = false;    // ← Track if Window_Loaded has completed

        // 📝 Interaction tracking để log phản hồi UI trong lúc đang load sheets
        private bool _userInteractionLoggedDuringSheetLoad = false;
        private int _userInteractionEventCount = 0;
        
        // ⚡ Debounce timer for UpdateExportSummary - batch multiple PropertyChanged events
        private DispatcherTimer _summaryUpdateTimer;
        
        // ⏱️ CONTINUOUS UI MONITORING: Log every 5 seconds to detect freeze
        private DispatcherTimer _uiMonitorTimer;
        private DateTime _formShownTime;
        private bool _isFormFullyShown = false;
        
        private const int USER_INTERACTION_LOG_LIMIT = 5;
        
        // ⏱️ PERFORMANCE TRACKING: Total time from constructor start to complete load
        private System.Diagnostics.Stopwatch _totalLoadTimer;
        
        // Performance optimization constants
        private const int BATCH_SIZE = 50; // Load 50 items per batch for UI updates
        private const int PARALLEL_THRESHOLD = 100; // Use parallel processing for 100+ items
        
        // Cache for sheet sizes to avoid repeated API calls
        private Dictionary<ElementId, string> _sheetSizeCache = new Dictionary<ElementId, string>();
        
        public ObservableCollection<ViewSheetSetInfo> ViewSheetSets
        {
            get => _viewSheetSets;
            set
            {
                _viewSheetSets = value;
                OnPropertyChanged(nameof(ViewSheetSets));
                OnPropertyChanged(nameof(SelectedSetsDisplay));
            }
        }
        
        /// <summary>
        /// Display text for selected sets in ToggleButton
        /// </summary>
        public string SelectedSetsDisplay
        {
            get
            {
                if (_viewSheetSets == null || !_viewSheetSets.Any(s => s.IsSelected))
                    return "All V/S Sets";
                
                var selectedNames = _viewSheetSets
                    .Where(s => s.IsSelected)
                    .Select(s => s.Name)
                    .ToList();
                    
                if (selectedNames.Count == 1)
                    return selectedNames[0];
                else if (selectedNames.Count <= 3)
                    return string.Join(", ", selectedNames);
                else
                    return $"{selectedNames.Count} sets selected";
            }
        }

        // Enhanced properties for data binding
        public int SelectedSheetsCount 
        { 
            get 
            { 
                return Sheets?.Count(s => s.IsSelected) ?? 0; 
            } 
        }

        public int SelectedViewsCount 
        { 
            get 
            { 
                return Views?.Count(v => v.IsSelected) ?? 0; 
            } 
        }
        
        public ObservableRangeCollection<SheetItem> SheetItems => Sheets;
        
        // New property for XAML binding
        public ObservableRangeCollection<SheetItem> Sheets 
        { 
            get => _sheets; 
            set 
            {
                _sheets = value;
                OnPropertyChanged(nameof(Sheets));
                OnPropertyChanged(nameof(SelectedSheetsCount));
            }
        }

        private ObservableCollection<ViewItem> _views;
        public ObservableCollection<ViewItem> Views 
        { 
            get => _views; 
            set 
            {
                _views = value;
                OnPropertyChanged(nameof(Views));
                OnPropertyChanged(nameof(SelectedViewsCount));
            }
        }

        // Properties for Create tab
        private string _outputFolder;
        public string OutputFolder
        {
            get => _outputFolder;
            set
            {
                _outputFolder = value;
                OnPropertyChanged(nameof(OutputFolder));
            }
        }

        public ObservableCollection<object> SelectedItemsForExport
        {
            get
            {
                var selectedItems = new ObservableCollection<object>();
                
                // Add selected sheets
                if (Sheets != null)
                {
                    foreach (var sheet in Sheets.Where(s => s.IsSelected))
                    {
                        selectedItems.Add(new
                        {
                            Number = sheet.SheetNumber,
                            Name = sheet.SheetName,
                            CustomFileName = sheet.CustomFileName,
                            Type = "Sheet"
                        });
                    }
                }
                
                // Add selected views
                if (Views != null)
                {
                    foreach (var view in Views.Where(v => v.IsSelected))
                    {
                        selectedItems.Add(new
                        {
                            Number = view.ViewType,
                            Name = view.ViewName,
                            CustomFileName = view.CustomFileName,
                            Type = "View"
                        });
                    }
                }
                
                return selectedItems;
            }
        }

        // INotifyPropertyChanged implementation
        public event PropertyChangedEventHandler PropertyChanged;
        
        protected virtual void OnPropertyChanged(string propertyName)
        {
            WriteDebugLog($"Property '{propertyName}' changed - firing PropertyChanged event");
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // Export settings với data binding
        public ExportSettings ExportSettings { get; set; }
        
        // Navisworks export settings
        private NWCExportSettings _nwcSettings = new NWCExportSettings();
        public NWCExportSettings NWCSettings
        {
            get => _nwcSettings;
            set
            {
                _nwcSettings = value;
                OnPropertyChanged(nameof(NWCSettings));
            }
        }
        
        // IFC export settings
        private IFCExportSettings _ifcSettings = new IFCExportSettings();
        public IFCExportSettings IFCSettings
        {
            get => _ifcSettings;
            set
            {
                _ifcSettings = value;
                OnPropertyChanged(nameof(IFCSettings));
            }
        }
        
        // IFC Setup Profiles Collection
        private ObservableCollection<string> _ifcCurrentSetups;
        public ObservableCollection<string> IFCCurrentSetups
        {
            get => _ifcCurrentSetups;
            set
            {
                _ifcCurrentSetups = value;
                OnPropertyChanged(nameof(IFCCurrentSetups));
            }
        }
        
        // Selected IFC Setup
        private string _selectedIFCSetup = "<In-Session Setup>";
        public string SelectedIFCSetup
        {
            get => _selectedIFCSetup;
            set
            {
                if (_selectedIFCSetup != value)
                {
                    _selectedIFCSetup = value;
                    OnPropertyChanged(nameof(SelectedIFCSetup));
                    
                    // Auto-load IFC setup from Revit when user selects
                    if (value != "<In-Session Setup>")
                    {
                        try
                        {
                            WriteDebugLog($"Loading IFC setup from Revit: {value}");
                            var loadedSettings = IFCExportManager.LoadIFCSetupFromRevit(_document, value);
                            
                            // Apply loaded settings to current IFCSettings
                            IFCSettings = loadedSettings;
                            
                            WriteDebugLog($"✓ IFC setup '{value}' loaded successfully");
                            System.Windows.MessageBox.Show(
                                $"IFC setup '{value}' loaded from Revit successfully!",
                                "Setup Loaded",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Information);
                        }
                        catch (Exception ex)
                        {
                            WriteDebugLog($"✗ Error loading IFC setup: {ex.Message}");
                            System.Windows.MessageBox.Show(
                                $"Could not load setup '{value}' from Revit.\n\nError: {ex.Message}",
                                "Setup Load Error",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Warning);
                        }
                    }
                }
            }
        }
        
        // IFC Setup Configuration Paths (mapping setup name to file path)
        private Dictionary<string, string> _ifcSetupConfigPaths;
        
        // Export Queue Items for Create tab
        private ObservableCollection<ExportQueueItem> _exportQueueItems;
        public ObservableCollection<ExportQueueItem> ExportQueueItems
        {
            get => _exportQueueItems;
            set
            {
                _exportQueueItems = value;
                OnPropertyChanged(nameof(ExportQueueItems));
            }
        }
        
        
        public ExportPlusMainWindow(Document document) : this(document, null)
        {
        }

        public ExportPlusMainWindow(Document document, UIApplication uiApp)
        {
            // ⏱️ START TOTAL LOAD TIMER
            _totalLoadTimer = System.Diagnostics.Stopwatch.StartNew();
            
            // ⚡ FIX: Only log once to avoid 5x duplication
            WriteDebugLog("===== EXPORT + CONSTRUCTOR STARTED =====");
            WriteDebugLog($"Document: {document?.Title ?? "NULL"}");
            WriteDebugLog($"⏱️  Total load timer started at {DateTime.Now:HH:mm:ss.fff}");
            
            _document = document;
            _uiApp = uiApp;
            
            // ✅ Initialize RevitAsyncHelper for safe async Revit API calls
            Quoc_MEP.Lib.RevitAsyncHelper.Initialize(_uiApp);
            WriteDebugLog("RevitAsyncHelper initialized for async operations");
            
            // Initialize External Event for export operations
            if (_uiApp != null)
            {
                _exportHandler = new ExportHandler();
                _exportEvent = ExternalEvent.Create(_exportHandler);
                WriteDebugLog("ExternalEvent initialized for export operations");
                
                // Initialize PDF Export External Event
                _pdfExportHandler = new Events.PDFExportEventHandler();
                _pdfExportEvent = ExternalEvent.Create(_pdfExportHandler);
                WriteDebugLog("PDF Export ExternalEvent initialized");
                
                // Initialize IFC Export External Event
                _ifcExportHandler = new Commands.IFCExportHandler();
                _ifcExportEvent = ExternalEvent.Create(_ifcExportHandler);
                WriteDebugLog("IFC Export ExternalEvent initialized");
            }
            
            // Initialize export settings with data binding
            ExportSettings = new ExportSettings();
            WriteDebugLog("ExportSettings initialized");
            
            // Initialize IFC Setup Profiles - Load from Revit dynamically
            WriteDebugLog("=== ATTEMPTING TO LOAD IFC SETUPS FROM REVIT ===");
            try
            {
                WriteDebugLog("Calling IFCExportManager.GetAvailableIFCSetups()...");
                var availableSetups = IFCExportManager.GetAvailableIFCSetups(_document);
                WriteDebugLog($"GetAvailableIFCSetups() returned {availableSetups.Count} setups");
                IFCCurrentSetups = new ObservableCollection<string>(availableSetups);
                SelectedIFCSetup = "<In-Session Setup>";
                WriteDebugLog($"✓ IFC Setup Profiles loaded from Revit: {availableSetups.Count} setups found");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"✗ ERROR loading IFC setups from Revit: {ex.Message}");
                WriteDebugLog($"   Stack Trace: {ex.StackTrace}");
                // Fallback to hardcoded list
                IFCCurrentSetups = new ObservableCollection<string>
                {
                    "<In-Session Setup>",
                    "IFC 2x3 Coordination View 2.0",
                    "IFC 2x3 Coordination View",
                    "IFC 2x3 GSA Concept Design BIM 2010",
                    "IFC 2x3 Basic FM Handover View",
                    "IFC 2x2 Coordination View",
                    "IFC 2x2 Singapore BCA e-Plan Check",
                    "IFC 2x3 COBie 2.4 Design Deliverable View",
                    "IFC4 Reference View [Architecture]",
                    "IFC4 Reference View [Structural]",
                    "IFC4 Reference View [BuildingService]",
                    "IFC4 Design Transfer View",
                    "Typical Setup"
                };
                SelectedIFCSetup = "<In-Session Setup>";
            }
            
            // Initialize output folder to Desktop
            OutputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            WriteDebugLog($"Default output folder set to: {OutputFolder}");
            
            // Initialize Export Queue with empty collection
            ExportQueueItems = new ObservableCollection<ExportQueueItem>();
            WriteDebugLog("ExportQueueItems initialized");
            
            InitializeComponent();
            WriteDebugLog("InitializeComponent completed");
            
            // Load DWG Export Setups from Revit
            // TODO: Temporarily disabled - will re-enable after fixing WPF compiler issues
            // LoadDWGExportSetups();
            // WriteDebugLog("DWG Export Setups loaded");
            
            // TODO: Wire up IFC Import/Export buttons after WPF build issues resolved
            // WireUpIFCButtons();
            
            // Configure window for non-modal operation
            ConfigureNonModalWindow();
            
            // Set DataContext for binding - should point to this window, not ExportSettings
            this.DataContext = this;
            WriteDebugLog("DataContext set to this window");
            
            InitializeProfiles();
            
            // ⚡⚡⚡ CRITICAL FIX: KHÔNG load sheets trong constructor!
            // Form phải hiện NGAY LẬP TỨC, load data SAU trong Loaded event
            WriteDebugLog("⚡ Sheets will be loaded AFTER form is shown (in Loaded event)");
            
            // Khởi tạo collection rỗng để DataGrid có thể bind
            Sheets = new ObservableRangeCollection<SheetItem>();
            
            _sheetsLoaded = false;
            _viewsLoaded = false;
            WriteDebugLog("⚡ Lazy loading enabled for both sheets and views");
            
            UpdateFormatSelection();
            UpdateNavigationButtons();
            
            // ⚡ Initialize View/Sheet Set Manager (fast - just create object)
            _viewSheetSetManager = new ViewSheetSetManager(_document);
            
            // ⚡ Initialize empty ViewSheetSets for binding
            ViewSheetSets = new ObservableCollection<ViewSheetSetInfo>();
            WriteDebugLog("ViewSheetSetManager created (sets will load in background)");
            
            AttachUserInteractionLoggingHandlers();
            WriteDebugLog("📝 Interaction logging attached to capture user actions during sheet load");
            
            // ⚡ Initialize UI monitoring timer - logs every 5 seconds to detect freeze
            _uiMonitorTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(5) // Log every 5 seconds
            };
            _uiMonitorTimer.Tick += UIMonitorTimer_Tick;
            WriteDebugLog("⏱️ UI Monitor timer initialized (will start after form shown)");
            
            // ⚡ Initialize debounce timer for UpdateExportSummary
            _summaryUpdateTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100) // 100ms delay
            };
            _summaryUpdateTimer.Tick += (s, e) =>
            {
                WriteDebugLog($"⏲️ DEBOUNCE TIMER FIRED at {DateTime.Now:HH:mm:ss.fff} - calling UpdateStatusText & UpdateExportSummary");
                _summaryUpdateTimer.Stop();
                
                var sw = System.Diagnostics.Stopwatch.StartNew();
                UpdateStatusText();
                sw.Stop();
                WriteDebugLog($"⏲️ UpdateStatusText took {sw.ElapsedMilliseconds}ms");
                
                sw.Restart();
                UpdateExportSummary();
                sw.Stop();
                WriteDebugLog($"⏲️ UpdateExportSummary took {sw.ElapsedMilliseconds}ms");
                
                WriteDebugLog($"⏲️ DEBOUNCE COMPLETE at {DateTime.Now:HH:mm:ss.fff}");
            };
            WriteDebugLog("⚡ Debounce timer initialized for UpdateExportSummary");

            WriteDebugLog("===== EXPORT + CONSTRUCTOR COMPLETED SUCCESSFULLY =====");
            
            // ⚡⚡⚡ CRITICAL: Unsubscribe trước khi subscribe để tránh duplicate
            this.Loaded -= ExportPlusMainWindow_Loaded;  // Remove nếu đã tồn tại
            this.Loaded += ExportPlusMainWindow_Loaded;  // Add mới
            
            // ⚡ CONTINUOUS USER INTERACTION MONITORING (not just during load!)
            AttachPermanentUserInteractionHandlers();
            WriteDebugLog("✅ Permanent user interaction monitoring attached");
        }
        
        /// <summary>
        /// Attach PERMANENT user interaction handlers to log all user actions
        /// </summary>
        private void AttachPermanentUserInteractionHandlers()
        {
            this.PreviewMouseDown += OnUserMouseDown;
            this.PreviewMouseWheel += OnUserMouseWheel;
            this.PreviewKeyDown += OnUserKeyDown;
            this.PreviewMouseMove += OnUserMouseMove;
            
            WriteDebugLog("🖱️ Permanent interaction handlers attached");
        }
        
        private void OnUserMouseDown(object sender, MouseButtonEventArgs e)
        {
            var control = e.OriginalSource?.GetType().Name ?? "Unknown";
            WriteDebugLog($"🖱️ USER CLICK: {e.ChangedButton} on {control} at {DateTime.Now:HH:mm:ss.fff}");
        }
        
        private void OnUserMouseWheel(object sender, MouseWheelEventArgs e)
        {
            var direction = e.Delta > 0 ? "UP" : "DOWN";
            WriteDebugLog($"🖱️ USER SCROLL: {direction} (delta={e.Delta}) at {DateTime.Now:HH:mm:ss.fff}");
            WriteDebugLog($"🖱️ SCROLL TARGET: {e.OriginalSource?.GetType().Name ?? "Unknown"}");
        }
        
        private void OnUserKeyDown(object sender, KeyEventArgs e)
        {
            WriteDebugLog($"⌨️ USER KEYPRESS: {e.Key} at {DateTime.Now:HH:mm:ss.fff}");
        }
        
        private DateTime _lastMouseMoveLog = DateTime.MinValue;
        private void OnUserMouseMove(object sender, MouseEventArgs e)
        {
            // Throttle mouse move logs (every 2 seconds)
            var now = DateTime.Now;
            if ((now - _lastMouseMoveLog).TotalSeconds >= 2)
            {
                _lastMouseMoveLog = now;
                var pos = e.GetPosition(this);
                WriteDebugLog($"🖱️ USER MOUSE MOVE: ({pos.X:F0}, {pos.Y:F0}) at {now:HH:mm:ss.fff}");
            }
        }

        private void AttachUserInteractionLoggingHandlers()
        {
            this.PreviewMouseDown -= OnPreviewMouseDownDuringLoad;
            this.PreviewMouseDown += OnPreviewMouseDownDuringLoad;

            this.PreviewMouseWheel -= OnPreviewMouseWheelDuringLoad;
            this.PreviewMouseWheel += OnPreviewMouseWheelDuringLoad;

            this.PreviewKeyDown -= OnPreviewKeyDownDuringLoad;
            this.PreviewKeyDown += OnPreviewKeyDownDuringLoad;

            this.TouchDown -= OnTouchDownDuringLoad;
            this.TouchDown += OnTouchDownDuringLoad;
        }

        private void DetachUserInteractionLoggingHandlers()
        {
            // ⚡ ONLY detach "during load" handlers, NOT permanent handlers!
            this.PreviewMouseDown -= OnPreviewMouseDownDuringLoad;
            this.PreviewMouseWheel -= OnPreviewMouseWheelDuringLoad;
            this.PreviewKeyDown -= OnPreviewKeyDownDuringLoad;
            this.TouchDown -= OnTouchDownDuringLoad;
            
            WriteDebugLog("🛑 Temporary load-time interaction handlers detached (permanent handlers still active)");
        }

        private void OnPreviewMouseDownDuringLoad(object sender, MouseButtonEventArgs e)
        {
            var button = e.ChangedButton.ToString();
            HandleUserInteractionDuringLoad($"MouseDown ({button})");
        }

        private void OnPreviewMouseWheelDuringLoad(object sender, MouseWheelEventArgs e)
        {
            var direction = e.Delta > 0 ? "WheelUp" : "WheelDown";
            HandleUserInteractionDuringLoad($"MouseWheel ({direction})");
        }

        private void OnPreviewKeyDownDuringLoad(object sender, KeyEventArgs e)
        {
            HandleUserInteractionDuringLoad($"KeyDown ({e.Key})");
        }

        private void OnTouchDownDuringLoad(object sender, TouchEventArgs e)
        {
            HandleUserInteractionDuringLoad("TouchDown");
        }

        private void HandleUserInteractionDuringLoad(string interactionSource)
        {
            if (_sheetsLoaded)
            {
                return;
            }

            _userInteractionEventCount++;

            if (!_userInteractionLoggedDuringSheetLoad)
            {
                _userInteractionLoggedDuringSheetLoad = true;
                WriteDebugLog("🙌 USER INPUT DETECTED while sheets are still loading - UI is responsive");
            }

            if (_userInteractionEventCount <= USER_INTERACTION_LOG_LIMIT)
            {
                WriteDebugLog($"🙋 Interaction #{_userInteractionEventCount}: {interactionSource} captured during sheet load");

                if (_userInteractionEventCount == USER_INTERACTION_LOG_LIMIT)
                {
                    WriteDebugLog("ℹ️ Interaction log limit reached while loading - further events will be suppressed");
                }
            }
        }

        private async Task BindSheetsInChunksAsync(IReadOnlyList<SheetItem> sheets)
        {
            const int CHUNK_SIZE = 20;

            if (sheets == null || sheets.Count == 0 || Sheets == null)
            {
                return;
            }

            int total = sheets.Count;
            int chunkIndex = 0;
            var chunkBuffer = new List<SheetItem>(CHUNK_SIZE);

            while (chunkIndex < total)
            {
                if (_cancelLoading)
                {
                    WriteDebugLog("⚠️ Binding cancelled mid-way - user closed window");
                    return;
                }

                chunkBuffer.Clear();
                int count = Math.Min(CHUNK_SIZE, total - chunkIndex);

                for (int i = 0; i < count; i++)
                {
                    chunkBuffer.Add(sheets[chunkIndex + i]);
                }

                Sheets.AddRange(chunkBuffer);
                chunkIndex += count;

                WriteDebugLog($"🧩 Bound chunk up to item {chunkIndex}/{total}");

                if (chunkIndex < total)
                {
                    await Dispatcher.Yield(DispatcherPriority.Background);
                }
            }
            
            // ⚡ CRITICAL: Subscribe PropertyChanged AFTER all chunks bound
            // BUT do it in BATCHES to avoid blocking UI thread
            WriteDebugLog($"⚡ Binding complete - subscribing PropertyChanged for {total} sheets...");
            
            int subscribedCount = 0;
            const int SUBSCRIBE_BATCH_SIZE = 20;
            
            for (int i = 0; i < Sheets.Count; i++)
            {
                Sheets[i].PropertyChanged += SheetItem_PropertyChanged;
                subscribedCount++;
                
                // Yield every 20 subscriptions to keep UI responsive
                if (subscribedCount % SUBSCRIBE_BATCH_SIZE == 0 && i < Sheets.Count - 1)
                {
                    await Dispatcher.Yield(DispatcherPriority.Background);
                    WriteDebugLog($"⚡ Subscribed {subscribedCount}/{total} handlers...");
                }
            }
            
            WriteDebugLog($"✅ PropertyChanged handlers subscribed (debounced) for {subscribedCount} sheets");
        }
        
        /// <summary>
        /// ⚡ Loaded event - Load ViewSheetSets và FastGrid trong background
        /// Form đã hiện → user có thể thao tác ngay
        /// </summary>
        private void ExportPlusMainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            WriteDebugLog($"===== WINDOW LOADED - _windowFullyLoaded={_windowFullyLoaded} =====");
            
            // ⚡⚡⚡ CRITICAL: GUARD - chỉ chạy 1 lần duy nhất!
            if (_windowFullyLoaded)
            {
                WriteDebugLog("⚠️ ExportPlusMainWindow_Loaded already executed - SKIPPING duplicate call");
                return;
            }
            
            // ⚡ SET FLAG: Window đã fully loaded, giờ mới cho phép loading data
            _windowFullyLoaded = true;
            WriteDebugLog("✅ _windowFullyLoaded = true - Form is ready for user interaction");
            
            // ⚡ Background loading ViewSheetSets + Progressive binding for Sheets
            System.Threading.Tasks.Task.Run(async () =>
            {
                try
                {
                    // Load ViewSheetSets from Revit (blocking Revit API call)
                    WriteDebugLog("📋 Loading View/Sheet Sets in background...");
                    var sets = _viewSheetSetManager?.GetAllViewSheetSets();
                    
                    // Switch to UI thread to update UI
                    await this.Dispatcher.InvokeAsync(() =>
                    {
                        if (sets != null && sets.Any())
                        {
                            ViewSheetSets = new ObservableCollection<ViewSheetSetInfo>(sets);
                            WriteDebugLog($"✅ Loaded {sets.Count} View/Sheet Sets");
                        }
                    });
                    
                    // ⏰ Show form first, let user interact BEFORE loading data
                    WriteDebugLog("⏰ Waiting 200ms for form to fully render...");
                    await System.Threading.Tasks.Task.Delay(200);
                    WriteDebugLog("✅ Form is VISIBLE - user can interact now");
                    WriteDebugLog("🔓 UI UNLOCKED: User can click buttons, type, scroll - form is responsive");
                    
                    // ⏱️ START UI MONITOR: Track UI state every 5 seconds
                    await this.Dispatcher.InvokeAsync(() =>
                    {
                        _formShownTime = DateTime.Now;
                        _isFormFullyShown = true;
                        _uiMonitorTimer.Start();
                        WriteDebugLog($"⏱️ UI MONITOR STARTED at {_formShownTime:HH:mm:ss.fff} - will log every 5 seconds");
                    });
                    
                    // ⚡⚡⚡ CRITICAL CHECK: Nếu sheets đã load rồi (do SelectionChanged event), SKIP!
                    if (_sheetsLoaded)
                    {
                        WriteDebugLog("⚠️ Sheets already loaded (from SelectionChanged event) - SKIPPING Window_Loaded loading");
                        return;
                    }
                    
                    // ===== PHASE 1: Load sheets data in background =====
                    WriteDebugLog("� [Phase 1] Loading sheets from Revit (background)...");
                    var totalTimer = System.Diagnostics.Stopwatch.StartNew();
                    
                    List<SheetItem> sheets = null;
                    await System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            var bgTimer = System.Diagnostics.Stopwatch.StartNew();
                            sheets = LoadSheetsInBackground();
                            bgTimer.Stop();
                            WriteDebugLog($"✅ [Phase 1] Loaded {sheets?.Count ?? 0} sheets in {bgTimer.ElapsedMilliseconds}ms");
                        }
                        catch (Exception ex)
                        {
                            WriteDebugLog($"❌ Background sheet loading error: {ex.Message}");
                        }
                    });
                    
                    if (sheets == null || sheets.Count == 0)
                    {
                        WriteDebugLog("⚠️ No sheets loaded from background");
                        return;
                    }
                    
                    // ===== PHASE 2: Bind to UI thread (NON-BLOCKING) =====
                    WriteDebugLog($"🔗 [Phase 2] Binding {sheets.Count} sheets in chunks...");

                    var bindingTimer = System.Diagnostics.Stopwatch.StartNew();

                    this.Dispatcher.InvokeAsync(async () =>
                    {
                        Sheets.Clear();
                        await BindSheetsInChunksAsync(sheets);
                        _sheetsLoaded = true;

                        DetachUserInteractionLoggingHandlers();

                        if (_userInteractionEventCount > 0)
                        {
                            WriteDebugLog($"🙌 Recorded {_userInteractionEventCount} user interaction(s) while sheets were loading");
                        }
                        else
                        {
                            WriteDebugLog("ℹ️ No user interaction captured during sheet load");
                        }

                        bindingTimer.Stop();
                        totalTimer.Stop();

                        WriteDebugLog($"✅ [Phase 2] Binding complete in {bindingTimer.ElapsedMilliseconds}ms ({sheets.Count} items)");
                        WriteDebugLog($"🎉 [COMPLETE] All loading done in {totalTimer.ElapsedMilliseconds}ms");
                        WriteDebugLog("🛑 Interaction logging detached after sheet load completion");
                        WriteDebugLog($"✅ UI was RESPONSIVE during entire load!");
                        WriteDebugLog($"⏱️ UI MONITOR TIMER: IsEnabled={_uiMonitorTimer?.IsEnabled ?? false} - will continue logging every 5s");
                    }, System.Windows.Threading.DispatcherPriority.Background);

                    // ⚡ Background thread returns IMMEDIATELY - UI thread processes binding when idle
                    WriteDebugLog("🚀 Background thread finished - UI thread will bind chunk by chunk when ready");
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"❌ Background loading error: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// Window loaded handler - FastGrid already initialized in background, no action needed
        /// </summary>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("===== WINDOW LOADED ===== (FastGrid already initialized)");
        }

        /// <summary>
        /// Initialize FastWpfGrid control for high-performance sheet list rendering.
        /// Falls back to DataGrid if FastWpfGrid types are not available.
        /// </summary>
        private void InitializeFastGrid()
        {
            // ⚡ CRITICAL: Prevent duplicate initialization
            if (_fastGridInitialized)
            {
                WriteDebugLog("⏭️ FastGrid already initialized - skipping duplicate call");
                return;
            }
            
            try
            {
                WriteDebugLog("✓ Initializing FastWpfGrid for sheet list rendering...");

                // ⚡ CRITICAL FIX: Ensure we use the CURRENT Sheets collection (not old reference)
                var currentSheets = Sheets; // Get current reference
                if (currentSheets == null || currentSheets.Count == 0)
                {
                    WriteDebugLog($"⚠️ WARNING: Sheets collection is empty ({currentSheets?.Count ?? 0} items) - FastGrid will show blank");
                    WriteDebugLog("⚠️ This means LoadSheets() hasn't run yet or Sheets was cleared.");
                    return;
                }
                
                // Create custom model for SheetItem binding
                var model = new SheetListGridModel(currentSheets);
                WriteDebugLog($"✓ SheetListGridModel created with {currentSheets?.Count ?? 0} sheets.");
                
                // Set model to pre-defined FastGrid control (from XAML - no InitializeComponent issues!)
                SheetsFastGrid.Model = model;
                WriteDebugLog("✓ FastGrid model set successfully.");
                
                // ⚡ Subscribe to FastGrid Loaded event to collapse DataGrid AFTER render
                SheetsFastGrid.Loaded += (s, e) =>
                {
                    WriteDebugLog("✓ FastGrid.Loaded fired - scheduling DataGrid collapse...");
                    
                    // ✅ Dùng ApplicationIdle để đảm bảo UI thread hoàn toàn rảnh (sau khi render xong)
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        WriteDebugLog("✓ UI thread idle - now hiding DataGrid");
                        SheetsDataGrid.Visibility = System.Windows.Visibility.Collapsed;
                    }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                };
                
                // Show FastGrid (defined in XAML, DataGrid vẫn visible để tránh flicker)
                SheetsFastGrid.Visibility = System.Windows.Visibility.Visible;
                
                _fastGridInitialized = true;  // ← Mark as initialized
                WriteDebugLog("✓✓✓ FastGrid initialized and will display after render.");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"? Exception initializing FastGrid: {ex.Message}");
                WriteDebugLog($"Stack: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Update ExportSettings from UI controls before export
        /// This ensures UI selections are properly synchronized with settings object
        /// </summary>
        private void UpdateExportSettingsFromUI()
        {
            try
            {
                WriteDebugLog("Updating ExportSettings from UI controls...");
                
                // Update Raster Quality from ComboBox
                if (RasterQualityCombo.SelectedItem is ComboBoxItem rasterItem)
                {
                    string rasterText = rasterItem.Content?.ToString() ?? "High";
                    WriteDebugLog($"UI Raster Quality: {rasterText}");
                    
                    switch (rasterText)
                    {
                        case "Low":
                            ExportSettings.RasterQuality = PSRasterQuality.Low;
                            WriteDebugLog("✓ RasterQuality set to LOW (72 DPI)");
                            break;
                        case "Medium":
                            ExportSettings.RasterQuality = PSRasterQuality.Medium;
                            WriteDebugLog("✓ RasterQuality set to MEDIUM (150 DPI)");
                            break;
                        case "High":
                            ExportSettings.RasterQuality = PSRasterQuality.High;
                            WriteDebugLog("✓ RasterQuality set to HIGH (300 DPI)");
                            break;
                        case "Presentation":
                            ExportSettings.RasterQuality = PSRasterQuality.Maximum;
                            WriteDebugLog("✓ RasterQuality set to PRESENTATION/MAXIMUM (600 DPI)");
                            break;
                        default:
                            ExportSettings.RasterQuality = PSRasterQuality.High;
                            WriteDebugLog("⚠ Unknown raster quality, defaulting to HIGH");
                            break;
                    }
                }
                
                // Update Colors from ComboBox
                if (ColorsCombo.SelectedItem is ComboBoxItem colorItem)
                {
                    string colorText = colorItem.Content?.ToString() ?? "Color";
                    WriteDebugLog($"UI Colors: {colorText}");
                    
                    switch (colorText)
                    {
                        case "Color":
                            ExportSettings.Colors = PSColors.Color;
                            WriteDebugLog("✓ Colors set to COLOR");
                            break;
                        case "Black and White":
                        case "Black and white":
                            ExportSettings.Colors = PSColors.BlackAndWhite;
                            WriteDebugLog("✓ Colors set to BLACK AND WHITE");
                            break;
                        case "Grayscale":
                            ExportSettings.Colors = PSColors.Grayscale;
                            WriteDebugLog("✓ Colors set to GRAYSCALE");
                            break;
                        default:
                            ExportSettings.Colors = PSColors.Color;
                            WriteDebugLog("⚠ Unknown color mode, defaulting to COLOR");
                            break;
                    }
                }
                
                // Update Output Folder
                if (!string.IsNullOrEmpty(CreateFolderPathTextBox?.Text))
                {
                    ExportSettings.OutputFolder = CreateFolderPathTextBox.Text;
                    WriteDebugLog($"✓ Output folder: {ExportSettings.OutputFolder}");
                }
                
                // Update Paper Placement settings
                if (CenterRadio?.IsChecked == true)
                {
                    ExportSettings.PaperPlacement = PSPaperPlacement.Center;
                    WriteDebugLog("✓ Paper Placement: CENTER");
                }
                else if (OffsetRadio?.IsChecked == true)
                {
                    ExportSettings.PaperPlacement = PSPaperPlacement.OffsetFromCorner;
                    WriteDebugLog("✓ Paper Placement: OFFSET FROM CORNER");
                }
                
                // Update Paper Margin
                if (MarginCombo.SelectedItem is ComboBoxItem marginItem)
                {
                    string marginText = marginItem.Content?.ToString() ?? "No Margin";
                    WriteDebugLog($"UI Margin: {marginText}");
                    
                    switch (marginText)
                    {
                        case "No Margin":
                            ExportSettings.PaperMargin = PSPaperMargin.NoMargin;
                            WriteDebugLog("✓ Paper Margin: NO MARGIN");
                            break;
                        case "Printer Limit":
                            ExportSettings.PaperMargin = PSPaperMargin.PrinterLimit;
                            WriteDebugLog("✓ Paper Margin: PRINTER LIMIT");
                            break;
                        case "User Defined":
                            ExportSettings.PaperMargin = PSPaperMargin.UserDefined;
                            WriteDebugLog("✓ Paper Margin: USER DEFINED");
                            break;
                        default:
                            ExportSettings.PaperMargin = PSPaperMargin.NoMargin;
                            WriteDebugLog("⚠ Unknown margin type, defaulting to NO MARGIN");
                            break;
                    }
                }
                
                // Update Offset X and Y values
                if (double.TryParse(OffsetXTextBox?.Text, out double offsetX))
                {
                    ExportSettings.OffsetX = offsetX;
                    WriteDebugLog($"✓ Offset X: {offsetX}");
                }
                
                if (double.TryParse(OffsetYTextBox?.Text, out double offsetY))
                {
                    ExportSettings.OffsetY = offsetY;
                    WriteDebugLog($"✓ Offset Y: {offsetY}");
                }
                
                // Update Combine Files setting
                if (CombineFilesRadio?.IsChecked == true)
                {
                    ExportSettings.CombineFiles = true;
                    WriteDebugLog("✓ File Mode: COMBINE multiple sheets into single file");
                }
                else if (SeparateFilesRadio?.IsChecked == true)
                {
                    ExportSettings.CombineFiles = false;
                    WriteDebugLog("✓ File Mode: CREATE SEPARATE files");
                }
                
                // Update Keep Paper Size & Orientation setting
                if (KeepPaperSizeCheckBox?.IsChecked == true)
                {
                    ExportSettings.KeepPaperSize = true;
                    WriteDebugLog("✓ Keep Paper Size & Orientation: ENABLED");
                }
                else
                {
                    ExportSettings.KeepPaperSize = false;
                    WriteDebugLog("✓ Keep Paper Size & Orientation: DISABLED");
                }
                
                WriteDebugLog("===== ExportSettings Updated Successfully =====");
                WriteDebugLog($"Final Settings: RasterQuality={ExportSettings.RasterQuality}, Colors={ExportSettings.Colors}");
                WriteDebugLog($"Paper Placement: {ExportSettings.PaperPlacement}, Margin: {ExportSettings.PaperMargin}, Offset: ({ExportSettings.OffsetX}, {ExportSettings.OffsetY})");
                WriteDebugLog($"Combine Files: {ExportSettings.CombineFiles}, Keep Paper Size: {ExportSettings.KeepPaperSize}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR updating ExportSettings from UI: {ex.Message}");
            }
        }

        /// <summary>
        /// ⏱️ UI MONITOR TIMER: Logs every 5 seconds to detect UI freeze patterns
        /// Tracks elapsed time since form shown and checks if user can interact
        /// </summary>
        private void UIMonitorTimer_Tick(object sender, EventArgs e)
        {
            if (!_isFormFullyShown) return;
            
            var elapsed = DateTime.Now - _formShownTime;
            var elapsedSeconds = (int)elapsed.TotalSeconds;
            
            // Log current state every 5 seconds
            WriteDebugLog($"⏱️ UI MONITOR [{elapsedSeconds}s]: " +
                         $"Sheets={Sheets?.Count ?? 0}, " +
                         $"Selected={Sheets?.Count(s => s.IsSelected) ?? 0}, " +
                         $"FormEnabled={this.IsEnabled}, " +
                         $"Focusable={this.Focusable}");
            
            // Test if UI is responsive by trying to access controls
            try
            {
                var dgEnabled = SheetsDataGrid?.IsEnabled ?? false;
                var dgItemCount = SheetsDataGrid?.Items?.Count ?? 0;
                
                WriteDebugLog($"⏱️ UI CHECK: DataGrid enabled={dgEnabled}, items={dgItemCount}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"⏱️ ❌ UI FREEZE DETECTED: {ex.Message}");
            }
        }

        /// <summary>
        /// ⚡ DEBOUNCED PropertyChanged handler - batches multiple updates into single UI refresh
        /// Prevents cascading UpdateExportSummary calls (89 sheets × UpdateExportSummary = freeze!)
        /// </summary>
        private void SheetItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsSelected")
            {
                WriteDebugLog($"🔔 PropertyChanged: IsSelected changed at {DateTime.Now:HH:mm:ss.fff}");
                
                // Restart timer on each PropertyChanged event
                _summaryUpdateTimer.Stop();
                _summaryUpdateTimer.Start();
                
                WriteDebugLog($"⏲️ Debounce timer restarted (will fire in 100ms)");
            }
        }
        
        private void ConfigureNonModalWindow()
        {
            try
            {
                // Configure window to work well as non-modal
                this.ShowInTaskbar = true;
                this.Topmost = false;
                this.WindowState = WindowState.Normal;
                
                // Handle window closing event
                this.Closing += ExportPlusMainWindow_Closing;
                
                // Handle window activated/deactivated for better UX
                this.Activated += ExportPlusMainWindow_Activated;
                this.Deactivated += ExportPlusMainWindow_Deactivated;
                
                WriteDebugLog("Non-modal window configuration completed");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error configuring non-modal window: {ex.Message}");
            }
        }

        private void ExportPlusMainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            WriteDebugLog("ExportPlus window is closing");
            DetachUserInteractionLoggingHandlers();
            WriteDebugLog("🛑 Interaction logging detached due to window closing");
            
            // ⏱️ Stop UI monitor timer
            if (_uiMonitorTimer != null && _uiMonitorTimer.IsEnabled)
            {
                _uiMonitorTimer.Stop();
                WriteDebugLog("⏱️ UI Monitor timer stopped");
            }
            
            // ⚡⚡⚡ CRITICAL: Set cancel flag NGAY LẬP TỨC để dừng loading
            _cancelLoading = true;
            _isLoadingSheets = false;
            WriteDebugLog("Set _cancelLoading flag - form will close immediately");
            
            // ⚡ KHÔNG chặn đóng form - để user tắt ngay
            // Không cần đợi loading xong
            
            // 🚀 OPTIMIZATION: Nếu e.Cancel = true (from cache logic), chỉ cleanup minimal
            // Nếu e.Cancel = false, đây là close thật → cleanup toàn bộ
            
            try
            {
                // Cancel any ongoing export operations first
                if (_exportCancellationTokenSource != null && !_exportCancellationTokenSource.IsCancellationRequested)
                {
                    WriteDebugLog("Cancelling ongoing export operations...");
                    try
                    {
                        _exportCancellationTokenSource.Cancel();
                    }
                    catch (Exception cancelEx)
                    {
                        WriteDebugLog($"Error cancelling export: {cancelEx.Message}");
                    }
                }
                
                // Give a brief moment for any pending operations to complete
                System.Threading.Thread.Sleep(100);
                
                // ⚠️ CHỈ dispose resources nếu THẬT SỰ đóng (không phải Hide)
                // Kiểm tra sau khi event handlers chạy xong
                if (!e.Cancel)
                {
                    WriteDebugLog("Real close detected - disposing all resources");
                    DisposeResources();
                }
                else
                {
                    WriteDebugLog("Window hidden (cached) - keeping resources alive");
                    // Reset closing flag để lần mở lại có thể load nếu cần
                    _isClosing = false;
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error during window cleanup: {ex.Message}");
                WriteDebugLog($"Stack trace: {ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// Dispose tất cả resources khi THẬT SỰ đóng window
        /// </summary>
        private void DisposeResources()
        {
            try
            {
                // Dispose CancellationTokenSource first
                if (_exportCancellationTokenSource != null)
                {
                    WriteDebugLog("Disposing CancellationTokenSource...");
                    try
                    {
                        _exportCancellationTokenSource.Dispose();
                        _exportCancellationTokenSource = null;
                        WriteDebugLog("CancellationTokenSource disposed");
                    }
                    catch (Exception disposeEx)
                    {
                        WriteDebugLog($"Error disposing CancellationTokenSource: {disposeEx.Message}");
                    }
                }
                
                // Dispose External Events
                if (_pdfExportEvent != null)
                {
                    WriteDebugLog("Disposing PDF Export Event...");
                    try
                    {
                        _pdfExportEvent.Dispose();
                        _pdfExportEvent = null;
                        WriteDebugLog("PDF Export Event disposed");
                    }
                    catch (Exception disposeEx)
                    {
                        WriteDebugLog($"Error disposing PDF Export Event: {disposeEx.Message}");
                    }
                }
                
                if (_exportEvent != null)
                {
                    WriteDebugLog("Disposing Export Event...");
                    try
                    {
                        _exportEvent.Dispose();
                        _exportEvent = null;
                        WriteDebugLog("Export Event disposed");
                    }
                    catch (Exception disposeEx)
                    {
                        WriteDebugLog($"Error disposing Export Event: {disposeEx.Message}");
                    }
                }
                
                if (_ifcExportEvent != null)
                {
                    WriteDebugLog("Disposing IFC Export Event...");
                    try
                    {
                        _ifcExportEvent.Dispose();
                        _ifcExportEvent = null;
                        WriteDebugLog("IFC Export Event disposed");
                    }
                    catch (Exception disposeEx)
                    {
                        WriteDebugLog($"Error disposing IFC Export Event: {disposeEx.Message}");
                    }
                }
                
                WriteDebugLog("ExportPlus window cleanup completed successfully");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error disposing resources: {ex.Message}");
            }
        }

        private void ExportPlusMainWindow_Activated(object sender, EventArgs e)
        {
            WriteDebugLog("ExportPlus window activated");
            // Window brought to front - could refresh data if needed
        }

        private void ExportPlusMainWindow_Deactivated(object sender, EventArgs e)
        {
            WriteDebugLog("ExportPlus window deactivated");
            // Window lost focus - user might be working in Revit
        }

        /// <summary>
        /// Show export completed dialog with Open Folder button
        /// </summary>
        private void ShowExportCompletedDialog(string folderPath)
        {
            try
            {
                var dialog = new ExportCompletedDialog(folderPath);
                dialog.Owner = this;
                dialog.ShowDialog();
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error showing export completed dialog: {ex.Message}");
                // Fallback to simple message box
                MessageBox.Show($"Export completed.\n\nLocation: {folderPath}", 
                              "Export Completed", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // DllImport for OutputDebugStringA to work with DebugView
        [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Ansi)]
        private static extern void OutputDebugStringA(string lpOutputString);

        /// <summary>
        /// Enhanced debug logging method compatible with DebugView
        /// </summary>
        private void WriteDebugLog(string message)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                string fullMessage = $"[Export +] {timestamp} - {message}";
                
                // ⚡ FIX: Only output ONCE to avoid 5x duplication in DebugView
                // DebugView captures OutputDebugString natively - don't need Debug.WriteLine etc
                OutputDebugStringA(fullMessage + "\r\n");
                
                // ❌ REMOVED: These ALL show up in DebugView causing 5x duplication!
                // System.Diagnostics.Debug.WriteLine(fullMessage);
                // Console.WriteLine(fullMessage);
                // System.Diagnostics.Trace.WriteLine(fullMessage);
            }
            catch (Exception ex)
            {
                // Fallback logging
                OutputDebugStringA($"[Export +] Logging error: {ex.Message}\r\n");
            }
        }

        // ⚡ ASYNC LOADING: Use RevitAsyncHelper instead of DispatcherTimer
        private const int SHEET_CHUNK_SIZE = 50; // ⚡ Tăng lên 50 để ít lần delay hơn
        private bool _isLoadingSheets = false;
        private bool _cancelLoading = false; // ⚡ Flag để cancel loading khi đóng form
        
        // ✅ Throttle scroll events (từ RevitScheduleEditor)
        private DateTime _lastScrollLoadTime = DateTime.MinValue;
        private const int SCROLL_LOAD_THROTTLE_MS = 200;
        
        /// <summary>
        /// Check xem sheet có chứa model views không (không phải schedule)
        /// </summary>
        private bool HasModelViews(ViewSheet sheet)
        {
            var placedViews = sheet.GetAllPlacedViews();
            
            // Sheet trống - BỎ QUA
            if (placedViews.Count == 0)
                return false;
            
            // Kiểm tra TẤT CẢ views trên sheet
            foreach (var viewId in placedViews)
            {
                var view = _document.GetElement(viewId) as View;
                if (view == null) continue;
                
                // Nếu có ít nhất 1 view KHÔNG phải schedule → GIỮ sheet
                if (!(view is ViewSchedule) && view.ViewType != ViewType.Schedule)
                {
                    return true; // Có model view
                }
            }
            
            return false; // Tất cả đều là schedule
        }
        
        /// <summary>
        /// ⚡ ASYNC VERSION: Load sheets without blocking UI
        /// Dùng cho user interaction (click tab, button...)
        /// </summary>
        private void LoadSheetsSync()
        {
            WriteDebugLog("⚡ LoadSheetsSync started...");
            
            var totalTimer = System.Diagnostics.Stopwatch.StartNew();
            
            // ⚡⚡⚡ CRITICAL: Revit API MUST run on MAIN UI thread
            // SYNCHRONOUS execution - NO async/await to avoid context issues!
            WriteDebugLog("⚡ Loading sheets SYNCHRONOUSLY on MAIN thread (Revit API requirement)...");
            List<SheetItem> loadedSheets = null;
            
            try
            {
                WriteDebugLog("⚡ Main thread: Loading sheets START...");
                var loadTimer = System.Diagnostics.Stopwatch.StartNew();
                
                // ✅ Call directly on main thread - NO async, NO Task.Run()!
                loadedSheets = LoadSheetsInBackground();
                
                loadTimer.Stop();
                WriteDebugLog($"⚡ Main thread: Loaded {loadedSheets?.Count ?? 0} sheets in {loadTimer.ElapsedMilliseconds}ms");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"✗ Sheet loading error: {ex.Message}");
                WriteDebugLog($"✗ Stack: {ex.StackTrace}");
            }
            
            WriteDebugLog("✓ Sheet loading DONE");
            
            // ⚡ Update UI (already on UI thread)
            WriteDebugLog("⚡ Updating UI with loaded sheets...");
            var uiUpdateTimer = System.Diagnostics.Stopwatch.StartNew();
            
            if (loadedSheets != null && loadedSheets.Count > 0)
            {
                WriteDebugLog($"⚡ Creating ObservableRangeCollection from {loadedSheets.Count} items...");
                Sheets = new ObservableRangeCollection<SheetItem>(loadedSheets);
                WriteDebugLog($"✓ UI updated with {Sheets.Count} sheets");
            }
            else
            {
                WriteDebugLog("⚠️ No sheets loaded");
            }
            
            WriteDebugLog("✓ Starting FastGrid initialization...");
            var fastGridTimer = System.Diagnostics.Stopwatch.StartNew();
            InitializeFastGrid();
            fastGridTimer.Stop();
            WriteDebugLog($"⏱️ FastGrid init took {fastGridTimer.ElapsedMilliseconds}ms");
            
            uiUpdateTimer.Stop();
            totalTimer.Stop();
            WriteDebugLog($"⏱️ TOTAL LoadSheetsAsync took {totalTimer.ElapsedMilliseconds}ms");
            
            // ✅ Mark as loaded
            _sheetsLoaded = true;
        }
        
        private void LoadSheets()
        {
            if (_cancelLoading)
            {
                WriteDebugLog("LoadSheets cancelled by user");
                return;
            }
            
            WriteDebugLog("LoadSheets started - SYNC mode (already in Revit context)");
            var startTime = System.Diagnostics.Stopwatch.StartNew();
            
            try
            {
                // ⚡ CHECK CANCEL TRƯỚC KHI BẮT ĐẦU
                if (_cancelLoading)
                {
                    WriteDebugLog("❌ LoadSheets CANCELLED before starting");
                    return;
                }
                
                // ⚡ STEP 1: Get ALL sheets and filter out schedule sheets
                WriteDebugLog("⚡ STEP 1: Getting model sheets (excluding Schedule sheets)...");
                var idsSw = System.Diagnostics.Stopwatch.StartNew();
                
                // ⚡ CHECK CANCEL
                if (_cancelLoading)
                {
                    WriteDebugLog("❌ LoadSheets CANCELLED during setup");
                    return;
                }
                
                // ⚡⚡⚡ CHỈ lấy ElementId (KHÔNG load ViewSheet object!)
                var collectorSw = System.Diagnostics.Stopwatch.StartNew();
                var allSheetIds = new FilteredElementCollector(_document)
                    .OfClass(typeof(ViewSheet))
                    .ToElementIds()  // ← CHỈ lấy ID, KHÔNG load object!
                    .ToList();
                collectorSw.Stop();
                WriteDebugLog($"⏱️ FilteredElementCollector (IDs only) took {collectorSw.ElapsedMilliseconds}ms");
                
                // ⚡ CHECK CANCEL NGAY SAU KHI COLLECT
                if (_cancelLoading)
                {
                    WriteDebugLog("❌ LoadSheets CANCELLED after collecting sheets");
                    return;
                }
                
                WriteDebugLog($"Total sheet IDs collected: {allSheetIds.Count}");
                
                // ⚡⚡⚡ Lọc bỏ Schedule sheets - CHỈ load object khi cần check
                var filterSw = System.Diagnostics.Stopwatch.StartNew();
                var sheetIds = new List<ElementId>();
                int skippedSchedules = 0;
                
                foreach (var sheetId in allSheetIds)
                {
                    // ⚡ CHECK CANCEL
                    if (_cancelLoading)
                    {
                        WriteDebugLog("❌ LoadSheets CANCELLED by user during filtering");
                        return;
                    }
                    
                    // Load object CHỈ để check (vẫn nhanh hơn load hết)
                    var sheet = _document.GetElement(sheetId) as ViewSheet;
                    if (sheet == null) continue;
                    
                    // Kiểm tra xem sheet có chứa model views không
                    if (!HasModelViews(sheet))
                    {
                        skippedSchedules++;
                        WriteDebugLog($"⚡ SKIP Schedule sheet: {sheet.SheetNumber} - {sheet.Name}");
                        continue;
                    }
                    
                    // Sheet có model views - lưu ID
                    sheetIds.Add(sheetId);
                }
                
                filterSw.Stop();
                WriteDebugLog($"⏱️ Schedule filtering took {filterSw.ElapsedMilliseconds}ms");
                
                idsSw.Stop();
                WriteDebugLog($"✅ Got {sheetIds.Count} model sheets (excluded {skippedSchedules} Schedule sheets) in {idsSw.ElapsedMilliseconds}ms");
                
                if (sheetIds.Count == 0)
                {
                    Sheets = new ObservableRangeCollection<SheetItem>();
                    WriteDebugLog("No sheets found");
                    return;
                }
                
                // ⚡⚡⚡ CRITICAL: Chỉ lấy SheetNumber + SheetName từ Parameter (NHANH!)
                WriteDebugLog($"⚡ STEP 2: Loading SheetNumber + SheetName from parameters...");
                var quickLoadSw = System.Diagnostics.Stopwatch.StartNew();
                
                // ⚡⚡⚡ Load vào List trước, sau đó tạo ObservableCollection MỘT LẦN
                var tempList = new List<SheetItem>(sheetIds.Count);
                
                int itemCount = 0;
                foreach (var sheetId in sheetIds)
                {
                    // ⚡ CHECK CANCEL
                    if (_cancelLoading)
                    {
                        WriteDebugLog("❌ LoadSheets CANCELLED by user during loading");
                        return;
                    }
                    
                    // ⚡ Lấy parameters TRỰC TIẾP (không cần ViewSheet object)
                    var element = _document.GetElement(sheetId);
                    var numberParam = element.get_Parameter(BuiltInParameter.SHEET_NUMBER);
                    var nameParam = element.get_Parameter(BuiltInParameter.SHEET_NAME);
                    
                    var number = numberParam?.AsString() ?? "";
                    var name = nameParam?.AsString() ?? "";
                    
                    tempList.Add(new SheetItem
                    {
                        Id = sheetId,
                        SheetNumber = number,
                        SheetName = name,
                        CustomFileName = $"{number} - {name}",
                        Size = "",
                        Revision = "",
                        IsSelected = false,
                        IsFullyLoaded = false
                    });
                    
                    itemCount++;
                }
                
                // ⚡⚡⚡ Tạo ObservableRangeCollection MỘT LẦN từ List (NHANH HƠN nhiều lần .Add())
                WriteDebugLog($"⚡ Creating ObservableRangeCollection from {tempList.Count} items...");
                Sheets = new ObservableRangeCollection<SheetItem>(tempList);
                
                quickLoadSw.Stop();
                WriteDebugLog($"⏱️ Loading {Sheets.Count} items took {quickLoadSw.ElapsedMilliseconds}ms");
                WriteDebugLog($"⏱️ Average per item: {(quickLoadSw.ElapsedMilliseconds / (double)Sheets.Count):F2}ms");
                
                startTime.Stop();
                WriteDebugLog($"🚀 DATA LOADED in {startTime.ElapsedMilliseconds}ms");
                WriteDebugLog($"⚡ Size/Revision will load on-demand if needed");
                WriteDebugLog($"✅ Form is now responsive - WPF will render DataGrid in background");
                
                _isLoadingSheets = false;
            }
            catch (Exception ex)
            {
                WriteDebugLog($"CRITICAL ERROR in LoadSheets: {ex.Message}");
                MessageBox.Show($"Error loading sheets: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                _isLoadingSheets = false;
            }
        }
        
        /// <summary>
        /// Load sheets in BACKGROUND THREAD - returns List (không update UI)
        /// </summary>
        private List<SheetItem> LoadSheetsInBackground()
        {
            WriteDebugLog($"⚡ LoadSheetsInBackground started - _sheetsLoaded={_sheetsLoaded}...");
            
            // ⚡⚡⚡ GUARD: Nếu đã load rồi, return empty
            if (_sheetsLoaded)
            {
                WriteDebugLog("⚠️ Sheets ALREADY LOADED - returning empty list");
                return new List<SheetItem>();
            }
            
            var startTime = System.Diagnostics.Stopwatch.StartNew();
            
            try
            {
                // ⚡ STEP 1: Get sheet IDs (NHANH!)
                WriteDebugLog("⚡ STEP 1: Collecting sheet IDs...");
                var collectorSw = System.Diagnostics.Stopwatch.StartNew();
                var allSheetIds = new FilteredElementCollector(_document)
                    .OfClass(typeof(ViewSheet))
                    .ToElementIds()
                    .ToList();
                collectorSw.Stop();
                WriteDebugLog($"⏱️ Collected {allSheetIds.Count} sheet IDs in {collectorSw.ElapsedMilliseconds}ms");
                
                // ⚡ STEP 2: Filter Schedule sheets
                WriteDebugLog("⚡ STEP 2: Filtering Schedule sheets...");
                var filterSw = System.Diagnostics.Stopwatch.StartNew();
                var sheetIds = new List<ElementId>();
                int skippedSchedules = 0;
                int processedCount = 0;
                
                foreach (var sheetId in allSheetIds)
                {
                    if (_cancelLoading) 
                    {
                        WriteDebugLog("⚠️ CANCELLED during filtering");
                        break;
                    }
                    
                    processedCount++;
                    if (processedCount % 20 == 0)
                    {
                        WriteDebugLog($"⚡ Filtering progress: {processedCount}/{allSheetIds.Count} sheets...");
                    }
                    
                    var sheet = _document.GetElement(sheetId) as ViewSheet;
                    if (sheet == null) continue;
                    
                    if (!HasModelViews(sheet))
                    {
                        skippedSchedules++;
                        continue;
                    }
                    
                    sheetIds.Add(sheetId);
                }
                
                filterSw.Stop();
                WriteDebugLog($"⏱️ Filtered {sheetIds.Count} model sheets (skipped {skippedSchedules}) in {filterSw.ElapsedMilliseconds}ms");
                
                if (sheetIds.Count == 0)
                {
                    WriteDebugLog("⚠️ No model sheets found");
                    return new List<SheetItem>();
                }
                
                // ⚡ STEP 3: Load parameters SEQUENTIALLY (Revit API is NOT thread-safe!)
                WriteDebugLog($"⚡ STEP 3: Loading parameters for {sheetIds.Count} sheets (SEQUENTIAL - Revit API restriction)...");
                var loadSw = System.Diagnostics.Stopwatch.StartNew();
                
                // ⚡⚡⚡ SEQUENTIAL PROCESSING: Revit API MUST run on main thread
                var tempList = new List<SheetItem>();
                int loadedCount = 0;
                
                WriteDebugLog($"🔄 Starting loop: {sheetIds.Count} sheets to process");
                
                foreach (var sheetId in sheetIds)
                {
                    if (_cancelLoading) break;
                    
                    try
                    {
                        // 🔍 DEBUG: Log EACH step
                        var sheet = _document.GetElement(sheetId) as ViewSheet;
                        if (sheet == null)
                        {
                            WriteDebugLog($"⚠️ Sheet {sheetId} is NULL - skipping");
                            continue;
                        }
                        
                        WriteDebugLog($"📄 Processing [{loadedCount + 1}/{sheetIds.Count}]: {sheet.SheetNumber} - {sheet.Name}");
                        
                        // ⚡ Load ALL parameters at once
                        WriteDebugLog($"   → Step 1: Loading parameters...");
                        var parameters = GetSheetParametersFast(sheet);
                        WriteDebugLog($"   → Step 2: Parameters loaded OK");
                        
                        // ⚡⚡⚡ SKIP GetCachedSheetSize() - TOO SLOW (8-20 seconds per sheet!)
                        // FilteredElementCollector in SheetSizeDetector causes massive delays
                        // Use fast pattern-based detection instead
                        string sheetSize = GuessSheetSizeFromNumber(sheet.SheetNumber);
                        WriteDebugLog($"   → Step 3: Size = {sheetSize} (guessed from number)");
                        
                        WriteDebugLog($"   → Step 4: Creating SheetItem...");
                        tempList.Add(new SheetItem
                        {
                            Id = sheetId,
                            SheetNumber = sheet.SheetNumber ?? "",
                            SheetName = sheet.Name ?? "",
                            CustomFileName = $"{sheet.SheetNumber} - {sheet.Name}",
                            Size = sheetSize,
                            Revision = parameters.Revision,
                            // ⚡ NEW: Extended parameters
                            DrawnBy = parameters.DrawnBy,
                            CheckedBy = parameters.CheckedBy,
                            ApprovedBy = parameters.ApprovedBy,
                            IssueDate = parameters.IssueDate,
                            DesignOption = parameters.DesignOption,
                            Phase = parameters.Phase,
                            IsSelected = false,
                            IsFullyLoaded = true // ⚡ All data loaded!
                        });
                        WriteDebugLog($"   ✅ Sheet {loadedCount + 1} added");
                        
                        loadedCount++;
                        if (loadedCount % 20 == 0)
                        {
                            WriteDebugLog($"⚡ Progress: {loadedCount}/{sheetIds.Count} sheets loaded");
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteDebugLog($"❌ CRITICAL ERROR loading sheet {sheetId}: {ex.Message}");
                        WriteDebugLog($"❌ Stack trace: {ex.StackTrace}");
                    }
                }
                
                loadSw.Stop();
                WriteDebugLog($"⏱️ Loaded {tempList.Count} sheets with FULL parameters in {loadSw.ElapsedMilliseconds}ms (SEQUENTIAL)");
                WriteDebugLog($"⏱️ Average: {(loadSw.ElapsedMilliseconds / (double)tempList.Count):F2}ms per sheet");
                
                startTime.Stop();
                WriteDebugLog($"🚀 LoadSheetsInBackground DONE in {startTime.ElapsedMilliseconds}ms");
                
                return tempList;
            }
            catch (Exception ex)
            {
                WriteDebugLog($"✗ CRITICAL ERROR in LoadSheetsInBackground: {ex.Message}");
                return new List<SheetItem>();
            }
        }
        
        /// <summary>
        /// Load sheets ĐỒNG BỘ theo chunk - Nhanh vì không có overhead của async/await
        /// </summary>
        private void LoadSheetsSync(List<ElementId> sheetIds)
        {
            try
            {
                int loadedCount = 0;
                
                // Load từng chunk ĐỒNG BỘ - Nhanh!
                for (int i = 0; i < sheetIds.Count; i += SHEET_CHUNK_SIZE)
                {
                    // ⚡ Check cancel flag trước mỗi chunk
                    if (_cancelLoading)
                    {
                        WriteDebugLog($"⚠️ Loading cancelled by user at {loadedCount}/{sheetIds.Count} sheets");
                        return;
                    }
                    
                    var endIndex = Math.Min(i + SHEET_CHUNK_SIZE, sheetIds.Count);
                    var chunkSw = System.Diagnostics.Stopwatch.StartNew();
                    
                    // ✅ Load chunk TRỰC TIẾP từ document
                    for (int j = i; j < endIndex; j++)
                    {
                        // ⚡ Check cancel flag trong loop
                        if (_cancelLoading)
                        {
                            WriteDebugLog($"⚠️ Loading cancelled by user at sheet {j}/{sheetIds.Count}");
                            return;
                        }
                        
                        try
                        {
                            var sheet = _document.GetElement(sheetIds[j]) as ViewSheet;
                            if (sheet != null && j < Sheets.Count)
                            {
                                var item = Sheets[j];
                                item.SheetNumber = sheet.SheetNumber ?? "";
                                item.SheetName = sheet.Name ?? "";
                                item.Revision = GetRevisionFast(sheet);
                                item.CustomFileName = $"{item.SheetNumber} - {item.SheetName}";
                                item.Size = ""; // Load on-demand
                                item.IsFullyLoaded = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteDebugLog($"Error loading sheet {j}: {ex.Message}");
                        }
                    }
                    
                    chunkSw.Stop();
                    loadedCount = endIndex;
                    
                    var progress = (loadedCount * 100) / sheetIds.Count;
                    WriteDebugLog($"   Progress: {loadedCount}/{sheetIds.Count} ({progress}%) - chunk took {chunkSw.ElapsedMilliseconds}ms");
                }
                
                // ⏱️ COMPLETE!
                if (_totalLoadTimer != null)
                {
                    _totalLoadTimer.Stop();
                    WriteDebugLog("╔════════════════════════════════════════════════════════════════╗");
                    WriteDebugLog("║           ✅ SYNC LOADING COMPLETE!                            ║");
                    WriteDebugLog("╚════════════════════════════════════════════════════════════════╝");
                    WriteDebugLog($"📊 FINAL STATISTICS:");
                    WriteDebugLog($"   • Total Sheets: {Sheets?.Count ?? 0}");
                    WriteDebugLog($"   • Total Time: {_totalLoadTimer.ElapsedMilliseconds}ms ({_totalLoadTimer.Elapsed.TotalSeconds:F2}s)");
                    WriteDebugLog($"   • Average/Sheet: {(Sheets?.Count > 0 ? _totalLoadTimer.ElapsedMilliseconds / (double)Sheets.Count : 0):F2}ms");
                    WriteDebugLog($"   • Chunk Size: {SHEET_CHUNK_SIZE} sheets");
                    WriteDebugLog("╔════════════════════════════════════════════════════════════════╗");
                    WriteDebugLog("║  🚀 SYNC pattern - No async overhead!                          ║");
                    WriteDebugLog("╚════════════════════════════════════════════════════════════════╝");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in LoadSheetsSync: {ex.Message}");
            }
        }
        
        // ⚡ TECHNIQUE 3: Load visible rows on scroll (ProSheets on-demand loading)
        private void SheetsDataGrid_ScrollChanged(object sender, System.Windows.Controls.ScrollChangedEventArgs e)
        {
            WriteDebugLog($"📜 SCROLL EVENT: VerticalChange={e.VerticalChange}, ExtentHeightChange={e.ExtentHeightChange} at {DateTime.Now:HH:mm:ss.fff}");
            
            // ✅ Throttle scroll events - chỉ load khi scroll dừng 200ms (giảm spam)
            if (e.VerticalChange != 0)
            {
                var now = DateTime.Now;
                if ((now - _lastScrollLoadTime).TotalMilliseconds >= SCROLL_LOAD_THROTTLE_MS)
                {
                    _lastScrollLoadTime = now;
                    WriteDebugLog($"📜 SCROLL THROTTLE PASSED - calling LoadVisibleSheetRows at {now:HH:mm:ss.fff}");
                    LoadVisibleSheetRows();
                    WriteDebugLog($"📜 LoadVisibleSheetRows RETURNED at {DateTime.Now:HH:mm:ss.fff}");
                }
                else
                {
                    WriteDebugLog($"📜 SCROLL THROTTLED - skipping LoadVisibleSheetRows (too soon)");
                }
            }
        }
        
        /// <summary>
        /// Handle ViewsDataGrid scroll to load visible view parameters on demand
        /// </summary>
        private void ViewsDataGrid_ScrollChanged(object sender, System.Windows.Controls.ScrollChangedEventArgs e)
        {
            WriteDebugLog($"👁️ VIEW SCROLL: VerticalChange={e.VerticalChange}, ExtentHeightChange={e.ExtentHeightChange} at {DateTime.Now:HH:mm:ss.fff}");
            
            // ✅ Throttle scroll events - only load when scroll stops for 200ms
            if (e.VerticalChange != 0)
            {
                var now = DateTime.Now;
                if ((now - _lastScrollLoadTime).TotalMilliseconds >= SCROLL_LOAD_THROTTLE_MS)
                {
                    _lastScrollLoadTime = now;
                    WriteDebugLog($"👁️ VIEW SCROLL THROTTLE PASSED - calling LoadVisibleViewRows");
                    LoadVisibleViewRows();
                }
                else
                {
                    WriteDebugLog($"👁️ VIEW SCROLL THROTTLED - skipping LoadVisibleViewRows (too soon)");
                }
            }
        }
        
        private void LoadVisibleSheetRows()
        {
            WriteDebugLog($"📜 LoadVisibleSheetRows STARTED at {DateTime.Now:HH:mm:ss.fff}");
            
            try
            {
                // Get ScrollViewer from DataGrid
                var scrollViewer = FindVisualChild<System.Windows.Controls.ScrollViewer>(SheetsDataGrid);
                if (scrollViewer == null)
                {
                    WriteDebugLog("📜 ScrollViewer not found - returning");
                    return;
                }
                
                int firstVisibleIndex = (int)scrollViewer.VerticalOffset;
                int visibleCount = (int)scrollViewer.ViewportHeight + 10; // +10 buffer
                int lastVisibleIndex = Math.Min(firstVisibleIndex + visibleCount, Sheets.Count);
                
                WriteDebugLog($"📜 Checking rows {firstVisibleIndex} to {lastVisibleIndex} (total {Sheets.Count})");
                
                // ⚡ NOTE: All sheets should already have Size loaded in background (IsFullyLoaded = true)
                // This method should rarely find items needing reload
                int itemsNeedingLoad = 0;
                for (int i = firstVisibleIndex; i < lastVisibleIndex; i++)
                {
                    if (i >= 0 && i < Sheets.Count)
                    {
                        var item = Sheets[i];
                        if (!item.IsFullyLoaded)
                        {
                            itemsNeedingLoad++;
                        }
                    }
                }
                
                if (itemsNeedingLoad > 0)
                {
                    WriteDebugLog($"⚠️ Found {itemsNeedingLoad} items not fully loaded - this should NOT happen!");
                }
                else
                {
                    WriteDebugLog($"✅ All visible rows already loaded (IsFullyLoaded = true)");
                }
                
                WriteDebugLog($"📜 LoadVisibleSheetRows COMPLETED at {DateTime.Now:HH:mm:ss.fff}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"❌ Error in LoadVisibleSheetRows: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Load view parameters (Scale, DetailLevel, Discipline) for visible rows ONLY
        /// Called when Views DataGrid is scrolled
        /// </summary>
        private void LoadVisibleViewRows()
        {
            WriteDebugLog($"👁️ LoadVisibleViewRows STARTED at {DateTime.Now:HH:mm:ss.fff}");
            
            try
            {
                if (Views == null || Views.Count == 0)
                {
                    WriteDebugLog("👁️ No views to load - returning");
                    return;
                }
                
                // Get ScrollViewer from DataGrid
                var scrollViewer = FindVisualChild<System.Windows.Controls.ScrollViewer>(ViewsDataGrid);
                if (scrollViewer == null)
                {
                    WriteDebugLog("👁️ ScrollViewer not found - returning");
                    return;
                }
                
                // Calculate visible row range
                // VerticalOffset is in rows (virtualized), ViewportHeight is in pixels
                int firstVisibleIndex = (int)scrollViewer.VerticalOffset;
                
                // Estimate rows visible: ViewportHeight (pixels) / estimated row height (35 pixels per row)
                const int ESTIMATED_ROW_HEIGHT_PIXELS = 35;
                int visibleRowCount = (int)(scrollViewer.ViewportHeight / ESTIMATED_ROW_HEIGHT_PIXELS) + 5; // +5 buffer
                int lastVisibleIndex = Math.Min(firstVisibleIndex + visibleRowCount, Views.Count);
                
                WriteDebugLog($"👁️ ScrollViewer: Offset={scrollViewer.VerticalOffset}, ViewportHeight={scrollViewer.ViewportHeight}px");
                WriteDebugLog($"👁️ Checking rows {firstVisibleIndex} to {lastVisibleIndex} (estimated {visibleRowCount} visible, total {Views.Count})");
                
                // Load view details for visible rows
                int itemsLoaded = 0;
                var loadedItems = new System.Collections.Generic.List<ViewItem>();
                
                for (int i = firstVisibleIndex; i < lastVisibleIndex; i++)
                {
                    if (i >= 0 && i < Views.Count)
                    {
                        var view = Views[i];
                        if (!view.IsFullyLoaded)
                        {
                            // 🆕 CRITICAL: Pass Document to LoadFullDetails() for lazy View retrieval
                            view.LoadFullDetails(_document);
                            loadedItems.Add(view);
                            itemsLoaded++;
                        }
                    }
                }
                
                // 🆕 CRITICAL: Force UI refresh for loaded items
                if (itemsLoaded > 0)
                {
                    WriteDebugLog($"✅ Loaded details for {itemsLoaded} views - refreshing UI...");
                    
                    // Force DataGrid to refresh by triggering PropertyChanged on UI thread
                    Dispatcher.Invoke(() =>
                    {
                        foreach (var item in loadedItems)
                        {
                            // Force UI to re-read the properties
                            item.RefreshUI();
                        }
                        WriteDebugLog($"✅ UI refresh completed for {itemsLoaded} items");
                    }, System.Windows.Threading.DispatcherPriority.Render);
                }
                else
                {
                    WriteDebugLog($"✅ All visible rows already loaded");
                }
                
                WriteDebugLog($"👁️ LoadVisibleViewRows COMPLETED at {DateTime.Now:HH:mm:ss.fff}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"❌ Error in LoadVisibleViewRows: {ex.Message}");
            }
        }
        
        // Helper to find ScrollViewer in visual tree
        private T FindVisualChild<T>(System.Windows.DependencyObject parent) where T : System.Windows.DependencyObject
        {
            if (parent == null) return null;
            
            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                {
                    return typedChild;
                }
                
                var result = FindVisualChild<T>(child);
                if (result != null)
                {
                    return result;
                }
            }
            
            return null;
        }
        
        /// <summary>
        /// ⚡ OPTIMIZED: Load sheets from ElementIds (sequential)
        /// Only loads Element when needed - no upfront graphics generation
        /// </summary>
        private void LoadSheetsSequentialFromIds(List<ElementId> sheetIds)
        {
            var newSheets = new List<SheetItem>();
            int processedCount = 0;
            
            foreach (var elementId in sheetIds)
            {
                if (_isClosing)
                {
                    WriteDebugLog($"⚠️ Window closing - stopped at {processedCount}/{sheetIds.Count}");
                    return;
                }
                
                try
                {
                    // ✅ LAZY LOAD: Only get element when processing
                    var sheet = _document.GetElement(elementId) as ViewSheet;
                    if (sheet == null || sheet.IsTemplate) continue;
                    
                    var sheetItem = ProcessSheetFast(sheet);
                    if (sheetItem != null)
                    {
                        newSheets.Add(sheetItem);
                        processedCount++;
                        
                        // Batch logging
                        if (processedCount % BATCH_SIZE == 0)
                        {
                            WriteDebugLog($"   Processed: {processedCount}/{sheetIds.Count}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"ERROR processing sheet {elementId}: {ex.Message}");
                }
            }
            
            FinalizeSheets(newSheets);
        }
        
        /// <summary>
        /// ⚡ OPTIMIZED: Load sheets from ElementIds (parallel)
        /// </summary>
        private void LoadSheetsParallelFromIds(List<ElementId> sheetIds)
        {
            WriteDebugLog($"⚡ Extracting sheet data from {sheetIds.Count} IDs...");
            
            // Step 1: Extract data in main thread (Revit API not thread-safe)
            var sheetDataList = new List<SheetDataFast>();
            int extractedCount = 0;
            
            foreach (var elementId in sheetIds)
            {
                if (_isClosing) return;
                
                try
                {
                    var sheet = _document.GetElement(elementId) as ViewSheet;
                    if (sheet == null || sheet.IsTemplate) continue;
                    
                    var data = new SheetDataFast
                    {
                        ElementId = elementId,
                        SheetNumber = sheet.SheetNumber ?? "NO_NUMBER",
                        SheetName = sheet.Name ?? "NO_NAME",
                        Revision = GetRevisionFast(sheet),
                        SheetSize = GetCachedSheetSize(sheet)
                    };
                    
                    sheetDataList.Add(data);
                    extractedCount++;
                    
                    if (extractedCount % BATCH_SIZE == 0)
                    {
                        WriteDebugLog($"   Extracted: {extractedCount}/{sheetIds.Count}");
                    }
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"ERROR extracting {elementId}: {ex.Message}");
                }
            }
            
            WriteDebugLog($"✅ Extracted {sheetDataList.Count} sheets");
            
            if (_isClosing) return;
            
            // Step 2: Process data in parallel (no Revit API calls)
            WriteDebugLog($"⚡ Processing {sheetDataList.Count} sheets in parallel...");
            var newSheets = new System.Collections.Concurrent.ConcurrentBag<SheetItem>();
            int processedCount = 0;
            
            System.Threading.Tasks.Parallel.ForEach(sheetDataList,
                new System.Threading.Tasks.ParallelOptions { MaxDegreeOfParallelism = 4 },
                (data) =>
                {
                    if (_isClosing) return;
                    
                    var sheetItem = CreateSheetItemFromDataFast(data);
                    if (sheetItem != null)
                    {
                        newSheets.Add(sheetItem);
                        
                        int count = System.Threading.Interlocked.Increment(ref processedCount);
                        if (count % BATCH_SIZE == 0)
                        {
                            WriteDebugLog($"   Parallel: {count}/{sheetDataList.Count}");
                        }
                    }
                });
            
            WriteDebugLog($"✅ Parallel processing done: {newSheets.Count} sheets");
            FinalizeSheets(newSheets.ToList());
        }
        
        // Helper class for fast data extraction
        private class SheetDataFast
        {
            public ElementId ElementId { get; set; }
            public string SheetNumber { get; set; }
            public string SheetName { get; set; }
            public string Revision { get; set; }
            public string SheetSize { get; set; }
            
            // ⚡ NEW: Extended parameters
            public string DrawnBy { get; set; }
            public string CheckedBy { get; set; }
            public string ApprovedBy { get; set; }
            public string IssueDate { get; set; }
            public string DesignOption { get; set; }
            public string Phase { get; set; }
        }
        
        /// <summary>
        /// ⚡ NEW: Load ALL sheet parameters at once (more efficient than 7 separate calls)
        /// </summary>
        private SheetParameters GetSheetParametersFast(ViewSheet sheet)
        {
            try
            {
                return new SheetParameters
                {
                    Revision = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION)?.AsString() ?? "",
                    DrawnBy = sheet.get_Parameter(BuiltInParameter.SHEET_DRAWN_BY)?.AsString() ?? "",
                    CheckedBy = sheet.get_Parameter(BuiltInParameter.SHEET_CHECKED_BY)?.AsString() ?? "",
                    ApprovedBy = sheet.get_Parameter(BuiltInParameter.SHEET_APPROVED_BY)?.AsString() ?? "",
                    IssueDate = sheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE)?.AsString() ?? "",
                    DesignOption = sheet.get_Parameter(BuiltInParameter.DESIGN_OPTION_ID)?.AsValueString() ?? "",
                    Phase = sheet.get_Parameter(BuiltInParameter.PHASE_CREATED)?.AsValueString() ?? ""
                };
            }
            catch
            {
                return new SheetParameters(); // Return empty if any error
            }
        }
        
        /// <summary>
        /// Helper class to hold all sheet parameters
        /// </summary>
        private class SheetParameters
        {
            public string Revision { get; set; } = "";
            public string DrawnBy { get; set; } = "";
            public string CheckedBy { get; set; } = "";
            public string ApprovedBy { get; set; } = "";
            public string IssueDate { get; set; } = "";
            public string DesignOption { get; set; } = "";
            public string Phase { get; set; } = "";
        }
        
        /// <summary>
        /// ⚡ SAFE FALLBACK: Guess sheet size from sheet number when detection fails
        /// </summary>
        private string GuessSheetSizeFromNumber(string sheetNumber)
        {
            if (string.IsNullOrEmpty(sheetNumber)) return "A1";
            
            // Common patterns: A0, A1, A2, A3, A4
            if (sheetNumber.Contains("A0")) return "A0";
            if (sheetNumber.Contains("A1")) return "A1";
            if (sheetNumber.Contains("A2")) return "A2";
            if (sheetNumber.Contains("A3")) return "A3";
            if (sheetNumber.Contains("A4")) return "A4";
            
            // Default to A1 (most common)
            return "A1";
        }
        
        /// <summary>
        /// ⚡ FAST: Get revision without exception handling overhead (DEPRECATED - use GetSheetParametersFast)
        /// </summary>
        private string GetRevisionFast(ViewSheet sheet)
        {
            try
            {
                Parameter revParam = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION);
                return revParam?.AsString() ?? "";
            }
            catch
            {
                return "";
            }
        }
        
        /// <summary>
        /// ⚡ FAST: Process sheet with minimal overhead
        /// </summary>
        private SheetItem ProcessSheetFast(ViewSheet sheet)
        {
            var sheetItem = new SheetItem
            {
                Id = sheet.Id,
                RevitSheet = sheet,
                IsSelected = false,
                SheetNumber = sheet.SheetNumber ?? "NO_NUMBER",
                SheetName = sheet.Name ?? "NO_NAME",
                Revision = GetRevisionFast(sheet),
                Size = GetCachedSheetSize(sheet),
                CustomFileName = $"{sheet.SheetNumber ?? "NO_NUMBER"}_{(sheet.Name ?? "NO_NAME").Replace(" ", "_")}",
                IsFullyLoaded = true
            };
            
            // ⚡ NO PropertyChanged here - will subscribe AFTER binding completes
            
            return sheetItem;
        }
        
        /// <summary>
        /// ⚡ FAST: Get sheet size from cache (no FilteredElementCollector call)
        /// </summary>
        private string GetCachedSheetSize(ViewSheet sheet)
        {
            return Utils.SheetSizeDetector.GetSheetSize(sheet);
        }
        
        /// <summary>
        /// ⚡ FAST: Create SheetItem from pre-extracted data (parallel-safe)
        /// </summary>
        private SheetItem CreateSheetItemFromDataFast(SheetDataFast data)
        {
            var sheetItem = new SheetItem
            {
                Id = data.ElementId,
                RevitSheet = null, // Will be set when needed
                IsSelected = false,
                SheetNumber = data.SheetNumber,
                SheetName = data.SheetName,
                Revision = data.Revision,
                Size = data.SheetSize,
                // ⚡ NEW: Extended parameters
                DrawnBy = data.DrawnBy,
                CheckedBy = data.CheckedBy,
                ApprovedBy = data.ApprovedBy,
                IssueDate = data.IssueDate,
                DesignOption = data.DesignOption,
                Phase = data.Phase,
                CustomFileName = $"{data.SheetNumber}_{data.SheetName.Replace(" ", "_")}",
                IsFullyLoaded = true // ⚡ Already loaded all data (Size, Revision, + 6 extended params) in background
            };
            
            // ⚡ NO PropertyChanged here - will subscribe AFTER binding completes
            
            return sheetItem;
        }
        
        /// <summary>
        /// Load sheets sequentially with batch UI updates - DEPRECATED
        /// Use LoadSheetsSequentialFromIds instead
        /// </summary>
        private void LoadSheetsSequential(List<ViewSheet> sheets)
        {
            var newSheets = new List<SheetItem>();
            int addedCount = 0;
            
            foreach (var sheet in sheets)
            {
                if (_isClosing)
                {
                    WriteDebugLog($"⚠️ Window closing detected - stopping at sheet #{addedCount}");
                    return;
                }
                
                var sheetItem = ProcessSheet(sheet);
                if (sheetItem != null)
                {
                    newSheets.Add(sheetItem);
                    addedCount++;
                    
                    // Batch update UI every BATCH_SIZE items
                    if (addedCount % BATCH_SIZE == 0)
                    {
                        WriteDebugLog($"📦 Batch update: {addedCount}/{sheets.Count} sheets loaded");
                    }
                }
            }
            
            FinalizeSheets(newSheets);
        }
        
        /// <summary>
        /// Load sheets in parallel for faster processing (100+ sheets)
        /// </summary>
        private void LoadSheetsParallel(List<ViewSheet> sheets)
        {
            // CRITICAL: Revit API is NOT thread-safe!
            // We MUST extract all data from Revit API in the main thread first
            // Then we can process that data in parallel threads
            
            WriteDebugLog($"⚡ Extracting data from {sheets.Count} sheets in main thread (Revit API is NOT thread-safe)...");
            
            // Step 1: Extract ALL data from Revit API in main thread (thread-safe)
            var sheetDataList = new List<SheetData>();
            int extractedCount = 0;
            
            foreach (var sheet in sheets)
            {
                if (_isClosing) return;
                
                try
                {
                    var data = new SheetData
                    {
                        Sheet = sheet,
                        SheetNumber = sheet.SheetNumber ?? "NO_NUMBER",
                        SheetName = sheet.Name ?? "NO_NAME",
                        SheetId = sheet.Id.IntegerValue.ToString(),
                        Revision = GetRevision(sheet),
                        SheetSize = GetCachedSheetSize(sheet)
                    };
                    
                    sheetDataList.Add(data);
                    extractedCount++;
                    
                    if (extractedCount % BATCH_SIZE == 0)
                    {
                        WriteDebugLog($"📊 Extracted: {extractedCount}/{sheets.Count} sheets");
                    }
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"ERROR extracting sheet data: {ex.Message}");
                }
            }
            
            WriteDebugLog($"✅ Data extraction complete: {sheetDataList.Count} sheets");
            
            if (_isClosing) return;
            
            // Step 2: Now process the extracted data in parallel (NO Revit API calls here!)
            WriteDebugLog($"⚡ Processing {sheetDataList.Count} sheets in parallel...");
            var newSheets = new System.Collections.Concurrent.ConcurrentBag<SheetItem>();
            int processedCount = 0;
            
            System.Threading.Tasks.Parallel.ForEach(sheetDataList, 
                new System.Threading.Tasks.ParallelOptions { MaxDegreeOfParallelism = 4 },
                (data) =>
                {
                    if (_isClosing) return;
                    
                    var sheetItem = CreateSheetItemFromData(data);
                    if (sheetItem != null)
                    {
                        newSheets.Add(sheetItem);
                        
                        int count = System.Threading.Interlocked.Increment(ref processedCount);
                        if (count % BATCH_SIZE == 0)
                        {
                            WriteDebugLog($"⚡ Parallel progress: {count}/{sheetDataList.Count} sheets processed");
                        }
                    }
                });
            
            if (_isClosing) return;
            
            WriteDebugLog($"⚡ Parallel processing complete: {newSheets.Count} sheets");
            FinalizeSheets(newSheets.ToList());
        }
        
        // Helper class to store sheet data extracted from Revit API
        private class SheetData
        {
            public ViewSheet Sheet { get; set; }
            public string SheetNumber { get; set; }
            public string SheetName { get; set; }
            public string SheetId { get; set; }
            public string Revision { get; set; }
            public string SheetSize { get; set; }
        }
        
        // Helper method to extract revision (main thread only)
        private string GetRevision(ViewSheet sheet)
        {
            try
            {
                Parameter revParam = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION);
                return revParam?.AsString() ?? "";
            }
            catch
            {
                return "";
            }
        }
        
        // Create SheetItem from pre-extracted data (can run in parallel thread)
        private SheetItem CreateSheetItemFromData(SheetData data)
        {
            try
            {
                var sheetItem = new SheetItem
                {
                    Id = data.Sheet.Id,  // ElementId, not string
                    RevitSheet = data.Sheet,
                    SheetNumber = data.SheetNumber,
                    SheetName = data.SheetName,
                    Revision = data.Revision,
                    Size = data.SheetSize,
                    IsSelected = false,
                    IsFullyLoaded = true,
                    CustomFileName = $"{data.SheetNumber} - {data.SheetName}"
                };
                
                // ⚡ NO PropertyChanged here - will subscribe AFTER binding completes
                
                return sheetItem;
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR creating sheet item: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Process a single sheet - optimized with caching
        /// </summary>
        private SheetItem ProcessSheet(ViewSheet sheet)
        {
            try
            {
                string sheetNumber = sheet.SheetNumber ?? "NO_NUMBER";
                string sheetName = sheet.Name ?? "NO_NAME";
                
                // Get revision (fast)
                string revision = "";
                try
                {
                    Parameter revParam = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION);
                    revision = revParam?.AsString() ?? "";
                }
                catch { }
                
                // Get size with caching
                string sheetSize = GetCachedSheetSize(sheet);
                
                var sheetItem = new SheetItem
                {
                    Id = sheet.Id,
                    RevitSheet = sheet,
                    IsSelected = false,
                    SheetNumber = sheetNumber,
                    SheetName = sheetName,
                    Revision = revision,
                    Size = sheetSize,
                    CustomFileName = $"{sheetNumber}_{sheetName.Replace(" ", "_")}",
                    IsFullyLoaded = true
                };
                
                // ⚡ NO PropertyChanged here - will subscribe AFTER binding completes
                
                return sheetItem;
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error processing sheet {sheet.SheetNumber}: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Finalize sheets collection - sort and update UI
        /// </summary>
        private async void FinalizeSheets(List<SheetItem> sheets)
        {
            WriteDebugLog($"Finalizing {sheets.Count} sheets - sorting...");
            
            // Sort sheets
            var sortedSheets = sheets.OrderBy(s => s.SheetNumber, new AlphanumericComparer()).ToList();
            
            // Update UI on dispatcher thread (NON-BLOCKING)
            await Dispatcher.InvokeAsync(async () =>
            {
                Sheets = new ObservableRangeCollection<SheetItem>(sortedSheets);
                
                WriteDebugLog($"⚡ NOW subscribing PropertyChanged for {sortedSheets.Count} sheets...");
                
                // ⚡ CRITICAL: Subscribe PropertyChanged in BATCHES to avoid blocking
                int subscribedCount = 0;
                const int SUBSCRIBE_BATCH_SIZE = 20;
                
                for (int i = 0; i < Sheets.Count; i++)
                {
                    Sheets[i].PropertyChanged += SheetItem_PropertyChanged;
                    subscribedCount++;
                    
                    // Yield every 20 subscriptions
                    if (subscribedCount % SUBSCRIBE_BATCH_SIZE == 0 && i < Sheets.Count - 1)
                    {
                        await Dispatcher.Yield(DispatcherPriority.Background);
                        WriteDebugLog($"⚡ Subscribed {subscribedCount}/{sortedSheets.Count} handlers...");
                    }
                }
                
                WriteDebugLog($"✅ PropertyChanged handlers subscribed (debounced) for {subscribedCount} sheets");
                
                UpdateStatusText();
                UpdateExportSummary();
                WriteDebugLog("✅ Sheets collection updated in UI");
            });
        }

        private void LoadViews()
        {
            WriteDebugLog("LoadViews started");
            var startTime = System.Diagnostics.Stopwatch.StartNew();
            
            try
            {
                // Store existing custom filenames
                var existingCustomNames = Views?.Where(v => !string.IsNullOrEmpty(v.ViewId))
                                                .ToDictionary(v => v.ViewId, v => v.CustomFileName) 
                                         ?? new Dictionary<string, string>();
                WriteDebugLog($"Preserved {existingCustomNames.Count} existing custom filenames");

                // ⚡ STEP 1: Get ElementIds only (FAST - no graphics generation!)
                var step1 = System.Diagnostics.Stopwatch.StartNew();
                
                var viewIds = new FilteredElementCollector(_document)
                    .OfCategory(BuiltInCategory.OST_Views) // ✅ Quick Filter (faster than OfClass)
                    .WhereElementIsNotElementType() // Nice3point extension
                    .ToElementIds() // ✅ No graphics generation!
                    .Where(id =>
                    {
                        var view = _document.GetElement(id) as View;
                        return view != null &&
                               !view.IsTemplate &&
                               view.ViewType != ViewType.DrawingSheet &&
                               view.ViewType != ViewType.ProjectBrowser &&
                               view.ViewType != ViewType.SystemBrowser &&
                               view.CanBePrinted;
                    })
                    .ToList();
                
                step1.Stop();
                int viewCount = viewIds.Count;
                WriteDebugLog($"✅ STEP 1: Got {viewCount} ElementIds in {step1.ElapsedMilliseconds}ms (no graphics!)");

                // Choose loading strategy based on count
                if (viewCount >= PARALLEL_THRESHOLD)
                {
                    WriteDebugLog($"⚡ PARALLEL LOADING: {viewCount} views (threshold: {PARALLEL_THRESHOLD})");
                    LoadViewsParallelFromIds(viewIds, existingCustomNames);
                }
                else
                {
                    WriteDebugLog($"⚡ SEQUENTIAL LOADING: {viewCount} views");
                    LoadViewsSequentialFromIds(viewIds, existingCustomNames);
                }
                
                startTime.Stop();
                WriteDebugLog($"✅ LoadViews completed - Total: {Views?.Count ?? 0} views in {startTime.ElapsedMilliseconds}ms");
                WriteDebugLog($"⚡ Performance: {(Views?.Count > 0 ? startTime.ElapsedMilliseconds / (double)Views.Count : 0):F2}ms per view");
            }
            catch (Exception ex)
            {
                startTime.Stop();
                WriteDebugLog($"CRITICAL ERROR in LoadViews: {ex.Message}");
                WriteDebugLog($"StackTrace: {ex.StackTrace}");
                MessageBox.Show($"Error loading views: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// ⚡ OPTIMIZED: Load views from ElementIds (sequential)
        /// Following Schedule_Editor pattern: Load first 200 with full details, rest lazy load
        /// </summary>
        private void LoadViewsSequentialFromIds(List<ElementId> viewIds, Dictionary<string, string> existingCustomNames)
        {
            const int INITIAL_BATCH_SIZE = 200; // Load first 200 with full details (like Schedule_Editor)
            
            var newViews = new List<ViewItem>();
            int addedCount = 0;
            int restoredCount = 0;
            
            foreach (var elementId in viewIds)
            {
                if (_isClosing)
                {
                    WriteDebugLog($"⚠️ Window closing - stopped at {addedCount}/{viewIds.Count}");
                    return;
                }
                
                try
                {
                    // ✅ LAZY LOAD: Only get element when processing
                    var view = _document.GetElement(elementId) as View;
                    if (view == null) continue;
                    
                    // ✅ IMPROVED: First 200 items load full details, rest lazy load
                    bool shouldLoadFullDetails = (addedCount < INITIAL_BATCH_SIZE);
                    var viewItem = ProcessView(view, existingCustomNames, ref restoredCount, loadFullDetails: shouldLoadFullDetails);
                    
                    newViews.Add(viewItem);
                    addedCount++;
                    
                    if (addedCount % BATCH_SIZE == 0)
                    {
                        WriteDebugLog($"   Processed: {addedCount}/{viewIds.Count} (Full details: {Math.Min(addedCount, INITIAL_BATCH_SIZE)})");
                    }
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"ERROR processing view {elementId}: {ex.Message}");
                }
            }
            
            WriteDebugLog($"Restored custom filenames: {restoredCount}/{addedCount}");
            WriteDebugLog($"✅ Loaded full details for first {Math.Min(addedCount, INITIAL_BATCH_SIZE)} views");
            WriteDebugLog($"✅ Lazy loaded remaining {Math.Max(0, addedCount - INITIAL_BATCH_SIZE)} views");
            FinalizeViews(newViews);
        }
        
        /// <summary>
        /// ⚡ OPTIMIZED: Load views from ElementIds (parallel)
        /// </summary>
        private void LoadViewsParallelFromIds(List<ElementId> viewIds, Dictionary<string, string> existingCustomNames)
        {
            WriteDebugLog($"⚡ Extracting view data from {viewIds.Count} IDs...");
            
            // Step 1: Extract data in main thread
            var viewDataList = new List<ViewData>();
            int extractedCount = 0;
            
            foreach (var elementId in viewIds)
            {
                if (_isClosing) return;
                
                try
                {
                    var view = _document.GetElement(elementId) as View;
                    if (view == null) continue;
                    
                    var data = new ViewData
                    {
                        ElementId = elementId,
                        ViewId = view.Id.ToString(),
                        ViewName = view.Name ?? "NO_NAME",
                        ViewType = view.ViewType.ToString(),
                        ViewScale = view.Scale.ToString()
                    };
                    
                    viewDataList.Add(data);
                    extractedCount++;
                    
                    if (extractedCount % BATCH_SIZE == 0)
                    {
                        WriteDebugLog($"   Extracted: {extractedCount}/{viewIds.Count}");
                    }
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"ERROR extracting {elementId}: {ex.Message}");
                }
            }
            
            WriteDebugLog($"✅ Extracted {viewDataList.Count} views");
            
            if (_isClosing) return;
            
            // Step 2: Process in parallel
            WriteDebugLog($"⚡ Processing {viewDataList.Count} views in parallel...");
            var newViews = new System.Collections.Concurrent.ConcurrentBag<ViewItem>();
            int processedCount = 0;
            int restoredCount = 0;
            
            System.Threading.Tasks.Parallel.ForEach(viewDataList,
                new System.Threading.Tasks.ParallelOptions { MaxDegreeOfParallelism = 4 },
                (data) =>
                {
                    if (_isClosing) return;
                    
                    var viewItem = CreateViewItemFromData(data, existingCustomNames, ref restoredCount);
                    if (viewItem != null)
                    {
                        newViews.Add(viewItem);
                        
                        int count = System.Threading.Interlocked.Increment(ref processedCount);
                        if (count % BATCH_SIZE == 0)
                        {
                            WriteDebugLog($"   Parallel: {count}/{viewDataList.Count}");
                        }
                    }
                });
            
            WriteDebugLog($"✅ Parallel processing done: {newViews.Count} views");
            WriteDebugLog($"Restored custom filenames: {restoredCount}/{newViews.Count}");
            FinalizeViews(newViews.ToList());
        }
        
        /// <summary>
        /// Load views sequentially with batch updates - DEPRECATED
        /// Use LoadViewsSequentialFromIds instead
        /// </summary>
        private void LoadViewsSequential(List<View> views, Dictionary<string, string> existingCustomNames)
        {
            const int INITIAL_BATCH_SIZE = 200; // Load first 200 with full details
            
            var newViews = new List<ViewItem>();
            int addedCount = 0;
            int restoredCount = 0;
            
            foreach (var view in views)
            {
                // Check if window is closing
                if (_isClosing)
                {
                    WriteDebugLog($"⚠️ Window closing detected - stopping at view #{addedCount}");
                    return;
                }
                
                // ✅ First 200 items load full details, rest lazy load
                bool shouldLoadFullDetails = (addedCount < INITIAL_BATCH_SIZE);
                var viewItem = ProcessView(view, existingCustomNames, ref restoredCount, loadFullDetails: shouldLoadFullDetails);
                newViews.Add(viewItem);
                addedCount++;
                
                // Batch update UI every BATCH_SIZE items
                if (addedCount % BATCH_SIZE == 0)
                {
                    WriteDebugLog($"📦 Batch update: {addedCount}/{views.Count} views loaded");
                }
            }
            
            WriteDebugLog($"Restored custom filenames: {restoredCount}/{addedCount}");
            FinalizeViews(newViews);
        }
        
        /// <summary>
        /// Load views in parallel - MUST extract Revit API data in main thread first!
        /// </summary>
        private void LoadViewsParallel(List<View> views, Dictionary<string, string> existingCustomNames)
        {
            // CRITICAL: Revit API is NOT thread-safe!
            // Extract ALL data from Revit API in main thread first
            
            WriteDebugLog($"⚡ Extracting view data from {views.Count} views in main thread...");
            
            // Step 1: Extract data in main thread (safe with Revit API)
            var viewDataList = new List<ViewData>();
            int extractedCount = 0;
            
            foreach (var view in views)
            {
                if (_isClosing) return;
                
                try
                {
                    var data = ExtractViewData(view, existingCustomNames);
                    viewDataList.Add(data);
                    extractedCount++;
                    
                    if (extractedCount % BATCH_SIZE == 0)
                    {
                        WriteDebugLog($"📊 Extracted: {extractedCount}/{views.Count} views");
                    }
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"ERROR extracting view {view?.Name}: {ex.Message}");
                }
            }
            
            WriteDebugLog($"✅ View data extraction complete: {viewDataList.Count} views");
            
            if (_isClosing) return;
            
            // Step 2: Process extracted data in parallel (NO Revit API calls!)
            WriteDebugLog($"⚡ Processing {viewDataList.Count} views in parallel...");
            var newViews = new System.Collections.Concurrent.ConcurrentBag<ViewItem>();
            int processedCount = 0;
            int restoredCount = 0;
            
            Parallel.ForEach(viewDataList, new ParallelOptions { MaxDegreeOfParallelism = 4 }, (data, state) =>
            {
                if (_isClosing)
                {
                    state.Stop();
                    return;
                }
                
                var viewItem = CreateViewItemFromData(data);
                if (viewItem != null)
                {
                    newViews.Add(viewItem);
                    
                    if (data.HasCustomFileName)
                    {
                        System.Threading.Interlocked.Increment(ref restoredCount);
                    }
                    
                    int current = System.Threading.Interlocked.Increment(ref processedCount);
                    if (current % BATCH_SIZE == 0)
                    {
                        WriteDebugLog($"⚡ Parallel batch: {current}/{viewDataList.Count} views processed");
                    }
                }
            });
            
            WriteDebugLog($"Restored custom filenames: {restoredCount}/{processedCount}");
            FinalizeViews(newViews.ToList());
        }
        
        // Helper class to store view data
        private class ViewData
        {
            public ElementId ElementId { get; set; } // ⚡ NEW: For ElementIds optimization
            public View View { get; set; }
            public string ViewId { get; set; }
            public string ViewName { get; set; }
            public string ViewType { get; set; }
            public string ViewScale { get; set; } // ⚡ NEW: For fast extraction
            public string Scale { get; set; }
            public string DetailLevel { get; set; }
            public string Discipline { get; set; }
            public string CustomFileName { get; set; }
            public bool HasCustomFileName { get; set; }
        }
        
        /// <summary>
        /// Convert Revit ViewType enum to human-readable string matching AvailableViewTypes list
        /// </summary>
        private string ConvertViewTypeToString(Autodesk.Revit.DB.ViewType viewType)
        {
            switch (viewType)
            {
                case Autodesk.Revit.DB.ViewType.ThreeD:
                    return "3D";
                case Autodesk.Revit.DB.ViewType.FloorPlan:
                    return "Floor Plan";
                case Autodesk.Revit.DB.ViewType.CeilingPlan:
                    return "Ceiling Plan";
                case Autodesk.Revit.DB.ViewType.Elevation:
                    return "Elevation";
                case Autodesk.Revit.DB.ViewType.Section:
                    return "Section";
                case Autodesk.Revit.DB.ViewType.Detail:
                    return "Detail";
                case Autodesk.Revit.DB.ViewType.Rendering:
                    return "Rendering";
                case Autodesk.Revit.DB.ViewType.Legend:
                    return "Legend";
                case Autodesk.Revit.DB.ViewType.EngineeringPlan:
                    return "Engineering Plan";
                case Autodesk.Revit.DB.ViewType.AreaPlan:
                    return "Area Plan";
                default:
                    return viewType.ToString();
            }
        }
        
        // Extract view data in main thread (Revit API calls)
        private ViewData ExtractViewData(View view, Dictionary<string, string> existingCustomNames)
        {
            var data = new ViewData
            {
                View = view,
                ViewId = view.Id.IntegerValue.ToString(),
                ViewName = view.Name ?? "Unnamed",
                ViewType = ConvertViewTypeToString(view.ViewType)  // ✅ FIX: Convert to human-readable
            };
            
            // Extract scale, detail level, discipline
            try
            {
                Parameter scaleParam = view.get_Parameter(BuiltInParameter.VIEW_SCALE);
                data.Scale = scaleParam != null && scaleParam.HasValue ? $"1:{scaleParam.AsInteger()}" : "N/A";
                
                data.DetailLevel = view.DetailLevel.ToString();
                
                Parameter disciplineParam = view.get_Parameter(BuiltInParameter.VIEW_DISCIPLINE);
                data.Discipline = disciplineParam?.AsValueString() ?? "N/A";
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR extracting view details: {ex.Message}");
                data.Scale = "N/A";
                data.DetailLevel = "N/A";
                data.Discipline = "N/A";
            }
            
            // Check for custom filename
            if (!string.IsNullOrEmpty(data.ViewId) && 
                existingCustomNames.TryGetValue(data.ViewId, out string customName))
            {
                data.CustomFileName = customName;
                data.HasCustomFileName = true;
            }
            else
            {
                data.CustomFileName = data.ViewName;
                data.HasCustomFileName = false;
            }
            
            return data;
        }
        
        // Create ViewItem from extracted data (can run in parallel)
        private ViewItem CreateViewItemFromData(ViewData data)
        {
            try
            {
                var viewItem = new ViewItem
                {
                    RevitView = data.View,
                    RevitViewId = data.View.Id,  // CRITICAL: Set ElementId for NWC/IFC export
                    ViewId = data.ViewId,
                    ViewName = data.ViewName,
                    ViewType = data.ViewType,
                    Scale = data.Scale,
                    DetailLevel = data.DetailLevel,
                    Discipline = data.Discipline,
                    CustomFileName = data.CustomFileName,
                    IsSelected = false
                };
                
                // Subscribe to PropertyChanged
                viewItem.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == "IsSelected")
                    {
                        // Auto-enable NWC/IFC for 3D views
                        if (viewItem.IsSelected && viewItem.ViewType != null && 
                            (viewItem.ViewType.Contains("ThreeD") || viewItem.ViewType.Contains("3D")))
                        {
                            if (!ExportSettings.IsNwcSelected)
                                ExportSettings.IsNwcSelected = true;
                            if (!ExportSettings.IsIfcSelected)
                                ExportSettings.IsIfcSelected = true;
                        }
                        UpdateStatusText();
                        UpdateExportSummary();
                    }
                };
                
                return viewItem;
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR creating view item: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// ⚡ FAST: Create ViewItem from pre-extracted data with custom filename restore
        /// </summary>
        private ViewItem CreateViewItemFromData(ViewData data, Dictionary<string, string> existingCustomNames, ref int restoredCount)
        {
            try
            {
                string customFileName = $"{data.ViewName.Replace(" ", "_")}";
                
                // Restore custom filename if exists
                if (existingCustomNames.TryGetValue(data.ViewId, out string existingName))
                {
                    customFileName = existingName;
                    System.Threading.Interlocked.Increment(ref restoredCount);
                }
                
                var viewItem = new ViewItem
                {
                    RevitView = null, // Will be set when needed (lazy)
                    RevitViewId = data.ElementId,
                    ViewId = data.ViewId,
                    ViewName = data.ViewName,
                    ViewType = data.ViewType,
                    Scale = data.ViewScale,
                    DetailLevel = "Not Loaded", // Lazy load
                    Discipline = "Not Loaded",  // Lazy load
                    CustomFileName = customFileName,
                    IsSelected = false
                };
                
                // Subscribe to PropertyChanged
                viewItem.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == "IsSelected")
                    {
                        // Auto-enable NWC/IFC for 3D views
                        if (viewItem.IsSelected && viewItem.ViewType != null && 
                            (viewItem.ViewType.Contains("ThreeD") || viewItem.ViewType.Contains("3D")))
                        {
                            if (!ExportSettings.IsNwcSelected)
                                ExportSettings.IsNwcSelected = true;
                            if (!ExportSettings.IsIfcSelected)
                                ExportSettings.IsIfcSelected = true;
                        }
                        UpdateStatusText();
                        UpdateExportSummary();
                    }
                };
                
                return viewItem;
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR creating view item: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Process a single view item
        /// </summary>
        private ViewItem ProcessView(View view, Dictionary<string, string> existingCustomNames, ref int restoredCount, bool loadFullDetails = false)
        {
            try
            {
                // ✅ IMPROVED: Support both full and lazy loading based on parameter
                // First 200 items: Load full details immediately for fast display
                // Remaining items: Lazy load when scrolled into view
                var viewItem = new ViewItem(view, loadFullDetails: loadFullDetails);
                
                // Restore custom filename
                if (!string.IsNullOrEmpty(viewItem.ViewId) && 
                    existingCustomNames.TryGetValue(viewItem.ViewId, out string customName))
                {
                    viewItem.CustomFileName = customName;
                    restoredCount++;
                }
                
                // ⚡ NO PropertyChanged here - will subscribe AFTER binding completes
                
                return viewItem;
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR processing view {view?.Name}: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Finalize views - sort and update UI
        /// Following Schedule_Editor pattern: Load remaining views in background
        /// </summary>
        private void FinalizeViews(List<ViewItem> views)
        {
            var validViews = views.Where(v => v != null).ToList();
            var sortedViews = validViews.OrderBy(v => v.ViewName, new AlphanumericComparer()).ToList();
            
            Dispatcher.Invoke(() =>
            {
                Views = new ObservableCollection<ViewItem>(sortedViews);
                
                WriteDebugLog($"⚡ NOW subscribing PropertyChanged for {sortedViews.Count} views...");
                
                // ⚡ CRITICAL: Subscribe PropertyChanged AFTER binding completes
                foreach (var view in Views)
                {
                    view.PropertyChanged += ViewItem_PropertyChanged;
                }
                
                WriteDebugLog("✅ View PropertyChanged handlers subscribed");
                UpdateStatusText();
                
                // ✅ IMPROVED: Load remaining views in background (Schedule_Editor pattern)
                LoadRemainingViewsInBackground();
            });
            
            WriteDebugLog($"✅ Finalized {sortedViews.Count} views");
        }
        
        /// <summary>
        /// Load remaining not-fully-loaded views in background (Schedule_Editor pattern)
        /// </summary>
        private async void LoadRemainingViewsInBackground()
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    var viewsToLoad = Views?.Where(v => !v.IsFullyLoaded).ToList() ?? new List<ViewItem>();
                    
                    if (viewsToLoad.Count == 0)
                    {
                        WriteDebugLog("✅ All views already loaded with full details");
                        return;
                    }
                    
                    WriteDebugLog($"🔄 Loading full details for remaining {viewsToLoad.Count} views in background...");
                    
                    const int BACKGROUND_BATCH_SIZE = 500; // Like Schedule_Editor
                    int loadedCount = 0;
                    
                    foreach (var viewItem in viewsToLoad)
                    {
                        if (_isClosing) break;
                        
                        try
                        {
                            // Load full details on UI thread (Revit API requires UI thread)
                            Dispatcher.Invoke(() =>
                            {
                                viewItem.LoadFullDetails(_document);
                            });
                            
                            loadedCount++;
                            
                            if (loadedCount % BACKGROUND_BATCH_SIZE == 0)
                            {
                                WriteDebugLog($"   Background loading: {loadedCount}/{viewsToLoad.Count}");
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteDebugLog($"ERROR loading view details: {ex.Message}");
                        }
                    }
                    
                    WriteDebugLog($"✅ Background loading completed: {loadedCount} views fully loaded");
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"ERROR in LoadRemainingViewsInBackground: {ex.Message}");
                }
            });
        }
        
        // ⚡ Centralized ViewItem PropertyChanged handler
        private void ViewItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsSelected")
            {
                var viewItem = sender as ViewItem;
                if (viewItem == null) return;
                
                // Auto-enable 3D formats when 3D view is selected
                if (viewItem.IsSelected && viewItem.ViewType != null && 
                    (viewItem.ViewType.Contains("ThreeD") || viewItem.ViewType.Contains("3D")))
                {
                    if (!ExportSettings.IsNwcSelected)
                        ExportSettings.IsNwcSelected = true;
                    if (!ExportSettings.IsIfcSelected)
                        ExportSettings.IsIfcSelected = true;
                }
                
                UpdateStatusText();
                UpdateExportSummary();
            }
        }

        private void UpdateViewStatusText()
        {
            var selectedCount = Views?.Count(v => v.IsSelected) ?? 0;
            var totalCount = Views?.Count ?? 0;
            WriteDebugLog($"[Export +] UpdateViewStatusText called - Selected: {selectedCount}, Total: {totalCount}");
            UpdateCreateTabSummary();
        }

        private void UpdateStatusText()
        {
            var selectedSheetsCount = Sheets?.Count(s => s.IsSelected) ?? 0;
            var totalSheetsCount = Sheets?.Count ?? 0;
            var selectedViewsCount = Views?.Count(v => v.IsSelected) ?? 0;
            var totalViewsCount = Views?.Count ?? 0;
            
            WriteDebugLog($"[Export +] UpdateStatusText called - Sheets: {selectedSheetsCount}/{totalSheetsCount}, Views: {selectedViewsCount}/{totalViewsCount}");
            
            // Check if user has selected 3D views
            var has3DViews = Views?.Any(v => v.IsSelected && 
                v.ViewType != null && 
                (v.ViewType.Contains("ThreeD") || v.ViewType.Contains("3D"))) ?? false;
            
            // Disable NWC and IFC if no 3D views selected or if only sheets are selected
            var shouldDisableNwcIfc = !has3DViews || (selectedSheetsCount > 0 && selectedViewsCount == 0);
            
            // Update NWC and IFC checkbox states
            if (ExportSettings != null)
            {
                if (shouldDisableNwcIfc)
                {
                    if (ExportSettings.IsNwcSelected)
                    {
                        ExportSettings.IsNwcSelected = false;
                        WriteDebugLog("NWC export disabled - requires 3D views");
                    }
                    if (ExportSettings.IsIfcSelected)
                    {
                        ExportSettings.IsIfcSelected = false;
                        WriteDebugLog("IFC export disabled - requires 3D views");
                    }
                }
            }
            
            // Disable/enable UI checkboxes with visual feedback
            try
            {
                if (NWCCheck != null)
                {
                    NWCCheck.IsEnabled = !shouldDisableNwcIfc;
                    if (shouldDisableNwcIfc)
                    {
                        NWCCheck.ToolTip = "NWC export requires 3D views to be selected";
                        NWCCheck.Opacity = 0.5;
                    }
                    else
                    {
                        NWCCheck.ToolTip = "Export to Navisworks NWC format";
                        NWCCheck.Opacity = 1.0;
                    }
                }
                
                if (IFCCheck != null)
                {
                    IFCCheck.IsEnabled = !shouldDisableNwcIfc;
                    if (shouldDisableNwcIfc)
                    {
                        IFCCheck.ToolTip = "IFC export requires 3D views to be selected";
                        IFCCheck.Opacity = 0.5;
                    }
                    else
                    {
                        IFCCheck.ToolTip = "Export to IFC format";
                        IFCCheck.Opacity = 1.0;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error updating NWC/IFC checkbox states: {ex.Message}");
            }
            
            // Update status text controls
            try
            {
                // Update sheet count text
                if (SheetsCountText != null)
                {
                    SheetsCountText.Text = $"{selectedSheetsCount} sheets selected";
                }
                
                // Update views count text  
                if (ViewsCountText != null)
                {
                    ViewsCountText.Text = $"{selectedViewsCount} views selected";
                }
                
                // Update total selected items text
                if (TotalItemsText != null)
                {
                    var totalSelected = selectedSheetsCount + selectedViewsCount;
                    TotalItemsText.Text = $"Total: {totalSelected} selected";
                }
                
                // Don't auto-sync "All" checkbox - let user control it manually
                
                // Use the same totalSelected from above for title
                var totalItemsForTitle = totalSheetsCount + totalViewsCount;
                var totalSelectedForTitle = selectedSheetsCount + selectedViewsCount;
                this.Title = $"Export + - {totalSelectedForTitle} of {totalItemsForTitle} items selected ({selectedSheetsCount} sheets, {selectedViewsCount} views)";
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error updating status controls: {ex.Message}");
            }
            
            UpdateCreateTabSummary();
        }

        private void UpdateCreateTabSummary()
        {
            try
            {
                // Update selection summary
                var sheetsSelected = Sheets?.Count(s => s.IsSelected) ?? 0;
                var viewsSelected = Views?.Count(v => v.IsSelected) ?? 0;
                var totalSelected = sheetsSelected + viewsSelected;
                
                // NOTE: SelectionSummaryText removed from new Create tab design
                // Status is shown in DataGrid instead
                
                // Update format summary
                // NOTE: FormatSummaryText removed from new Create tab design  
                // Formats shown in Export Queue DataGrid instead
                
                // Refresh SelectedItemsForExport binding
                OnPropertyChanged(nameof(SelectedItemsForExport));
                
                WriteDebugLog($"Create tab summary updated: {totalSelected} items");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error updating Create tab summary: {ex.Message}");
            }
        }

        private void UpdateFormatSelection()
        {
            WriteDebugLog("[Export +] UpdateFormatSelection called");
            
            try
            {
                WriteDebugLog($"[Export +] Current format states: {string.Join(", ", ExportSettings?.GetSelectedFormatsList() ?? new List<string>())}");
                
                // Format selection is handled by data binding in XAML
                WriteDebugLog("[Export +] Format selection updated via data binding");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"[Export +] ERROR in UpdateFormatSelection: {ex.Message}");
            }
        }

        /// <summary>
        /// Reset addin after export completed - cho phép user chọn lại và export tiếp
        /// QUAN TRỌNG: KHÔNG reset format selection (PDF/DWG checkboxes) - giữ nguyên lựa chọn của user
        /// </summary>
        private void ResetAddinAfterExport()
        {
            WriteDebugLog("🔄 [Reset] Starting addin reset after export completion");
            
            try
            {
                // 1. Clear Export Queue
                if (ExportQueueDataGrid?.ItemsSource is ObservableCollection<ExportQueueItem> queueItems)
                {
                    queueItems.Clear();
                    WriteDebugLog("✓ Export queue cleared");
                }
                
                // 2. Reset Progress UI
                if (ExportProgressBar != null)
                {
                    ExportProgressBar.Value = 0;
                }
                if (ProgressPercentageText != null)
                {
                    ProgressPercentageText.Text = "Completed 0%";
                }
                WriteDebugLog("✓ Progress UI reset");
                
                // 3. Reset Export Button
                if (StartExportButton != null)
                {
                    StartExportButton.IsEnabled = true;
                    StartExportButton.Content = "START EXPORT";
                }
                WriteDebugLog("✓ Export button reset");
                
                // 4. Update Export Queue với selection mới
                // NOTE: UpdateExportQueue() sẽ tự động lấy formats từ ExportSettings hiện tại
                // Không reset ExportSettings → giữ nguyên PDF/DWG checkboxes
                WriteDebugLog($"📋 Current format selection: {string.Join(", ", ExportSettings?.GetSelectedFormatsList() ?? new List<string>())}");
                UpdateExportQueue();
                WriteDebugLog("✓ Export queue updated with new selection");
                
                WriteDebugLog("🎉 [Reset] Addin reset completed - Ready for next export!");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"❌ [Reset] Error during addin reset: {ex.Message}");
            }
        }

        private void UpdateExportSummary()
        {
            try
            {
                var selectedCount = SelectedSheetsCount;
                var selectedFormats = ExportSettings?.GetSelectedFormatsList() ?? new List<string>();
                var estimatedFiles = selectedCount * selectedFormats.Count;

                // Update export settings with current selection count
                if (ExportSettings != null)
                {
                    ExportSettings.SelectedSheetsCount = selectedCount;
                }

                WriteDebugLog($"[Export +] Export summary updated: {selectedCount} sheets, {selectedFormats.Count} formats, {estimatedFiles} files");
                
                // Update Export Queue for Create tab
                UpdateExportQueue();
            }
            catch (Exception ex)
            {
                WriteDebugLog($"[Export +] ERROR in UpdateExportSummary: {ex.Message}");
            }
        }

        /// <summary>
        /// Update Export Queue DataGrid based on selected sheets/views and formats
        /// </summary>
        private void UpdateExportQueue()
        {
            try
            {
                if (ExportQueueItems == null) return;

                ExportQueueItems.Clear();

                var selectedFormats = ExportSettings?.GetSelectedFormatsList() ?? new List<string>();
                
                // DEBUG: Log chi tiết format states
                WriteDebugLog("========== FORMAT SELECTION DEBUG ==========");
                WriteDebugLog($"PDF Checkbox: {ExportSettings?.IsPdfSelected}");
                WriteDebugLog($"DWG Checkbox: {ExportSettings?.IsDwgSelected}");
                WriteDebugLog($"NWC Checkbox: {ExportSettings?.IsNwcSelected}");
                WriteDebugLog($"IFC Checkbox: {ExportSettings?.IsIfcSelected}");
                WriteDebugLog($"Selected formats count: {selectedFormats.Count}");
                WriteDebugLog($"Selected formats list: [{string.Join(", ", selectedFormats)}]");
                WriteDebugLog("==========================================");
                
                if (selectedFormats.Count == 0)
                {
                    WriteDebugLog("⚠️ No formats selected, Export Queue cleared");
                    return;
                }
                
                WriteDebugLog($"UpdateExportQueue: Processing {Sheets?.Count ?? 0} sheets, {Views?.Count ?? 0} views");
                WriteDebugLog($"Selected formats: {string.Join(", ", selectedFormats)}");

                // Add selected sheets to queue (only PDF/DWG formats, skip NWC/IFC)
                if (Sheets != null)
                {
                    var selectedSheets = Sheets.Where(s => s.IsSelected).ToList();
                    WriteDebugLog($"Found {selectedSheets.Count} selected sheets");
                    
                    foreach (var sheet in selectedSheets)
                    {
                        foreach (var format in selectedFormats)
                        {
                            // Sheets only support PDF/DWG/DWF/IMG, skip NWC/IFC
                            var formatUpper = format.ToUpper();
                            if (formatUpper == "NWC" || formatUpper == "IFC")
                            {
                                WriteDebugLog($"Skipping {formatUpper} for sheet {sheet.SheetNumber} - sheets don't support NWC/IFC export");
                                continue;
                            }
                            
                            // Determine display name: use CustomFileName if available, else SheetName
                            string displayName = sheet.SheetName;
                            if (!string.IsNullOrWhiteSpace(sheet.CustomFileName))
                            {
                                displayName = sheet.CustomFileName;
                            }
                            
                            var queueItem = new ExportQueueItem
                            {
                                IsSelected = true,
                                ViewSheetNumber = sheet.SheetNumber,
                                ViewSheetName = displayName,
                                Format = format.ToUpper(),
                                Size = GetSheetSize(sheet),
                                Orientation = GetSheetOrientation(sheet),
                                Progress = 0,
                                Status = "Pending"
                            };
                            ExportQueueItems.Add(queueItem);
                            WriteDebugLog($"  ✓ Added to queue: {sheet.SheetNumber} - {displayName} - Format: {format.ToUpper()}");
                        }
                    }
                }

                // Add selected views to queue
                if (Views != null)
                {
                    var selectedViews = Views.Where(v => v.IsSelected).ToList();
                    WriteDebugLog($"Found {selectedViews.Count} selected views");
                    
                    foreach (var view in selectedViews)
                    {
                        // Check if this is a 3D view
                        bool is3DView = view.ViewType != null && 
                                       (view.ViewType.Contains("ThreeD") || view.ViewType.Contains("3D"));
                        
                        WriteDebugLog($"Processing view: {view.ViewName} (Type: {view.ViewType}, is3D: {is3DView})");
                        
                        foreach (var format in selectedFormats)
                        {
                            var formatUpper = format.ToUpper();
                            
                            // 3D views: only NWC/IFC, skip PDF/DWG
                            if (is3DView && (formatUpper == "PDF" || formatUpper == "DWG" || formatUpper == "DWF" || formatUpper == "IMG"))
                            {
                                WriteDebugLog($"Skipping {formatUpper} for 3D view {view.ViewName} - 3D views only support NWC/IFC export");
                                continue;
                            }
                            
                            // 2D views: only PDF/DWG, skip NWC/IFC
                            if (!is3DView && (formatUpper == "NWC" || formatUpper == "IFC"))
                            {
                                WriteDebugLog($"Skipping {formatUpper} for 2D view {view.ViewName} - 2D views don't support NWC/IFC export");
                                continue;
                            }
                            
                            // Determine display name: use CustomFileName if available, else ViewName
                            string displayName = view.ViewName;
                            if (!string.IsNullOrWhiteSpace(view.CustomFileName))
                            {
                                displayName = view.CustomFileName;
                            }
                            
                            // ✅ FIX: Convert "ThreeD" to "3D" for display
                            string viewTypeDisplay = view.ViewType;
                            if (viewTypeDisplay == "ThreeD")
                            {
                                viewTypeDisplay = "3D";
                            }
                            
                            var queueItem = new ExportQueueItem
                            {
                                IsSelected = true,
                                ViewSheetNumber = viewTypeDisplay,
                                ViewSheetName = displayName,
                                Format = format.ToUpper(),
                                Size = "-",
                                Orientation = "-",
                                Progress = 0,
                                Status = "Pending"
                            };
                            ExportQueueItems.Add(queueItem);
                            WriteDebugLog($"  ✓ Added to queue: {displayName} - Format: {formatUpper}");
                        }
                    }
                }

                WriteDebugLog($"Export Queue updated: {ExportQueueItems.Count} items");
                
                // DEBUG: List all items in queue
                WriteDebugLog("\n========== EXPORT QUEUE SUMMARY ==========");
                for (int i = 0; i < ExportQueueItems.Count; i++)
                {
                    var item = ExportQueueItems[i];
                    WriteDebugLog($"[{i+1}] {item.ViewSheetName}");
                    WriteDebugLog($"    Format: {item.Format}");
                    WriteDebugLog($"    Status: {item.Status}");
                }
                WriteDebugLog("==========================================\n");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR in UpdateExportQueue: {ex.Message}");
            }
        }

        /// <summary>
        /// Get sheet size (paper size) from sheet
        /// </summary>
        private string GetSheetSize(SheetItem sheet)
        {
            try
            {
                // ✅ FIX: Use the Size property from SheetItem which already has "A1", "A2", etc.
                // This ensures consistency between Sheets tab and Create tab
                if (!string.IsNullOrEmpty(sheet?.Size))
                {
                    return sheet.Size;
                }
                
                // Fallback: Try to get from Revit if Size is not available
                if (sheet?.Id == null || _document == null) return "-";

                var revitSheet = _document.GetElement(sheet.Id) as ViewSheet;
                if (revitSheet == null) return "-";

                // Use SheetSizeDetector for consistency
                string detectedSize = Utils.SheetSizeDetector.GetSheetSize(revitSheet);
                return !string.IsNullOrEmpty(detectedSize) ? detectedSize : "Custom";
            }
            catch
            {
                return sheet?.Size ?? "-";
            }
        }

        /// <summary>
        /// Get sheet orientation (Portrait/Landscape)
        /// </summary>
        private string GetSheetOrientation(SheetItem sheet)
        {
            try
            {
                if (sheet?.Id == null || _document == null) return "-";

                var revitSheet = _document.GetElement(sheet.Id) as ViewSheet;
                if (revitSheet == null) return "-";

                var titleBlock = new FilteredElementCollector(_document, revitSheet.Id)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .FirstOrDefault();

                if (titleBlock != null)
                {
                    var widthParam = titleBlock.get_Parameter(BuiltInParameter.SHEET_WIDTH);
                    var heightParam = titleBlock.get_Parameter(BuiltInParameter.SHEET_HEIGHT);

                    if (widthParam != null && heightParam != null)
                    {
                        double width = widthParam.AsDouble();
                        double height = heightParam.AsDouble();

                        return width > height ? "Landscape" : "Portrait";
                    }
                }

                return "-";
            }
            catch
            {
                return "-";
            }
        }

        // Event Handlers for Export + Interface - moved to Profile Manager Methods region

        private async Task ApplySheetSelectionAsync(IEnumerable<SheetItem> items, bool selectAll)
        {
            if (items == null)
            {
                return;
            }

            const int BATCH_SIZE = 50;
            int processed = 0;

            foreach (var sheet in items)
            {
                if (sheet == null)
                {
                    continue;
                }

                if (sheet.IsSelected != selectAll)
                {
                    sheet.IsSelected = selectAll;
                }

                processed++;

                if (processed % BATCH_SIZE == 0)
                {
                    await Dispatcher.Yield(DispatcherPriority.Background);
                }
            }
        }

        private async Task ApplyViewSelectionAsync(IEnumerable<ViewItem> items, bool selectAll)
        {
            if (items == null)
            {
                return;
            }

            const int BATCH_SIZE = 50;
            int processed = 0;

            foreach (var view in items)
            {
                if (view == null)
                {
                    continue;
                }

                if (view.IsSelected != selectAll)
                {
                    view.IsSelected = selectAll;
                }

                processed++;

                if (processed % BATCH_SIZE == 0)
                {
                    await Dispatcher.Yield(DispatcherPriority.Background);
                }
            }
        }

        private async void ToggleAll_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("[Export +] Toggle All clicked");
            
            if (Sheets != null && Sheets.Any())
            {
                bool selectAll = !Sheets.All(s => s.IsSelected);
                await ApplySheetSelectionAsync(Sheets, selectAll);
                WriteDebugLog($"[Export +] Toggled all sheets to: {selectAll}");
                UpdateStatusText();
                UpdateExportSummary();
            }
        }

        private void EditCustomDrawingNumber_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("[Export +] Edit Custom Drawing Number clicked");
            MessageBox.Show("Custom Drawing Number Editor sẽ được thêm trong phiên bản tiếp theo!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void FormatToggle_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Primitives.ToggleButton button && button.Tag is string format)
            {
                WriteDebugLog($"[Export +] Format {format} checked via ToggleButton");
                ExportSettings?.SetFormatSelection(format, true);
                UpdateExportSummary();
            }
        }

        private void FormatToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Primitives.ToggleButton button && button.Tag is string format)
            {
                WriteDebugLog($"[Export +] Format {format} unchecked via ToggleButton");
                ExportSettings?.SetFormatSelection(format, false);
                UpdateExportSummary();
            }
        }

        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("[Export +] Browse folder clicked - Export + Enhanced version");
            
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            
            if (!string.IsNullOrEmpty(ExportSettings?.OutputFolder))
            {
                dialog.SelectedPath = ExportSettings.OutputFolder;
                WriteDebugLog($"[Export +] Current folder: {ExportSettings.OutputFolder}");
            }
            
            dialog.Description = "Chọn thư mục xuất file Export +";
            dialog.ShowNewFolderButton = true;
            
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ExportSettings.OutputFolder = dialog.SelectedPath;
                WriteDebugLog($"[Export +] Folder updated: {dialog.SelectedPath}");
                UpdateExportSummary();
            }
        }

        private void Create_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("===== CREATE BUTTON CLICKED (EXPORT + ENHANCED) =====");
            
            var selectedSheets = Sheets?.Where(s => s.IsSelected).ToList();
            WriteDebugLog($"Found {selectedSheets?.Count ?? 0} selected sheets for export");
            
            // Log selected sheet details
            if (selectedSheets != null && selectedSheets.Any())
            {
                foreach (var sheet in selectedSheets)
                {
                    WriteDebugLog($"Selected sheet: {sheet.Number} - {sheet.Name}");
                }
            }
            
            if (selectedSheets == null || !selectedSheets.Any())
            {
                WriteDebugLog("VALIDATION ERROR: No sheets selected for export");
                MessageBox.Show("Vui lòng chọn ít nhất một sheet để export!", "Cảnh báo", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            var outputPath = ExportSettings?.OutputFolder ?? "";
            WriteDebugLog($"Output path validation: '{outputPath}'");
            
            if (string.IsNullOrEmpty(outputPath))
            {
                WriteDebugLog("VALIDATION ERROR: Empty or null output path");
                MessageBox.Show("Vui lòng chọn thư mục xuất file!", "Cảnh báo", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            var selectedFormats = ExportSettings?.GetSelectedFormatsList() ?? new List<string>();
            WriteDebugLog($"Selected export formats: [{string.Join(", ", selectedFormats)}]");
            
            if (!selectedFormats.Any())
            {
                WriteDebugLog("VALIDATION ERROR: No export formats selected");
                MessageBox.Show("Vui lòng chọn ít nhất một định dạng file!", "Cảnh báo", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            // Show detailed export summary
            var summary = $@"EXPORT + SUMMARY
            
Sheets: {selectedSheets.Count}
Formats: {string.Join(", ", selectedFormats)}
Output: {outputPath}
Estimated Files: {selectedSheets.Count * selectedFormats.Count}

Template: {ExportSettings?.FileNameTemplate ?? "Default"}
Combine Files: {ExportSettings?.CombineFiles ?? false}
Include Revision: {ExportSettings?.IncludeRevision ?? false}

Tiếp tục xuất file?";
            
            WriteDebugLog($"[Export +] Showing export summary dialog");
            var result = MessageBox.Show(summary, "Export + Confirmation", 
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            
            if (result == MessageBoxResult.Yes)
            {
                WriteDebugLog("[Export +] User confirmed export - Starting export process...");
                
                try
                {
                    // Update status via window title
                    this.Title = "Export + - Exporting...";

                    bool exportSuccess = false;
                    int totalExported = 0;

                    // Convert SheetItem to ViewSheet for export
                    var sheetsToExport = new List<ViewSheet>();
                    foreach (var sheetItem in selectedSheets)
                    {
                        // Find the actual ViewSheet from document
                        var collector = new FilteredElementCollector(_document);
                        var sheet = collector.OfClass(typeof(ViewSheet))
                                           .Cast<ViewSheet>()
                                           .FirstOrDefault(s => s.SheetNumber == sheetItem.Number);
                        if (sheet != null)
                        {
                            sheetsToExport.Add(sheet);
                        }
                    }

                    WriteDebugLog($"[Export +] Found {sheetsToExport.Count} ViewSheets for export");

                    // Export to different formats
                    foreach (var format in selectedFormats)
                    {
                        WriteDebugLog($"[Export +] Starting export to {format}");
                        
                        if (format.ToUpper() == "PDF")
                        {
#if REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                            var pdfManager = new PDFExportManager(_document);
                            bool pdfResult = pdfManager.ExportSheetsToPDF(sheetsToExport, outputPath, ExportSettings);
                            if (pdfResult)
                            {
                                totalExported += sheetsToExport.Count;
                                exportSuccess = true;
                                WriteDebugLog($"[Export +] PDF export completed successfully");
                            }
#else
                            WriteDebugLog($"[Export +] PDF export not supported in Revit {_document.Application.VersionNumber}");
#endif
                        }
                        else
                        {
                            WriteDebugLog($"[Export +] Format {format} not yet implemented");
                        }
                    }
                    
                    WriteDebugLog("[Export +] Export process completed");
                    
                    if (exportSuccess)
                    {
                        // Show export completed dialog with Open Folder button
                        ShowExportCompletedDialog(outputPath);
                    }
                    else
                    {
                        MessageBox.Show("Export failed or no files were exported.", 
                            "Export Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                        
                    // Update status via window title
                    this.Title = exportSuccess ? "Export + - Export completed" : "Export + - Export failed";
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"[Export +] ERROR in export process: {ex.Message}");
                    MessageBox.Show($"Export error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    
                    // Update status via window title
                    this.Title = "Export + - Export failed";
                }
            }
            else
            {
                WriteDebugLog("[Export +] User cancelled export");
            }
        }

        // ViewDebugLog_Click method removed - use DebugView instead to see OutputDebugStringA logs

        // Legacy event handlers for compatibility

        private void SheetsDataGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            WriteDebugLog("[ExportPlus] DataGrid mouse button down");
        }

        private void SheetsDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            WriteDebugLog("[ExportPlus] DataGrid cell edit ending");
        }

        private void SheetsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WriteDebugLog("[ExportPlus] DataGrid selection changed");
            
            // ✓ RESET ADDIN nếu user chọn lại sheet sau khi export xong
            if (_exportJustCompleted)
            {
                WriteDebugLog("🔄 Export completed detected - Resetting addin for reuse");
                ResetAddinAfterExport();
                _exportJustCompleted = false;
            }
            
            UpdateStatusText();
        }

        // New event handlers for enhanced UI
        private void SheetsRadio_Checked(object sender, RoutedEventArgs e)
        {
            WriteDebugLog($"Sheets radio button checked - IsLoaded={this.IsLoaded}, _windowFullyLoaded={_windowFullyLoaded}");
            if (SheetsDataGrid != null && ViewsDataGrid != null)
            {
                SheetsDataGrid.Visibility = System.Windows.Visibility.Visible;
                ViewsDataGrid.Visibility = System.Windows.Visibility.Collapsed;
                
                // ⚡⚡⚡ CRITICAL: CHỈ load nếu Window_Loaded ĐÃ CHẠY XONG
                // Nếu chưa, Window_Loaded sẽ lo việc load
                if (_windowFullyLoaded && !_sheetsLoaded)
                {
                    WriteDebugLog("⚡ First time loading sheets - SYNC mode (Revit API requirement)...");
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    
                    // ⚡⚡⚡ Load SYNCHRONOUSLY - Revit API must stay on main thread!
                    LoadSheetsSync();
                    
                    sw.Stop();
                    _sheetsLoaded = true;
                    WriteDebugLog($"✅ Sheets loaded in {sw.ElapsedMilliseconds}ms");
                }
                else if (!_windowFullyLoaded)
                {
                    WriteDebugLog("⚡ Window chưa fully loaded - SKIP load (ExportPlusMainWindow_Loaded sẽ lo)");
                }
                else
                {
                    WriteDebugLog($"⚡ Sheets already loaded (_sheetsLoaded={_sheetsLoaded})");
                }
                
                UpdateStatusText(); // Update checkbox state for Sheets
            }
        }

        private void ViewsRadio_Checked(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("🔵 Views radio button checked");
            if (SheetsDataGrid != null && ViewsDataGrid != null)
            {
                SheetsDataGrid.Visibility = System.Windows.Visibility.Collapsed;
                ViewsDataGrid.Visibility = System.Windows.Visibility.Visible;
                
                WriteDebugLog($"🔍 DEBUG: _viewsLoaded = {_viewsLoaded}, _windowFullyLoaded = {_windowFullyLoaded}");
                
                // ⚡ LAZY LOADING: Only load if not already loaded
                if (!_viewsLoaded && _windowFullyLoaded)
                {
                    WriteDebugLog("⚡ First time loading views...");
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    LoadViews();
                    sw.Stop();
                    _viewsLoaded = true;
                    WriteDebugLog($"✅ Views loaded in {sw.ElapsedMilliseconds}ms - Total: {Views?.Count ?? 0} views");
                    
                    // 🆕 Load visible rows immediately after views are loaded
                    Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new System.Action(() =>
                    {
                        WriteDebugLog("👁️ Loading initial visible view rows...");
                        LoadVisibleViewRows();
                    }));
                }
                else if (!_windowFullyLoaded)
                {
                    WriteDebugLog("⚠️ Window not fully loaded yet - views will load after form shown");
                }
                else
                {
                    WriteDebugLog($"ℹ️ Views already loaded - Count: {Views?.Count ?? 0}");
                }
                
                UpdateStatusText(); // Update checkbox state for Views
            }
        }

        private void ViewsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WriteDebugLog("[ExportPlus] Views DataGrid selection changed");
            
            // ✓ RESET ADDIN nếu user chọn lại view sau khi export xong
            if (_exportJustCompleted)
            {
                WriteDebugLog("🔄 Export completed detected - Resetting addin for reuse");
                ResetAddinAfterExport();
                _exportJustCompleted = false;
            }
            
            UpdateStatusText();
            UpdateExportSummary(); // This will call UpdateExportQueue()
        }

        private void ViewSheetSetCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Legacy method - no longer used with multi-select
            // Kept for compatibility
            WriteDebugLog("ViewSheetSetCombo_SelectionChanged called (legacy - multi-select now active)");
        }
        
        private void ViewSheetSetCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("ViewSheetSet checkbox changed");
            OnPropertyChanged(nameof(SelectedSetsDisplay));
            
            // Auto-apply filter if filter checkbox is checked
            if (FilterByVSCheckBox?.IsChecked == true)
            {
                ApplyMultiSetFilter();
            }
        }

        private void FilterByVSCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Filter by V/S checkbox checked - enabling filter");
            // Apply multi-set filter
            ApplyMultiSetFilter();
        }

        private void FilterByVSCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Filter by V/S checkbox unchecked - showing all items");
            // Reset to show all sheets/views
            ResetFilter_Click(sender, e);
        }

        /// <summary>
        /// Handle Save V/S Set button click - shows context menu
        /// </summary>
        private void SaveVSSetButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("SaveVSSetButton clicked - opening context menu");
            
            // Open context menu programmatically
            if (sender is Button button && button.ContextMenu != null)
            {
                button.ContextMenu.PlacementTarget = button;
                button.ContextMenu.IsOpen = true;
            }
        }
        
        /// <summary>
        /// Create new View/Sheet Set
        /// </summary>
        private void NewVSSet_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("New V/S Set clicked");
            SaveVSSet_Click(sender, e); // Call existing save logic
        }
        
        /// <summary>
        /// Add selected items to an existing View/Sheet Set
        /// </summary>
        private void AddToExistingVSSet_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Add to Existing V/S Set clicked");
            
            try
            {
                bool isSheetMode = SheetsRadio?.IsChecked == true;
                
                // Get selected items
                List<ElementId> selectedIds;
                int selectedCount;
                
                if (isSheetMode)
                {
                    selectedIds = Sheets?.Where(s => s.IsSelected).Select(s => s.Id).ToList();
                    selectedCount = selectedIds?.Count ?? 0;
                }
                else
                {
                    selectedIds = Views?.Where(v => v.IsSelected).Select(v => v.RevitViewId).ToList();
                    selectedCount = selectedIds?.Count ?? 0;
                }
                
                if (selectedCount == 0)
                {
                    MessageBox.Show(
                        $"Please select at least one {(isSheetMode ? "sheet" : "view")} to add to a set.",
                        "No Selection",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }
                
                // Show dialog to select existing set
                var dialog = new SelectExistingSetDialog(ViewSheetSets);
                if (dialog.ShowDialog() == true && !string.IsNullOrEmpty(dialog.SelectedSetName))
                {
                    WriteDebugLog($"Adding {selectedCount} items to set: {dialog.SelectedSetName}");
                    
                    bool success = _viewSheetSetManager.AddToExistingSet(dialog.SelectedSetName, selectedIds);
                    
                    if (success)
                    {
                        // Reload sets to update counts
                        LoadViewSheetSets();
                        
                        MessageBox.Show(
                            $"Added {selectedCount} {(isSheetMode ? "sheet" : "view")}{(selectedCount > 1 ? "s" : "")} to '{dialog.SelectedSetName}'.",
                            "Success",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show(
                            $"Failed to add items to '{dialog.SelectedSetName}'.",
                            "Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR in AddToExistingVSSet: {ex.Message}");
                MessageBox.Show(
                    $"Error adding to existing set:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// Delete a View/Sheet Set
        /// </summary>
        private void DeleteVSSet_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Delete V/S Set clicked");
            
            try
            {
                // Get selected set(s)
                var selectedSets = ViewSheetSets?.Where(s => s.IsSelected && !s.IsBuiltIn).ToList();
                
                if (selectedSets == null || !selectedSets.Any())
                {
                    MessageBox.Show(
                        "Please select a View/Sheet Set from the dropdown to delete.\n\n" +
                        "Note: Built-in sets (All Sheets, All Views) cannot be deleted.",
                        "No Set Selected",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }
                
                // Show confirmation dialog
                string setsToDelete = string.Join(", ", selectedSets.Select(s => s.Name));
                var result = MessageBox.Show(
                    $"Are you sure you want to delete the following View/Sheet Set(s)?\n\n{setsToDelete}\n\n" +
                    "This action cannot be undone.",
                    "Confirm Delete",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                    
                if (result == MessageBoxResult.Yes)
                {
                    int deletedCount = 0;
                    foreach (var set in selectedSets)
                    {
                        WriteDebugLog($"Deleting set: {set.Name}");
                        if (_viewSheetSetManager.DeleteViewSheetSet(set.Name))
                        {
                            deletedCount++;
                        }
                    }
                    
                    // Reload sets
                    LoadViewSheetSets();
                    
                    MessageBox.Show(
                        $"Deleted {deletedCount} View/Sheet Set{(deletedCount > 1 ? "s" : "")} successfully.",
                        "Success",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR in DeleteVSSet: {ex.Message}");
                MessageBox.Show(
                    $"Error deleting View/Sheet Set:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Original Save V/S Set logic (called by NewVSSet_Click)
        /// </summary>
        private void SaveVSSet_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Save View/Sheet Set clicked");
            
            try
            {
                bool isSheetMode = SheetsRadio?.IsChecked == true;
                int selectedCount = 0;
                List<ElementId> selectedIds = new List<ElementId>();
                
                if (isSheetMode)
                {
                    var selectedSheets = Sheets?.Where(s => s.IsSelected).ToList();
                    selectedCount = selectedSheets?.Count ?? 0;
                    selectedIds = selectedSheets?.Select(s => s.Id).ToList() ?? new List<ElementId>();
                }
                else
                {
                    var selectedViews = Views?.Where(v => v.IsSelected).ToList();
                    selectedCount = selectedViews?.Count ?? 0;
                    selectedIds = selectedViews?.Select(v => v.RevitViewId).ToList() ?? new List<ElementId>();
                }
                
                WriteDebugLog($"Selected {selectedCount} items for saving to View/Sheet Set");
                
                if (selectedCount == 0)
                {
                    MessageBox.Show(
                        $"No {(isSheetMode ? "sheets" : "views")} selected.\n\n" +
                        $"Please select at least one {(isSheetMode ? "sheet" : "view")} to save as a View/Sheet Set.",
                        "No Selection",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
                
                // Open Save dialog
                var dialog = new SaveViewSheetSetDialog(selectedCount, isSheetMode);
                dialog.Owner = this;
                
                if (dialog.ShowDialog() == true)
                {
                    string setName = dialog.SetName;
                    WriteDebugLog($"User entered set name: {setName}");
                    
                    // Check if name already exists
                    if (_viewSheetSetManager.SetNameExists(setName))
                    {
                        var result = MessageBox.Show(
                            $"A View/Sheet Set named '{setName}' already exists.\n\n" +
                            "Do you want to replace it?",
                            "Set Already Exists",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question);
                            
                        if (result == MessageBoxResult.No)
                            return;
                            
                        // Delete existing set
                        _viewSheetSetManager.DeleteViewSheetSet(setName);
                        WriteDebugLog($"Deleted existing set: {setName}");
                    }
                    
                    // Create new ViewSheetSet
                    try
                    {
                        var viewSheetSet = _viewSheetSetManager.CreateViewSheetSet(setName, selectedIds);
                        
                        if (viewSheetSet != null)
                        {
                            WriteDebugLog($"Successfully created ViewSheetSet: {setName}");
                            
                            // Reload the dropdown
                            LoadViewSheetSets();
                            
                            // Auto-select the newly created set
                            var newSet = ViewSheetSets?.FirstOrDefault(s => s.Name == setName);
                            if (newSet != null)
                            {
                                newSet.IsSelected = true;
                                OnPropertyChanged(nameof(SelectedSetsDisplay));
                            }
                            
                            MessageBox.Show(
                                $"View/Sheet Set '{setName}' saved successfully!\n\n" +
                                $"Contains {selectedCount} {(isSheetMode ? "sheet" : "view")}{(selectedCount > 1 ? "s" : "")}.",
                                "Success",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteDebugLog($"ERROR creating ViewSheetSet: {ex.Message}");
                        MessageBox.Show(
                            $"Failed to create View/Sheet Set:\n\n{ex.Message}",
                            "Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR in SaveVSSet_Click: {ex.Message}");
                MessageBox.Show(
                    $"An error occurred:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox searchBox = sender as TextBox;
            string searchText = searchBox?.Text?.ToLower() ?? "";
            WriteDebugLog($"Search text changed: '{searchText}'");
            
            if (string.IsNullOrWhiteSpace(searchText))
            {
                // Show all items when search is empty
                if (SheetsDataGrid.Visibility == System.Windows.Visibility.Visible)
                {
                    SheetsDataGrid.ItemsSource = Sheets;
                }
                else if (ViewsDataGrid.Visibility == System.Windows.Visibility.Visible)
                {
                    ViewsDataGrid.ItemsSource = Views;
                }
            }
            else
            {
                // Filter based on current view
                if (SheetsDataGrid.Visibility == System.Windows.Visibility.Visible && Sheets != null)
                {
                    var filtered = Sheets.Where(s => 
                        (s.SheetNumber?.ToLower().Contains(searchText) ?? false) ||
                        (s.SheetName?.ToLower().Contains(searchText) ?? false) ||
                        (s.CustomFileName?.ToLower().Contains(searchText) ?? false)
                    ).ToList();
                    
                    SheetsDataGrid.ItemsSource = filtered;
                    WriteDebugLog($"Filtered sheets: {filtered.Count} of {Sheets.Count}");
                }
                else if (ViewsDataGrid.Visibility == System.Windows.Visibility.Visible && Views != null)
                {
                    var filtered = Views.Where(v => 
                        (v.ViewName?.ToLower().Contains(searchText) ?? false) ||
                        (v.ViewType?.ToLower().Contains(searchText) ?? false) ||
                        (v.CustomFileName?.ToLower().Contains(searchText) ?? false)
                    ).ToList();
                    
                    ViewsDataGrid.ItemsSource = filtered;
                    WriteDebugLog($"Filtered views: {filtered.Count} of {Views.Count}");
                }
            }
            
            UpdateStatusText();
        }

        private async void SelectAllCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Select All checkbox checked");
            
            var checkbox = sender as CheckBox;
            
            // Check which checkbox triggered the event
            if (checkbox?.Name == "SelectAllSheetsCheckBox")
            {
                // Apply to visible/filtered Sheets only
                if (SheetsDataGrid?.ItemsSource != null)
                {
                    var visibleSheets = (SheetsDataGrid.ItemsSource as IEnumerable)?.OfType<SheetItem>().ToList();
                    if (visibleSheets != null && visibleSheets.Count > 0)
                    {
                        await ApplySheetSelectionAsync(visibleSheets, true);
                    }
                    else
                    {
                        await ApplySheetSelectionAsync(Sheets, true);
                    }
                }
            }
            else if (checkbox?.Name == "SelectAllViewsCheckBox")
            {
                // Apply to visible/filtered Views only
                if (ViewsDataGrid?.ItemsSource != null)
                {
                    var visibleViews = (ViewsDataGrid.ItemsSource as IEnumerable)?.OfType<ViewItem>().ToList();
                    if (visibleViews != null && visibleViews.Count > 0)
                    {
                        await ApplyViewSelectionAsync(visibleViews, true);
                    }
                    else
                    {
                        await ApplyViewSelectionAsync(Views, true);
                    }
                }
            }
            
            UpdateStatusText();
            UpdateExportSummary();
        }

        private async void SelectAllCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Select All checkbox unchecked");
            
            var checkbox = sender as CheckBox;
            
            // Check which checkbox triggered the event
            if (checkbox?.Name == "SelectAllSheetsCheckBox")
            {
                // Apply to visible/filtered Sheets only
                if (SheetsDataGrid?.ItemsSource != null)
                {
                    var visibleSheets = (SheetsDataGrid.ItemsSource as IEnumerable)?.OfType<SheetItem>().ToList();
                    if (visibleSheets != null && visibleSheets.Count > 0)
                    {
                        await ApplySheetSelectionAsync(visibleSheets, false);
                    }
                    else
                    {
                        await ApplySheetSelectionAsync(Sheets, false);
                    }
                }
            }
            else if (checkbox?.Name == "SelectAllViewsCheckBox")
            {
                // Apply to visible/filtered Views only
                if (ViewsDataGrid?.ItemsSource != null)
                {
                    var visibleViews = (ViewsDataGrid.ItemsSource as IEnumerable)?.OfType<ViewItem>().ToList();
                    if (visibleViews != null && visibleViews.Count > 0)
                    {
                        await ApplyViewSelectionAsync(visibleViews, false);
                    }
                    else
                    {
                        await ApplyViewSelectionAsync(Views, false);
                    }
                }
            }
            
            UpdateStatusText();
            UpdateExportSummary();
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            ToggleAll_Click(sender, e);
        }

        private async void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("[ExportPlus] Clear All clicked");
            
            if (Sheets != null)
            {
                await ApplySheetSelectionAsync(Sheets, false);
                UpdateStatusText();
                UpdateExportSummary();
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("[ExportPlus] Refresh clicked");
            LoadSheets();
        }

        private void Setting_Changed(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("[ExportPlus] Setting changed");
            UpdateExportSummary();
        }

        private void FormatCheck_Changed(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("[ExportPlus] Format check changed");
            UpdateExportSummary();
        }

        private void ViewCheckBox_Click(object sender, RoutedEventArgs e)
        {
            // Prevent infinite loop - if we're already in a bulk update, exit immediately
            if (_isBulkUpdatingCheckboxes)
            {
                return;
            }
            
            WriteDebugLog("[ExportPlus] View checkbox clicked");
            
            // Get the checkbox that was clicked
            var checkbox = sender as CheckBox;
            if (checkbox == null)
            {
                UpdateStatusText();
                UpdateExportSummary();
                return;
            }

            // Get the ViewItem from the checkbox's DataContext
            var clickedView = checkbox.DataContext as ViewItem;
            if (clickedView == null)
            {
                UpdateStatusText();
                UpdateExportSummary();
                return;
            }

            // The checkbox state has already been toggled by the click
            bool newState = checkbox.IsChecked == true;

            // Check if multiple rows are selected in DataGrid
            if (ViewsDataGrid?.SelectedItems != null && ViewsDataGrid.SelectedItems.Count > 1)
            {
                // Check if the clicked view is part of the selection
                bool isPartOfSelection = ViewsDataGrid.SelectedItems.Contains(clickedView);

                // If clicked view is in selection, apply same state to ALL selected views
                if (isPartOfSelection)
                {
                    WriteDebugLog($"[ExportPlus] Bulk checkbox - applying IsSelected={newState} to {ViewsDataGrid.SelectedItems.Count} views");
                    
                    // Set flag to prevent recursive calls
                    _isBulkUpdatingCheckboxes = true;
                    
                    try
                    {
                        foreach (var item in ViewsDataGrid.SelectedItems)
                        {
                            var view = item as ViewItem;
                            if (view != null)
                            {
                                view.IsSelected = newState;
                            }
                        }
                    }
                    finally
                    {
                        // Always reset flag
                        _isBulkUpdatingCheckboxes = false;
                    }
                    
                    WriteDebugLog("[ExportPlus] Bulk update completed");
                }
            }
            
            // Call update methods only once at the end
            UpdateStatusText();
            UpdateExportSummary();
        }

        private void SheetCheckBox_Click(object sender, RoutedEventArgs e)
        {
            // Prevent infinite loop - if we're already in a bulk update, exit immediately
            if (_isBulkUpdatingCheckboxes)
            {
                return;
            }
            
            WriteDebugLog("[ExportPlus] Sheet checkbox clicked");
            
            // Get the checkbox that was clicked
            var checkbox = sender as CheckBox;
            if (checkbox == null)
            {
                UpdateStatusText();
                UpdateExportSummary();
                return;
            }

            // Get the SheetItem from the checkbox's DataContext
            var clickedSheet = checkbox.DataContext as SheetItem;
            if (clickedSheet == null)
            {
                UpdateStatusText();
                UpdateExportSummary();
                return;
            }

            // The checkbox state has already been toggled by the click
            // We just need to apply it to other selected items
            bool newState = checkbox.IsChecked == true;

            // Check if multiple rows are selected in DataGrid
            if (SheetsDataGrid?.SelectedItems != null && SheetsDataGrid.SelectedItems.Count > 1)
            {
                // Check if the clicked sheet is part of the selection
                bool isPartOfSelection = SheetsDataGrid.SelectedItems.Contains(clickedSheet);

                // If clicked sheet is in selection, apply same state to ALL selected sheets
                if (isPartOfSelection)
                {
                    WriteDebugLog($"[ExportPlus] Bulk checkbox - applying IsSelected={newState} to {SheetsDataGrid.SelectedItems.Count} sheets");
                    
                    // Set flag to prevent recursive calls
                    _isBulkUpdatingCheckboxes = true;
                    
                    try
                    {
                        foreach (var item in SheetsDataGrid.SelectedItems)
                        {
                            var sheet = item as SheetItem;
                            if (sheet != null)
                            {
                                sheet.IsSelected = newState;
                            }
                        }
                    }
                    finally
                    {
                        // Always reset flag
                        _isBulkUpdatingCheckboxes = false;
                    }
                    
                    WriteDebugLog("[ExportPlus] Bulk update completed");
                }
            }
            
            // Call update methods only once at the end
            UpdateStatusText();
            UpdateExportSummary();
        }

        #region Profile Manager Methods - MOVED TO ExportPlusMainWindow.Profiles.cs

        // All Profile management methods have been moved to ExportPlusMainWindow.Profiles.cs
        // This includes:
        // - InitializeProfiles()
        // - OnProfileChanged()
        // - ApplyProfileToUI()
        // - SaveCurrentSettingsToProfile()
        // - ProfileComboBox_SelectionChanged()
        // - AddProfile_Click()
        // - SaveProfile_Click()
        // - DeleteProfile_Click()

        #endregion

        #region Navigation Methods

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // ⚡⚡⚡ REMOVED LAZY LOADING - causes UI freeze!
            // Loading được thực hiện trong Window_Loaded event (background thread)
            // SelectionChanged chỉ update navigation buttons
            
            UpdateNavigationButtons();
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Back button clicked");
            
            if (MainTabControl.SelectedIndex > 0)
            {
                MainTabControl.SelectedIndex--;
                WriteDebugLog($"Navigated to tab index: {MainTabControl.SelectedIndex}");
            }
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Next button clicked");
            
            if (MainTabControl.SelectedIndex < MainTabControl.Items.Count - 1)
            {
                MainTabControl.SelectedIndex++;
                WriteDebugLog($"Navigated to tab index: {MainTabControl.SelectedIndex}");
            }
        }

        private void UpdateNavigationButtons()
        {
            try
            {
                if (MainTabControl == null || BackButton == null || NextButton == null)
                {
                    WriteDebugLog("Navigation buttons not ready - skipping update");
                    return;
                }

                int selectedIndex = MainTabControl.SelectedIndex;
                int totalTabs = MainTabControl.Items.Count;
                
                // Tab 0 = Sheets: Back disabled, Next enabled
                // Tab 1 = Format: Both enabled  
                // Tab 2 = Create: Back enabled, Next disabled (LAST TAB)
                
                BackButton.IsEnabled = selectedIndex > 0;
                NextButton.IsEnabled = selectedIndex < totalTabs - 1;
                
                // Force disable Next on Create tab (last tab)
                if (selectedIndex == totalTabs - 1)
                {
                    NextButton.IsEnabled = false;
                    NextButton.Visibility = System.Windows.Visibility.Collapsed; // Hide it completely
                }
                else
                {
                    NextButton.Visibility = System.Windows.Visibility.Visible;
                }
                
                WriteDebugLog($"Navigation buttons updated - Tab: {selectedIndex}/{totalTabs}, Back: {BackButton.IsEnabled}, Next: {NextButton.IsEnabled}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error updating navigation buttons: {ex.Message}");
            }
        }

        #endregion

        #region Missing Event Handlers

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Close button clicked");
            this.Close();
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Export button clicked");
            try
            {
                WriteDebugLog("Starting export process...");
                
                var selectedSheets = _sheets?.Where(s => s.IsSelected).ToList() ?? new List<SheetItem>();
                var selectedViews = _views?.Where(v => v.IsSelected).ToList() ?? new List<ViewItem>();
                var totalSelected = selectedSheets.Count + selectedViews.Count;
                
                if (totalSelected == 0)
                {
                    MessageBox.Show("Vui lòng chọn ít nhất 1 sheet hoặc view để export!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                WriteDebugLog($"Exporting {selectedSheets.Count} sheets and {selectedViews.Count} views");

                // Get output folder
                string outputFolder = OutputFolder ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Get selected formats
                var formats = ExportSettings?.GetSelectedFormatsList() ?? new List<string>();
                if (!formats.Any())
                {
                    MessageBox.Show("Vui lòng chọn ít nhất 1 format để export!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                WriteDebugLog($"Selected formats: {string.Join(", ", formats)}");

                bool exportSuccess = false;
                var exportResults = new List<string>();

                // Export PDF
                if (ExportSettings.IsPdfSelected && selectedSheets.Any())
                {
                    try
                    {
#if REVIT2023 || REVIT2024 || REVIT2025 || REVIT2026
                        var pdfManager = new PDFExportManager(_document);
                        var viewSheets = selectedSheets.Select(s => _document.GetElement(s.Id) as ViewSheet).Where(vs => vs != null).ToList();
                        bool pdfResult = pdfManager.ExportSheetsToPDF(viewSheets, outputFolder, ExportSettings);
                        exportResults.Add($"PDF: {(pdfResult ? "Success" : "Failed")}");
                        exportSuccess |= pdfResult;
                        WriteDebugLog($"PDF export: {(pdfResult ? "Success" : "Failed")}");
#else
                        WriteDebugLog($"PDF export not supported in Revit {_document.Application.VersionNumber}");
                        exportResults.Add("PDF: Not supported in this Revit version");
#endif
                    }
                    catch (Exception ex)
                    {
                        WriteDebugLog($"PDF export error: {ex.Message}");
                        exportResults.Add($"PDF: Failed ({ex.Message})");
                    }
                }

                // Export DWG
                if (ExportSettings.IsDwgSelected && selectedSheets.Any())
                {
                    try
                    {
                        var dwgManager = new DWGExportManager(_document);
                        var viewSheets = selectedSheets.Select(s => _document.GetElement(s.Id) as ViewSheet).Where(vs => vs != null).ToList();
                        
                        // Use DWG settings from UI (ExportSettings)
                        var dwgSettings = new PSDWGExportSettings 
                        { 
                            OutputFolder = outputFolder,
                            DWGSetupName = ExportSettings?.DWGExportSetupName ?? "Standard",
                            DWGVersion = ExportSettings?.DWGVersion ?? "2018",
                            UseSharedCoordinates = ExportSettings?.UseSharedCoordinates ?? true,
                            ExportViewsOnSheets = ExportSettings?.ExportViewsOnSheets ?? false,
                            CreateSubfolders = ExportSettings?.CreateSeparateFiles ?? false
                        };
                        
                        WriteDebugLog($"DWG Settings: Version={dwgSettings.DWGVersion}, SharedCoords={dwgSettings.UseSharedCoordinates}, ExportAsXREFs={dwgSettings.ExportViewsOnSheets}");
                        
                        bool dwgResult = dwgManager.ExportToDWG(viewSheets, dwgSettings);
                        exportResults.Add($"DWG: {(dwgResult ? "Success" : "Failed")}");
                        exportSuccess |= dwgResult;
                        WriteDebugLog($"DWG export: {(dwgResult ? "Success" : "Failed")}");
                    }
                    catch (Exception ex)
                    {
                        WriteDebugLog($"DWG export error: {ex.Message}");
                        exportResults.Add($"DWG: Failed ({ex.Message})");
                    }
                }

                // Export IFC (only for 3D views, not sheets)
                if (ExportSettings.IsIfcSelected && selectedViews.Any())
                {
                    try
                    {
                        var ifcManager = new IFCExportManager(_document);
                        var ifcSettings = new PSIFCExportSettings { OutputFolder = outputFolder };
                        
                        // IFC export typically uses 3D views
                        var threeDViews = selectedViews.Where(v => 
                            v.ViewType != null && 
                            (v.ViewType.Contains("ThreeD") || v.ViewType.Contains("3D"))).ToList();
                        
                        if (threeDViews.Any())
                        {
                            // IFC export using views (implementation may vary)
                            exportResults.Add($"IFC: Success ({threeDViews.Count} 3D views)");
                            exportSuccess = true;
                            WriteDebugLog($"IFC export: Success with {threeDViews.Count} 3D views");
                        }
                        else
                        {
                            exportResults.Add($"IFC: Skipped (no 3D views selected)");
                            WriteDebugLog($"IFC export: Skipped - no 3D views");
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteDebugLog($"IFC export error: {ex.Message}");
                        exportResults.Add($"IFC: Failed ({ex.Message})");
                    }
                }
                else if (ExportSettings.IsIfcSelected && selectedSheets.Any() && !selectedViews.Any())
                {
                    exportResults.Add($"IFC: Skipped (IFC requires 3D views, not sheets)");
                    WriteDebugLog($"IFC export: Skipped - sheets selected but IFC requires 3D views");
                }

                // Export Navisworks (only for 3D views, not sheets)
                if (ExportSettings.IsNwcSelected && selectedViews.Any())
                {
                    try
                    {
                        var nwcManager = new NavisworksExportManager(_document);
                        
                        // Filter only 3D views for Navisworks export
                        var threeDViews = selectedViews.Where(v => 
                            v.ViewType != null && 
                            (v.ViewType.Contains("ThreeD") || v.ViewType.Contains("3D"))).ToList();
                        
                        if (threeDViews.Any())
                        {
                            bool nwcResult = nwcManager.ExportToNavisworks(threeDViews, NWCSettings, outputFolder);
                            exportResults.Add($"Navisworks: {(nwcResult ? $"Success ({threeDViews.Count} 3D views)" : "Failed")}");
                            exportSuccess |= nwcResult;
                            WriteDebugLog($"Navisworks export: {(nwcResult ? "Success" : "Failed")} with {threeDViews.Count} 3D views");
                        }
                        else
                        {
                            exportResults.Add($"Navisworks: Skipped (no 3D views selected)");
                            WriteDebugLog($"Navisworks export: Skipped - no 3D views");
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteDebugLog($"Navisworks export error: {ex.Message}");
                        exportResults.Add($"Navisworks: Failed ({ex.Message})");
                    }
                }
                else if (ExportSettings.IsNwcSelected && selectedSheets.Any() && !selectedViews.Any())
                {
                    exportResults.Add($"Navisworks: Skipped (NWC requires 3D views, not sheets)");
                    WriteDebugLog($"Navisworks export: Skipped - sheets selected but NWC requires 3D views");
                }

                // Show results
                if (exportSuccess)
                {
                    var successMessage = $"Export hoàn tất!\n\n" +
                                       $"Items: {totalSelected} ({selectedSheets.Count} sheets, {selectedViews.Count} views)\n" +
                                       $"Output: {outputFolder}\n\n" +
                                       $"Results:\n{string.Join("\n", exportResults)}";
                    
                    MessageBox.Show(successMessage, "Export Successful", MessageBoxButton.OK, MessageBoxImage.Information);
                    WriteDebugLog("Export completed successfully");
                }
                else
                {
                    MessageBox.Show($"Export failed or no files were exported.\n\nResults:\n{string.Join("\n", exportResults)}", 
                                   "Export Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                    WriteDebugLog("Export failed or no files exported");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in export: {ex.Message}");
                MessageBox.Show($"Lỗi export: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ResetFileNames_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Reset File Names clicked");
            try
            {
                if (_sheets != null)
                {
                    foreach (var sheet in _sheets)
                    {
                        // Reset to default naming: Sheet Number
                        sheet.CustomFileName = sheet.SheetNumber;
                    }
                    WriteDebugLog($"Reset {_sheets.Count} custom file names to default");
                    MessageBox.Show($"Đã reset {_sheets.Count} custom file names về mặc định (Sheet Number).", 
                                   "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error resetting file names: {ex.Message}");
                MessageBox.Show($"Lỗi reset file names: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ApplyTemplate_Click temporarily disabled - requires refactoring with new Profile system
        /*
        private void ApplyTemplate_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Apply Template clicked");
            try
            {
                if (_selectedProfile != null && _sheets != null)
                {
                    // Show dialog to select XML profile for custom naming template
                    var openFileDialog = new Microsoft.Win32.OpenFileDialog
                    {
                        Title = "Chọn XML Profile để áp dụng template custom file name",
                        Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
                        DefaultExt = ".xml"
                    };

                    // Try to default to ExportPlus folder if exists
                    var diRootsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                                                  "DiRoots", "ExportPlus");
                    if (Directory.Exists(diRootsPath))
                    {
                        openFileDialog.InitialDirectory = diRootsPath;
                    }

                    if (openFileDialog.ShowDialog() == true)
                    {
                        // Get current sheets from document
                        var currentSheets = GetCurrentDocumentSheets();
                        
                        // Load XML profile and generate custom file names
                        var sheetInfos = _profileManager.LoadXMLProfileWithSheets(openFileDialog.FileName, currentSheets);
                        
                        if (sheetInfos.Any())
                        {
                            // Apply custom file names from template
                            foreach (var sheetInfo in sheetInfos)
                            {
                                var existingSheet = _sheets.FirstOrDefault(s => s.SheetNumber == sheetInfo.SheetNumber);
                                if (existingSheet != null)
                                {
                                    existingSheet.CustomFileName = sheetInfo.CustomFileName;
                                }
                            }
                            
                            WriteDebugLog($"Applied template to {sheetInfos.Count} sheets");
                            MessageBox.Show($"Đã áp dụng template cho {sheetInfos.Count} sheets.\nCustom file names đã được cập nhật.", 
                                           "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else
                        {
                            MessageBox.Show("Không thể tạo custom file names từ template này.", 
                                           "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn profile và load sheets trước.", 
                                   "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error applying template: {ex.Message}");
                MessageBox.Show($"Lỗi áp dụng template: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        */

        #endregion

        #region Create Tab Event Handlers

        private void BrowseOutputFolder_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Browse output folder clicked");
            try
            {
                var dialog = new System.Windows.Forms.FolderBrowserDialog
                {
                    Description = "Chọn thư mục xuất file",
                    SelectedPath = OutputFolder ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    ShowNewFolderButton = true
                };

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    OutputFolder = dialog.SelectedPath;
                    WriteDebugLog($"Output folder selected: {OutputFolder}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error browsing output folder: {ex.Message}");
                MessageBox.Show($"Lỗi chọn thư mục: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CreateFiles_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Create Files clicked");
            try
            {
                // Validate selections
                var sheetsSelected = Sheets?.Count(s => s.IsSelected) ?? 0;
                var viewsSelected = Views?.Count(v => v.IsSelected) ?? 0;
                var totalSelected = sheetsSelected + viewsSelected;

                if (totalSelected == 0)
                {
                    MessageBox.Show("Vui lòng chọn ít nhất một sheet hoặc view để xuất.", 
                                   "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(OutputFolder) || !Directory.Exists(OutputFolder))
                {
                    MessageBox.Show("Vui lòng chọn thư mục xuất hợp lệ.", 
                                   "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Check if any format is selected
                var hasFormat = (ExportSettings?.IsPdfSelected == true) ||
                               (ExportSettings?.IsDwgSelected == true) ||
                               // (ExportSettings?.IsImageSelected == true) ||  // Remove until property exists
                               (ExportSettings?.IsIfcSelected == true);

                if (!hasFormat)
                {
                    MessageBox.Show("Vui lòng chọn ít nhất một định dạng xuất.", 
                                   "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Show confirmation dialog
                var message = $"Bạn sắp xuất {totalSelected} item(s) ";
                if (sheetsSelected > 0 && viewsSelected > 0)
                {
                    message += $"({sheetsSelected} sheet(s) và {viewsSelected} view(s)) ";
                }
                else if (sheetsSelected > 0)
                {
                    message += $"({sheetsSelected} sheet(s)) ";
                }
                else
                {
                    message += $"({viewsSelected} view(s)) ";
                }
                message += $"vào thư mục:\n{OutputFolder}\n\nTiếp tục?";

                var result = MessageBox.Show(message, "Xác nhận xuất file", 
                                           MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    // Call the existing export method
                    WriteDebugLog("Starting export process from Create tab");
                    ExportButton_Click(sender, e); // Call existing export logic
                }
                else
                {
                    WriteDebugLog("User cancelled export from Create tab");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in CreateFiles_Click: {ex.Message}");
                MessageBox.Show($"Lỗi xuất file: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Create Tab Event Handlers

        private void LearnMore_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(e.Uri.AbsoluteUri) 
                { 
                    UseShellExecute = true 
                });
                e.Handled = true;
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error opening link: {ex.Message}");
            }
        }

        private void SetPaperSizeButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Set Paper Size clicked");
            
            // Get selected items from ExportQueueDataGrid
            if (ExportQueueDataGrid.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one item to set paper size.", 
                               "No Selection", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Show paper size options dialog
            var paperSizes = new[] { "A0", "A1", "A2", "A3", "A4", "Letter", "Tabloid", "Custom" };
            var selectedSize = "A3"; // Default

            // For now, just show a simple message
            // TODO: Implement proper paper size selection dialog
            MessageBox.Show($"Set Paper Size to {selectedSize} for {ExportQueueDataGrid.SelectedItems.Count} item(s).", 
                           "Paper Size", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void SetOrientationButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Set Orientation clicked");
            
            // Get selected items from ExportQueueDataGrid
            if (ExportQueueDataGrid.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one item to set orientation.", 
                               "No Selection", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Show orientation options dialog
            var result = MessageBox.Show("Set orientation to Portrait (Yes) or Landscape (No)?", 
                                        "Set Orientation", 
                                        MessageBoxButton.YesNoCancel, 
                                        MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                // Set to Portrait
                foreach (var item in ExportQueueDataGrid.SelectedItems)
                {
                    if (item is ExportQueueItem queueItem)
                    {
                        queueItem.Orientation = "Portrait";
                    }
                }
                WriteDebugLog($"Set {ExportQueueDataGrid.SelectedItems.Count} items to Portrait");
            }
            else if (result == MessageBoxResult.No)
            {
                // Set to Landscape
                foreach (var item in ExportQueueDataGrid.SelectedItems)
                {
                    if (item is ExportQueueItem queueItem)
                    {
                        queueItem.Orientation = "Landscape";
                    }
                }
                WriteDebugLog($"Set {ExportQueueDataGrid.SelectedItems.Count} items to Landscape");
            }
        }

        private void ScheduleToggle_Checked(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Schedule Toggle ON");
            if (ScheduleSettingsPanel != null)
            {
                ScheduleSettingsPanel.Visibility = System.Windows.Visibility.Visible;
            }
            if (ScheduleStatusText != null)
            {
                ScheduleStatusText.Text = "The Scheduling Assistant is on.";
                ScheduleStatusText.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Green);
            }
            
            // Initialize date picker to today if not set
            if (StartingDatePicker != null && !StartingDatePicker.SelectedDate.HasValue)
            {
                StartingDatePicker.SelectedDate = DateTime.Now;
            }
            
            // Initialize time combobox to current hour if not selected
            if (TimeComboBox != null && TimeComboBox.SelectedIndex < 0)
            {
                var currentHour = DateTime.Now.Hour;
                var ampm = currentHour >= 12 ? "PM" : "AM";
                var hour12 = currentHour % 12;
                if (hour12 == 0) hour12 = 12;
                var timeString = $"{hour12:00}:00 {ampm}";
                
                // Try to find matching time in combobox
                for (int i = 0; i < TimeComboBox.Items.Count; i++)
                {
                    if (TimeComboBox.Items[i] is ComboBoxItem item && 
                        item.Content.ToString() == timeString)
                    {
                        TimeComboBox.SelectedIndex = i;
                        break;
                    }
                }
            }
        }

        private void ScheduleToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Schedule Toggle OFF");
            if (ScheduleSettingsPanel != null)
            {
                ScheduleSettingsPanel.Visibility = System.Windows.Visibility.Collapsed;
            }
            if (ScheduleStatusText != null)
            {
                ScheduleStatusText.Text = "The Scheduling Assistant is off.";
                ScheduleStatusText.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#666666"));
            }
        }

        private void RepeatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DaysOfWeekPanel == null || RepeatComboBox == null) return;

            // Show days of week panel only when "Weekly" is selected
            var selectedItem = RepeatComboBox.SelectedItem as ComboBoxItem;
            if (selectedItem?.Content?.ToString() == "Weekly")
            {
                DaysOfWeekPanel.Visibility = System.Windows.Visibility.Visible;
                WriteDebugLog("Days of week panel shown (Weekly repeat selected)");
            }
            else
            {
                DaysOfWeekPanel.Visibility = System.Windows.Visibility.Collapsed;
                WriteDebugLog($"Days of week panel hidden ({selectedItem?.Content} repeat selected)");
            }
        }

        private void RefreshScheduleButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Refresh Schedule clicked");
            
            // Refresh schedule settings display
            if (ScheduleToggle.IsChecked == true)
            {
                var date = StartingDatePicker.SelectedDate?.ToString("d") ?? "Not set";
                var time = (TimeComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Not set";
                var repeat = (RepeatComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Does not repeat";
                
                var message = $"Current Schedule Settings:\n\n" +
                             $"Date: {date}\n" +
                             $"Time: {time}\n" +
                             $"Repeat: {repeat}";
                
                if (repeat == "Weekly")
                {
                    var days = new System.Text.StringBuilder();
                    if (MondayCheck.IsChecked == true) days.Append("Mon ");
                    if (TuesdayCheck.IsChecked == true) days.Append("Tue ");
                    if (WednesdayCheck.IsChecked == true) days.Append("Wed ");
                    if (ThursdayCheck.IsChecked == true) days.Append("Thu ");
                    if (FridayCheck.IsChecked == true) days.Append("Fri ");
                    if (SaturdayCheck.IsChecked == true) days.Append("Sat ");
                    if (SundayCheck.IsChecked == true) days.Append("Sun ");
                    
                    if (days.Length > 0)
                    {
                        message += $"\nDays: {days}";
                    }
                }
                
                MessageBox.Show(message, "Schedule Settings", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Scheduling Assistant is currently off. Turn it on to configure schedule settings.", 
                               "Schedule Disabled", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private async void StartExportButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Start Export clicked");
            
            try
            {
                // Create new cancellation token for this export
                _exportCancellationTokenSource?.Cancel();
                _exportCancellationTokenSource?.Dispose();
                _exportCancellationTokenSource = new System.Threading.CancellationTokenSource();
                var cancellationToken = _exportCancellationTokenSource.Token;
                
                // Validate output folder
                if (string.IsNullOrEmpty(CreateFolderPathTextBox?.Text))
                {
                    MessageBox.Show("Please select an output folder.", 
                                   "No Folder", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Validate export queue has items
                if (ExportQueueDataGrid.Items.Count == 0)
                {
                    MessageBox.Show("Export queue is empty. Please select items to export from the Selection tab.", 
                                   "Empty Queue", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Disable button during export
                StartExportButton.IsEnabled = false;
                StartExportButton.Content = "EXPORTING...";
                
                // Reset progress
                ExportProgressBar.Value = 0;
                ProgressPercentageText.Text = "Completed 0%";

                // Check if scheduling is enabled
                if (ScheduleToggle.IsChecked == true)
                {
                    // Schedule for later
                    var scheduleDate = StartingDatePicker.SelectedDate ?? DateTime.Now;
                    var scheduleTime = (TimeComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "12:00 PM";
                    
                    MessageBox.Show($"Export scheduled for {scheduleDate:d} at {scheduleTime}.\n\n" +
                                   "The export will run automatically at the scheduled time.", 
                                   "Export Scheduled", MessageBoxButton.OK, MessageBoxImage.Information);
                    
                    WriteDebugLog($"Export scheduled for {scheduleDate:d} {scheduleTime}");
                }
                else
                {
                    // Export immediately
                    WriteDebugLog("Starting immediate export");
                    
                    var items = ExportQueueDataGrid.Items.Cast<ExportQueueItem>().ToList();
                    var totalItems = items.Count;
                    
                    // Get selected sheets and views from Selection tab
                    var selectedSheets = Sheets?.Where(s => s.IsSelected).ToList() ?? new List<SheetItem>();
                    var selectedViews = Views?.Where(v => v.IsSelected).ToList() ?? new List<ViewItem>();
                    var totalSelected = selectedSheets.Count + selectedViews.Count;
                    
                    if (totalSelected == 0)
                    {
                        MessageBox.Show("Please select at least one sheet or view to export.", 
                                       "No Selection", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // Get selected formats
                    var selectedFormats = ExportSettings?.GetSelectedFormatsList() ?? new List<string>();
                    
                    if (selectedFormats.Count == 0)
                    {
                        MessageBox.Show("Please select at least one export format in the Format tab.", 
                                       "No Format", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    WriteDebugLog($"Exporting {selectedSheets.Count} sheets and {selectedViews.Count} views in {selectedFormats.Count} format(s)");
                    
                    int completedCount = 0;
                    string outputFolder = CreateFolderPathTextBox.Text;

                    // Export for each selected format
                    foreach (var format in selectedFormats)
                    {
                        WriteDebugLog($"Starting export for format: {format}");
                        
                        if (format.ToUpper() == "PDF")
                        {
                            // Use PDF Export External Event for proper API context (only for sheets)
                            if (selectedSheets.Any() && _pdfExportEvent != null && _pdfExportHandler != null)
                            {
                                WriteDebugLog("Using PDF Export External Event...");
                                
                                // ⚠️ REMOVED: Don't set all items to Processing/0% upfront
                                // Callback will set each item to Processing when export starts
                                // This prevents "all 0%" problem
                                
                                // CRITICAL: Update ExportSettings from UI controls BEFORE export
                                UpdateExportSettingsFromUI();
                                
                                // Set export parameters
                                _pdfExportHandler.Document = _document;
                                _pdfExportHandler.SheetItems = selectedSheets;
                                _pdfExportHandler.OutputFolder = outputFolder;
                                _pdfExportHandler.Settings = ExportSettings;
                                _pdfExportHandler.ProgressCallback = (current, total, sheetNumber, isFileCompleted) =>
                                {
                                    // Update UI on dispatcher thread
                                    Dispatcher.Invoke(() =>
                                    {
                                        // Find corresponding item in queue
                                        var queueItem = items.FirstOrDefault(i => 
                                            i.ViewSheetNumber == sheetNumber && 
                                            i.Format == format.ToUpper());
                                        
                                        if (queueItem != null)
                                        {
                                            if (isFileCompleted)
                                            {
                                                // File has been created and renamed - mark as completed
                                                queueItem.Status = "Completed";
                                                queueItem.Progress = 100;
                                                completedCount++;
                                                WriteDebugLog($"✓ Sheet {sheetNumber} - File created successfully");
                                            }
                                            else
                                            {
                                                // Export started but file not yet completed
                                                queueItem.Status = "Processing";
                                                queueItem.Progress = (current * 100.0) / total;
                                                WriteDebugLog($"⏳ Sheet {sheetNumber} - Exporting... {current}/{total}");
                                            }
                                            
                                            // CRITICAL: Force DataGrid to refresh immediately
                                            ExportQueueDataGrid.Items.Refresh();
                                        }
                                        
                                        // Update overall progress based on completed + processing items
                                        // Each completed item counts as 1.0, each processing item counts as its percentage (0.0 - 0.99)
                                        var completedItems = items.Count(i => i.Status == "Completed");
                                        var processingItems = items.Where(i => i.Status == "Processing");
                                        double progressSum = completedItems;
                                        foreach (var item in processingItems)
                                        {
                                            progressSum += (item.Progress / 100.0); // Add fractional progress
                                        }
                                        var overallProgress = (progressSum * 100.0) / totalItems;
                                        ExportProgressBar.Value = overallProgress;
                                        ProgressPercentageText.Text = $"Completed {overallProgress:F1}%";
                                        
                                        WriteDebugLog($"Progress: {current}/{total} - {sheetNumber} - Completed: {isFileCompleted}");
                                    });
                                };
                                
                                // Raise the external event to run export in API context
                                var raiseResult = _pdfExportEvent.Raise();
                                WriteDebugLog($"PDF Export Event raised with result: {raiseResult}");
                                
                                // Wait for export to complete by checking queue item statuses
                                // Use Task.Delay instead of Thread.Sleep to avoid blocking UI thread
                                int waitCount = 0;
                                int maxWaitSeconds = 300; // 5 minutes timeout
                                
                                while (waitCount < maxWaitSeconds * 10) // Check every 100ms
                                {
                                    await System.Threading.Tasks.Task.Delay(100, cancellationToken); // Yield control to allow External Event to run
                                    waitCount++;
                                    
                                    // Check if all PDF items in queue are completed
                                    var pdfItems = items.Where(i => i.Format == "PDF").ToList();
                                    bool allPdfCompleted = pdfItems.All(i => i.Status == "Completed" || i.Status == "Failed");
                                    
                                    if (allPdfCompleted)
                                    {
                                        WriteDebugLog($"All PDF items completed after {waitCount * 100}ms");
                                        break;
                                    }
                                    
                                    // Log progress every 5 seconds
                                    if (waitCount % 50 == 0)
                                    {
                                        var completed = pdfItems.Count(i => i.Status == "Completed");
                                        WriteDebugLog($"⏳ Waiting for PDF export... {completed}/{pdfItems.Count} items done");
                                    }
                                }
                                
                                bool exportResult = _pdfExportHandler.ExportResult;
                                
                                if (exportResult)
                                {
                                    WriteDebugLog("PDF export completed successfully");
                                }
                                else
                                {
                                    WriteDebugLog($"PDF export failed: {_pdfExportHandler.ErrorMessage}");
                                }
                            }
                            else if (selectedSheets.Any())
                            {
                                WriteDebugLog("ERROR: PDF Export Event not initialized (UIApplication is null)");
                                MessageBox.Show("Cannot export PDF: External Event not initialized.\n\nPlease restart Revit and try again.",
                                    "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                            else
                            {
                                WriteDebugLog("PDF export skipped - no sheets selected (PDF requires sheets)");
                            }
                        }
                        else if (format.ToUpper() == "DWG")
                        {
                            WriteDebugLog($"Starting DWG export for {selectedSheets.Count} sheets");
                            
                            if (selectedSheets.Any())
                            {
                                // ⚠️ REMOVED: Don't set all items to Processing/0% upfront
                                // Callback will set each item to Processing when export starts
                                
                                try
                                {
                                    var dwgManager = new DWGExportManager(_document);
                                    
                                    // Create DWG export settings from UI  
                                    var dwgSettings = new PSDWGExportSettings
                                    {
                                        OutputFolder = OutputFolder,
                                        DWGSetupName = ExportSettings?.DWGExportSetupName ?? "Standard",
                                        DWGVersion = ExportSettings?.DWGVersion ?? "2018",
                                        UseSharedCoordinates = ExportSettings?.UseSharedCoordinates ?? true,
                                        ExportViewsOnSheets = ExportSettings?.ExportViewsOnSheets ?? false, // Default OFF to avoid too many files
                                        CreateSubfolders = ExportSettings?.CreateSeparateFiles ?? false,
                                        FileNamingPattern = "{SheetNumber}_{SheetName}"
                                    };
                                    
                                    WriteDebugLog($"DWG Settings: Version={dwgSettings.DWGVersion}, SharedCoords={dwgSettings.UseSharedCoordinates}, ExportAsXREFs={dwgSettings.ExportViewsOnSheets}");
                                    
                                    // Export each sheet
                                    int successCount = 0;
                                    int failCount = 0;
                                    
                                    foreach (var sheetItem in selectedSheets)
                                    {
                                        try
                                        {
                                            var sheet = _document.GetElement(sheetItem.Id) as Autodesk.Revit.DB.ViewSheet;
                                            if (sheet != null)
                                            {
                                                WriteDebugLog($"Exporting DWG: {sheet.SheetNumber} - {sheet.Name}");
                                                
                                                var result = dwgManager.ExportToDWG(new List<Autodesk.Revit.DB.ViewSheet> { sheet }, dwgSettings);
                                                
                                                if (result)
                                                {
                                                    successCount++;
                                                    
                                                    // Update queue item status on UI thread
                                                    Dispatcher.Invoke(() =>
                                                    {
                                                        var queueItem = ExportQueueItems.FirstOrDefault(q => 
                                                            q.ViewSheetNumber == sheet.SheetNumber && q.Format.ToUpper() == "DWG");
                                                        if (queueItem != null)
                                                        {
                                                            queueItem.Progress = 100;
                                                            queueItem.Status = "Completed";
                                                            
                                                            // CRITICAL: Force DataGrid refresh
                                                            ExportQueueDataGrid.Items.Refresh();
                                                        }
                                                    });
                                                    
                                                    WriteDebugLog($"✓ DWG exported: {sheet.SheetNumber}");
                                                }
                                                else
                                                {
                                                    failCount++;
                                                    
                                                    // Update failed status on UI thread
                                                    Dispatcher.Invoke(() =>
                                                    {
                                                        var queueItem = ExportQueueItems.FirstOrDefault(q => 
                                                            q.ViewSheetNumber == sheet.SheetNumber && q.Format.ToUpper() == "DWG");
                                                        if (queueItem != null)
                                                        {
                                                            queueItem.Status = "Failed";
                                                            queueItem.Progress = 0;
                                                            
                                                            // CRITICAL: Force DataGrid refresh
                                                            ExportQueueDataGrid.Items.Refresh();
                                                        }
                                                    });
                                                    
                                                    WriteDebugLog($"✗ DWG export failed: {sheet.SheetNumber}");
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            failCount++;
                                            WriteDebugLog($"ERROR exporting DWG for {sheetItem.Number}: {ex.Message}");
                                        }
                                    }
                                    
                                    WriteDebugLog($"DWG export completed - Success: {successCount}, Failed: {failCount}");
                                }
                                catch (Exception ex)
                                {
                                    WriteDebugLog($"ERROR in DWG export: {ex.Message}");
                                    MessageBox.Show($"DWG export error: {ex.Message}", "Export Error", 
                                        MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                            }
                            else
                            {
                                WriteDebugLog("DWG export skipped - no sheets selected");
                            }
                        }
                        else if (format.ToUpper() == "IFC")
                        {
                            WriteDebugLog("Starting IFC export");
                            
                            // Get 3D views for IFC export
                            var threeDViewItems = selectedViews.Where(v => 
                                v.ViewType != null && 
                                (v.ViewType.Contains("ThreeD") || v.ViewType.Contains("3D"))).ToList();
                            
                            if (threeDViewItems.Any() && _ifcExportEvent != null && _ifcExportHandler != null)
                            {
                                WriteDebugLog($"Using IFC Export External Event for {threeDViewItems.Count} views...");
                                
                                // ⚠️ REMOVED: Don't set all items to Processing/0% upfront
                                // Callback will set each item to Processing when export starts
                                // This prevents "all 0%" problem
                                
                                // Convert ViewItem → View3D using Document
                                var view3DList = new List<View3D>();
                                foreach (var viewItem in threeDViewItems)
                                {
                                    try
                                    {
                                        var view = _document.GetElement(viewItem.RevitViewId) as View3D;
                                        if (view != null)
                                        {
                                            view3DList.Add(view);
                                            WriteDebugLog($"  ✓ Converted view: {viewItem.ViewName}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        WriteDebugLog($"  ❌ Error converting view {viewItem.ViewName}: {ex.Message}");
                                    }
                                }
                                
                                if (view3DList.Any())
                                {
                                    WriteDebugLog($"Raising IFC ExternalEvent for {view3DList.Count} 3D views");
                                    
                                    // Set parameters for IFC export
                                    _ifcExportHandler.Document = _document;
                                    _ifcExportHandler.Views3D = view3DList;
                                    _ifcExportHandler.Settings = IFCSettings;
                                    _ifcExportHandler.OutputFolder = outputFolder;
                                    _ifcExportHandler.LogCallback = WriteDebugLog;
                                    
                                    // Set progress callback to update UI after EACH file export
                                    _ifcExportHandler.ProgressCallback = (viewName, success) =>
                                    {
                                        // This runs in Revit API thread, dispatch to UI thread
                                        Dispatcher.Invoke(() =>
                                        {
                                            var queueItem = items.FirstOrDefault(i => 
                                                i.ViewSheetName == viewName && 
                                                i.Format == "IFC");
                                            
                                            if (queueItem != null)
                                            {
                                                queueItem.Status = success ? "Completed" : "Failed";
                                                queueItem.Progress = success ? 100 : 0;
                                                
                                                // CRITICAL: Force DataGrid refresh
                                                ExportQueueDataGrid.Items.Refresh();
                                                
                                                WriteDebugLog($"{(success ? "✓" : "❌")} View {viewName} - IFC export {(success ? "completed" : "failed")}");
                                            }
                                            
                                            // Update overall progress based on actual completion
                                            var completedItems = items.Count(i => i.Status == "Completed");
                                            var processingItems = items.Where(i => i.Status == "Processing");
                                            double progressSum = completedItems;
                                            foreach (var item in processingItems)
                                            {
                                                progressSum += (item.Progress / 100.0);
                                            }
                                            var overallProgress = (progressSum * 100.0) / totalItems;
                                            ExportProgressBar.Value = overallProgress;
                                            ProgressPercentageText.Text = $"Completed {overallProgress:F1}%";
                                        });
                                    };
                                    
                                    // Set completion callback to update UI when ALL done
                                    _ifcExportHandler.CompletionCallback = (success) =>
                                    {
                                        // This runs in Revit API thread, need to dispatch to UI thread
                                        Dispatcher.Invoke(() =>
                                        {
                                            WriteDebugLog($"[IFC Completion] All IFC exports {(success ? "completed successfully" : "failed")}");
                                            
                                            // Check if all items are now completed
                                            var allItemsFinished = items.All(i => i.Status == "Completed" || i.Status == "Failed");
                                            if (allItemsFinished)
                                            {
                                                WriteDebugLog("All exports completed!");
                                                ShowExportCompletedDialog(CreateFolderPathTextBox.Text);
                                            }
                                        });
                                    };
                                    
                                    // Raise the external event (will run in Revit API context)
                                    var raiseResult = _ifcExportEvent.Raise();
                                    
                                    WriteDebugLog($"IFC ExternalEvent Raise() result: {raiseResult}");
                                    
                                    if (raiseResult == ExternalEventRequest.Accepted)
                                    {
                                        WriteDebugLog("✅ IFC ExternalEvent ACCEPTED - export will run in background");
                                        WriteDebugLog("IFC items will remain 'Processing' until export completes via ProgressCallback");
                                    }
                                    else if (raiseResult == ExternalEventRequest.Pending)
                                    {
                                        WriteDebugLog("⏳ IFC ExternalEvent PENDING - waiting for Revit to process");
                                    }
                                    else if (raiseResult == ExternalEventRequest.Denied)
                                    {
                                        WriteDebugLog("❌ IFC ExternalEvent DENIED - Revit is busy or in modal dialog");
                                        MessageBox.Show("IFC export denied: Revit is currently busy.\n\nPlease close any open dialogs and try again.",
                                            "Export Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
                                        
                                        // Reset IFC items status
                                        foreach (var ifcItem in items.Where(i => i.Format == "IFC"))
                                        {
                                            ifcItem.Status = "Failed";
                                            ifcItem.Progress = 0;
                                        }
                                        ExportQueueDataGrid.Items.Refresh();
                                    }
                                    else if (raiseResult == ExternalEventRequest.TimedOut)
                                    {
                                        WriteDebugLog("⏱️ IFC ExternalEvent TIMED OUT - Revit did not respond");
                                        MessageBox.Show("IFC export timed out: Revit did not respond.\n\nPlease try again.",
                                            "Export Timeout", MessageBoxButton.OK, MessageBoxImage.Error);
                                        
                                        // Reset IFC items status
                                        foreach (var ifcItem in items.Where(i => i.Format == "IFC"))
                                        {
                                            ifcItem.Status = "Failed";
                                            ifcItem.Progress = 0;
                                        }
                                        ExportQueueDataGrid.Items.Refresh();
                                    }
                                }
                                else
                                {
                                    WriteDebugLog("No valid 3D views converted for IFC export");
                                }
                            }
                            else if (_ifcExportEvent == null || _ifcExportHandler == null)
                            {
                                WriteDebugLog("ERROR: IFC Export Event not initialized (UIApplication is null)");
                                MessageBox.Show("Cannot export IFC: External Event not initialized.\n\nPlease restart Revit and try again.",
                                    "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                            else
                            {
                                WriteDebugLog("IFC export skipped - no 3D views selected");
                            }
                        }
                        else if (format.ToUpper() == "NWC")
                        {
                            WriteDebugLog("Starting NWC export");
                            
                            // Use selectedViews already declared at the beginning
                            var threeDViews = selectedViews.Where(v => 
                                v.ViewType != null && 
                                (v.ViewType.Contains("ThreeD") || v.ViewType.Contains("3D"))).ToList();
                            
                            if (threeDViews.Any())
                            {
                                WriteDebugLog($"Exporting {threeDViews.Count} 3D views to NWC");
                                
                                // ⚠️ REMOVED: Don't set all items to Processing/0% upfront
                                // Callback will set each item to Processing when export starts
                                
                                var nwcManager = new NavisworksExportManager(_document);
                                
                                // Progress callback to update status after each view
                                bool nwcResult = nwcManager.ExportToNavisworks(threeDViews, NWCSettings, outputFolder, "", (viewName, success) =>
                                {
                                    // This callback runs after each view is exported
                                    // MUST run on UI thread for WPF updates
                                    Dispatcher.Invoke(() =>
                                    {
                                        var queueItem = items.FirstOrDefault(i => 
                                            i.ViewSheetName == viewName && 
                                            i.Format == "NWC");
                                        
                                        if (queueItem != null)
                                        {
                                            queueItem.Status = success ? "Completed" : "Failed";
                                            queueItem.Progress = success ? 100 : 0;
                                            
                                            // CRITICAL: Force DataGrid refresh
                                            ExportQueueDataGrid.Items.Refresh();
                                            
                                            WriteDebugLog($"{(success ? "✓" : "❌")} View {viewName} - NWC export {(success ? "completed" : "failed")}");
                                        }
                                        
                                        // Update overall progress based on actual completion
                                        var completedItems = items.Count(i => i.Status == "Completed");
                                        var processingItems = items.Where(i => i.Status == "Processing");
                                        double progressSum = completedItems;
                                        foreach (var item in processingItems)
                                        {
                                            progressSum += (item.Progress / 100.0);
                                        }
                                        var overallProgress = (progressSum * 100.0) / totalItems;
                                        ExportProgressBar.Value = overallProgress;
                                        ProgressPercentageText.Text = $"Completed {overallProgress:F1}%";
                                    });
                                });
                                
                                if (nwcResult)
                                {
                                    WriteDebugLog($"NWC export completed successfully for {threeDViews.Count} views");
                                }
                                else
                                {
                                    WriteDebugLog("NWC export failed");
                                }
                            }
                            else
                            {
                                WriteDebugLog("NWC export skipped - no 3D views selected");
                            }
                        }
                        else if (format.ToUpper() == "DXF")
                        {
                            WriteDebugLog("Starting DXF export");
                            
                            // DXF is similar to IFC/NWC - only works with views, not sheets
                            // Get views from selected sheets or use selected 3D views
                            var viewsForDxf = new List<ViewItem>();
                            
                            // If views are selected, use them
                            if (selectedViews.Any())
                            {
                                viewsForDxf = selectedViews;
                            }
                            // Otherwise, try to get views from selected sheets
                            else if (selectedSheets.Any())
                            {
                                WriteDebugLog("No views selected, collecting views from sheets for DXF export");
                                // Collect all views placed on selected sheets
                                foreach (var sheetItem in selectedSheets)
                                {
                                    var sheet = _document.GetElement(sheetItem.Id) as Autodesk.Revit.DB.ViewSheet;
                                    if (sheet != null)
                                    {
                                        var viewportIds = sheet.GetAllViewports();
                                        foreach (var vpId in viewportIds)
                                        {
                                            var viewport = _document.GetElement(vpId) as Viewport;
                                            if (viewport != null)
                                            {
                                                var view = _document.GetElement(viewport.ViewId) as Autodesk.Revit.DB.View;
                                                if (view != null && !view.IsTemplate)
                                                {
                                                    // Create ViewItem wrapper
                                                    var viewItem = new ViewItem
                                                    {
                                                        RevitViewId = view.Id,
                                                        ViewName = view.Name,
                                                        ViewType = view.ViewType.ToString()
                                                    };
                                                    viewsForDxf.Add(viewItem);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            
                            if (viewsForDxf.Any())
                            {
                                WriteDebugLog($"Exporting {viewsForDxf.Count} views to DXF");
                                
                                // ⚠️ REMOVED: Don't set all items to Processing/0% upfront
                                // Callback will set each item to Processing when export starts
                                
                                try
                                {
                                    var dxfManager = new DXFExportManager(_document);
                                    
                                    // Create DXF export settings from UI  
                                    var dxfSettings = new PSDXFExportSettings
                                    {
                                        OutputFolder = outputFolder,
                                        ExportAllViews = false, // Use specific views
                                        Export3DViews = true,
                                        ExportPlanViews = true,
                                        ExportSectionViews = true,
                                        ExportSheetViews = false,
                                        ExcludeTemplateViews = true
                                    };
                                    
                                    WriteDebugLog($"DXF Settings: Export3D={dxfSettings.Export3DViews}, Views count={viewsForDxf.Count}");
                                    
                                    // Export all views at once
                                    var result = dxfManager.ExportViewsToDXF(outputFolder, dxfSettings);
                                    
                                    if (result)
                                    {
                                        WriteDebugLog($"✓ DXF export completed successfully");
                                        
                                        // Update all DXF queue items
                                        Dispatcher.Invoke(() =>
                                        {
                                            foreach (var dxfItem in items.Where(i => i.Format == "DXF"))
                                            {
                                                dxfItem.Progress = 100;
                                                dxfItem.Status = "Completed";
                                            }
                                            ExportQueueDataGrid.Items.Refresh();
                                        });
                                    }
                                    else
                                    {
                                        WriteDebugLog($"✗ DXF export failed");
                                        
                                        // Update failed status
                                        Dispatcher.Invoke(() =>
                                        {
                                            foreach (var dxfItem in items.Where(i => i.Format == "DXF"))
                                            {
                                                dxfItem.Status = "Failed";
                                                dxfItem.Progress = 0;
                                            }
                                            ExportQueueDataGrid.Items.Refresh();
                                        });
                                    }
                                }
                                catch (Exception ex)
                                {
                                    WriteDebugLog($"ERROR in DXF export: {ex.Message}");
                                    MessageBox.Show($"DXF export error: {ex.Message}", "Export Error", 
                                        MessageBoxButton.OK, MessageBoxImage.Error);
                                    
                                    // Update failed status
                                    Dispatcher.Invoke(() =>
                                    {
                                        foreach (var dxfItem in items.Where(i => i.Format == "DXF"))
                                        {
                                            dxfItem.Status = "Failed";
                                        }
                                        ExportQueueDataGrid.Items.Refresh();
                                    });
                                }
                            }
                            else
                            {
                                WriteDebugLog("DXF export skipped - no views available");
                                
                                // Mark as skipped
                                foreach (var dxfItem in items.Where(i => i.Format == "DXF"))
                                {
                                    dxfItem.Status = "Skipped";
                                }
                            }
                        }
                        else
                        {
                            WriteDebugLog($"Format {format} not yet implemented");
                        }
                    }
                    
                    // Final progress update - only if all items are done
                    var allCompleted = items.All(i => i.Status == "Completed");
                    if (allCompleted)
                    {
                        ExportProgressBar.Value = 100;
                        ProgressPercentageText.Text = "Completed 100%";
                        WriteDebugLog("All export items completed successfully");
                        
                        // Generate report if selected
                        var reportType = (ReportComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString();
                        if (reportType != "Don't Save Report")
                        {
                            WriteDebugLog($"Generating {reportType}");
                            // TODO: Implement report generation
                        }
                        
                        // ✓ SET FLAG: Export đã hoàn thành - sẵn sàng reset khi user chọn lại
                        _exportJustCompleted = true;
                        WriteDebugLog("🏁 Export completed - Flag set for auto-reset on next selection");
                        
                        // Show export completed dialog ONLY when all items are done
                        ShowExportCompletedDialog(CreateFolderPathTextBox.Text);
                    }
                    else
                    {
                        // Calculate actual progress based on completed items
                        var completedItems = items.Count(i => i.Status == "Completed");
                        var processingItems = items.Count(i => i.Status.Contains("Processing") || i.Status.Contains("External Event"));
                        var actualProgress = (completedItems * 100.0) / items.Count;
                        ExportProgressBar.Value = actualProgress;
                        ProgressPercentageText.Text = $"Completed {actualProgress:F0}%";
                        WriteDebugLog($"Export partially completed: {completedItems}/{items.Count} items");
                        
                        // Only show message if there are no items still processing (e.g., IFC via ExternalEvent)
                        // and some items actually failed
                        if (processingItems == 0)
                        {
                            var failedItems = items.Count(i => i.Status == "Failed");
                            if (failedItems > 0)
                            {
                                MessageBox.Show($"Export process finished, but some items failed.\n\n" +
                                               $"Completed: {completedItems}/{items.Count} items\n" +
                                               $"Failed: {failedItems} items\n" +
                                               $"Location: {CreateFolderPathTextBox.Text}",
                                               "Export Status", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                        }
                        else
                        {
                            WriteDebugLog($"Export in progress: {processingItems} items still processing via External Events");
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                WriteDebugLog("Export operation was cancelled by user");
                MessageBox.Show("Export was cancelled.", 
                               "Export Cancelled", MessageBoxButton.OK, MessageBoxImage.Information);
                
                // Update status for any pending items
                var items = ExportQueueDataGrid.ItemsSource as ObservableCollection<ExportQueueItem>;
                if (items != null)
                {
                    foreach (var item in items.Where(i => i.Status == "Processing" || i.Status == "Pending"))
                    {
                        item.Status = "Cancelled";
                        item.Progress = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in StartExportButton_Click: {ex.Message}");
                MessageBox.Show($"Error during export: {ex.Message}", 
                               "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // Re-enable button
                StartExportButton.IsEnabled = true;
                StartExportButton.Content = "START EXPORT";
            }
        }

        #endregion

        #region Enhanced UI Event Handlers

        private void FilterByVSSet_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Filter by View/Sheet Set clicked");
            try
            {
                // No longer used - multi-select handles filtering automatically
                WriteDebugLog("Legacy FilterByVSSet_Click - multi-select now active");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in FilterByVSSet_Click: {ex.Message}");
                MessageBox.Show($"Lỗi khi lọc: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ResetFilter_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Reset filter clicked");
            try
            {
                // Uncheck all set selections
                if (ViewSheetSets != null)
                {
                    foreach (var set in ViewSheetSets)
                    {
                        set.IsSelected = false;
                    }
                    OnPropertyChanged(nameof(SelectedSetsDisplay));
                }
                
                // Reload all data
                LoadSheets();
                LoadViews();
                
                MessageBox.Show("Filter reset - showing all items", "Filter Reset", 
                               MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in ResetFilter_Click: {ex.Message}");
                MessageBox.Show($"Lỗi khi reset filter: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SetCustomFileName_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Set Custom File Name clicked");
            try
            {
                if (sender is Button button && button.Tag is SheetItem sheetItem)
                {
                    // Create a simple parameter selection dialog
                    var parameterDialog = new ParameterSelectionDialog(_document, sheetItem);
                    if (parameterDialog.ShowDialog() == true)
                    {
                        string newFileName = parameterDialog.GeneratedFileName;
                        if (!string.IsNullOrEmpty(newFileName))
                        {
                            sheetItem.CustomFileName = newFileName;
                            WriteDebugLog($"Custom file name set to: {newFileName}");
                        }
                    }
                }
                else if (sender is Button buttonView && buttonView.Tag is ViewItem viewItem)
                {
                    // Handle view item parameter selection
                    var parameterDialog = new ParameterSelectionDialog(_document, viewItem);
                    if (parameterDialog.ShowDialog() == true)
                    {
                        string newFileName = parameterDialog.GeneratedFileName;
                        if (!string.IsNullOrEmpty(newFileName))
                        {
                            viewItem.CustomFileName = newFileName;
                            WriteDebugLog($"Custom file name set for view: {newFileName}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in SetCustomFileName_Click: {ex.Message}");
                MessageBox.Show($"Lỗi khi set custom file name: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void EditSelectedFilenames_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Edit Selected Filenames button clicked - Opening CustomFileNameDialog");
            
            try
            {
                bool isSheetMode = SheetsRadio?.IsChecked == true;
                
                if (isSheetMode)
                {
                    // Get ALL sheets (not just selected)
                    var allSheets = Sheets?.ToList();
                    
                    if (allSheets == null || !allSheets.Any())
                    {
                        MessageBox.Show("No sheets available.", "No Sheets", 
                                       MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                    
                    WriteDebugLog($"Applying custom filename to ALL {allSheets.Count} sheets");
                    
                    // Load existing configuration from current profile FOR SHEETS
                    List<Models.SelectedParameterInfo> existingConfig = null;
                    if (_profileManager?.CurrentProfile?.Settings != null)
                    {
                        // Try to load Sheet-specific config first, fallback to old config for backward compatibility
                        var configJson = _profileManager.CurrentProfile.Settings.CustomFileNameConfigJson_Sheets;
                        if (string.IsNullOrEmpty(configJson))
                        {
                            configJson = _profileManager.CurrentProfile.Settings.CustomFileNameConfigJson;
                        }
                        
                        if (!string.IsNullOrEmpty(configJson))
                        {
                            try
                            {
                                existingConfig = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Models.SelectedParameterInfo>>(configJson);
                                WriteDebugLog($"Loaded existing custom file name configuration for SHEETS with {existingConfig?.Count ?? 0} parameters");
                            }
                            catch (Exception jsonEx)
                            {
                                WriteDebugLog($"ERROR deserializing custom file name config: {jsonEx.Message}");
                            }
                        }
                    }
                    
                    // Open CustomFileNameDialog with existing configuration FOR SHEETS
                    var dialog = new CustomFileNameDialog(_document, existingConfig, isViewMode: false);
                    dialog.Owner = this;
                    
                    if (dialog.ShowDialog() == true)
                    {
                        // Save configuration to profile
                        if (_profileManager?.CurrentProfile?.Settings != null)
                        {
                            try
                            {
                                var configJson = Newtonsoft.Json.JsonConvert.SerializeObject(dialog.SelectedParameters);
                                _profileManager.CurrentProfile.Settings.CustomFileNameConfigJson_Sheets = configJson;
                                _profileManager.SaveProfile(_profileManager.CurrentProfile);
                                WriteDebugLog($"Saved custom file name configuration for SHEETS to profile ({dialog.SelectedParameters.Count} parameters)");
                            }
                            catch (Exception saveEx)
                            {
                                WriteDebugLog($"ERROR saving custom file name config to profile: {saveEx.Message}");
                            }
                        }
                        
                        // Apply custom file name configuration to ALL sheets
                        int updatedCount = ApplyCustomFileNameToSheets(allSheets, dialog.SelectedParameters);
                        
                        WriteDebugLog($"Updated {updatedCount} sheets with custom filename configuration");
                        
                        // IMPORTANT: Update Export Queue to reflect new custom names
                        UpdateExportQueue();
                        WriteDebugLog("Export Queue refreshed with updated custom file names");
                        
                        MessageBox.Show($"Successfully applied custom filename to ALL {updatedCount} sheet(s).", 
                                       "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                else
                {
                    // Get ALL views (not just selected)
                    var allViews = Views?.ToList();
                    
                    if (allViews == null || !allViews.Any())
                    {
                        MessageBox.Show("No views available.", "No Views", 
                                       MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                    
                    WriteDebugLog($"Applying custom filename to ALL {allViews.Count} views");
                    
                    // Load existing configuration from current profile FOR VIEWS
                    List<Models.SelectedParameterInfo> existingConfig = null;
                    if (_profileManager?.CurrentProfile?.Settings != null)
                    {
                        // Try to load View-specific config first, fallback to old config for backward compatibility
                        var configJson = _profileManager.CurrentProfile.Settings.CustomFileNameConfigJson_Views;
                        if (string.IsNullOrEmpty(configJson))
                        {
                            configJson = _profileManager.CurrentProfile.Settings.CustomFileNameConfigJson;
                        }
                        
                        if (!string.IsNullOrEmpty(configJson))
                        {
                            try
                            {
                                existingConfig = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Models.SelectedParameterInfo>>(configJson);
                                WriteDebugLog($"Loaded existing custom file name configuration for VIEWS with {existingConfig?.Count ?? 0} parameters");
                            }
                            catch (Exception jsonEx)
                            {
                                WriteDebugLog($"ERROR deserializing custom file name config: {jsonEx.Message}");
                            }
                        }
                    }
                    
                    // Open CustomFileNameDialog with existing configuration FOR VIEWS
                    var dialog = new CustomFileNameDialog(_document, existingConfig, isViewMode: true);
                    dialog.Owner = this;
                    
                    if (dialog.ShowDialog() == true)
                    {
                        // Save configuration to profile
                        if (_profileManager?.CurrentProfile?.Settings != null)
                        {
                            try
                            {
                                var configJson = Newtonsoft.Json.JsonConvert.SerializeObject(dialog.SelectedParameters);
                                _profileManager.CurrentProfile.Settings.CustomFileNameConfigJson_Views = configJson;
                                _profileManager.SaveProfile(_profileManager.CurrentProfile);
                                WriteDebugLog($"Saved custom file name configuration for VIEWS to profile ({dialog.SelectedParameters.Count} parameters)");
                            }
                            catch (Exception saveEx)
                            {
                                WriteDebugLog($"ERROR saving custom file name config to profile: {saveEx.Message}");
                            }
                        }
                        
                        // Apply custom file name configuration to ALL views
                        int updatedCount = ApplyCustomFileNameToViews(allViews, dialog.SelectedParameters);
                        
                        WriteDebugLog($"Updated {updatedCount} views with custom filename configuration");
                        
                        // IMPORTANT: Update Export Queue to reflect new custom names
                        UpdateExportQueue();
                        WriteDebugLog("Export Queue refreshed with updated custom file names");
                        
                        MessageBox.Show($"Successfully applied custom filename to ALL {updatedCount} view(s).", 
                                       "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error editing selected filenames: {ex.Message}");
                MessageBox.Show($"Error editing filenames: {ex.Message}", "Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Apply custom filename configuration to sheets
        /// </summary>
        private int ApplyCustomFileNameToSheets(List<SheetItem> sheets, ObservableCollection<SelectedParameterInfo> parameters)
        {
            int count = 0;
            
            foreach (var sheetItem in sheets)
            {
                try
                {
                    // Get the actual ViewSheet element
                    var sheet = _document.GetElement(sheetItem.Id) as ViewSheet;
                    if (sheet == null) continue;
                    
                    // Generate custom filename from parameters
                    string customFileName = GenerateCustomFileName(sheet, parameters);
                    
                    if (!string.IsNullOrWhiteSpace(customFileName))
                    {
                        sheetItem.CustomFileName = customFileName;
                        count++;
                        WriteDebugLog($"Sheet '{sheet.SheetNumber}' - Custom filename: {customFileName}");
                    }
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"Error applying custom filename to sheet: {ex.Message}");
                }
            }
            
            return count;
        }

        /// <summary>
        /// Apply custom filename configuration to views
        /// </summary>
        private int ApplyCustomFileNameToViews(List<ViewItem> views, ObservableCollection<SelectedParameterInfo> parameters)
        {
            int count = 0;
            
            foreach (var viewItem in views)
            {
                try
                {
                    // Get the actual View element
                    var view = _document.GetElement(viewItem.ViewId) as View;
                    if (view == null) continue;
                    
                    // Generate custom filename from parameters
                    string customFileName = GenerateCustomFileNameFromView(view, parameters);
                    
                    if (!string.IsNullOrWhiteSpace(customFileName))
                    {
                        viewItem.CustomFileName = customFileName;
                        count++;
                        WriteDebugLog($"View '{view.Name}' - Custom filename: {customFileName}");
                    }
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"Error applying custom filename to view: {ex.Message}");
                }
            }
            
            return count;
        }

        /// <summary>
        /// Generate custom filename from ViewSheet parameters
        /// </summary>
        private string GenerateCustomFileName(ViewSheet sheet, ObservableCollection<SelectedParameterInfo> parameters)
        {
            if (parameters == null || parameters.Count == 0)
                return null;
            
            var parts = new List<string>();
            
            foreach (var paramConfig in parameters)
            {
                string value = GetSheetParameterValue(sheet, paramConfig.ParameterName);
                
                if (!string.IsNullOrEmpty(value))
                {
                    string part = $"{paramConfig.Prefix}{value}{paramConfig.Suffix}";
                    parts.Add(part);
                }
            }
            
            string separator = parameters.FirstOrDefault()?.Separator ?? "-";
            return string.Join(separator, parts);
        }

        /// <summary>
        /// Generate custom filename from View parameters
        /// </summary>
        private string GenerateCustomFileNameFromView(View view, ObservableCollection<SelectedParameterInfo> parameters)
        {
            if (parameters == null || parameters.Count == 0)
                return null;
            
            var parts = new List<string>();
            
            foreach (var paramConfig in parameters)
            {
                string value = GetViewParameterValue(view, paramConfig.ParameterName);
                
                if (!string.IsNullOrEmpty(value))
                {
                    string part = $"{paramConfig.Prefix}{value}{paramConfig.Suffix}";
                    parts.Add(part);
                }
            }
            
            string separator = parameters.FirstOrDefault()?.Separator ?? "-";
            return string.Join(separator, parts);
        }

        /// <summary>
        /// Get parameter value from ViewSheet
        /// </summary>
        private string GetSheetParameterValue(ViewSheet sheet, string parameterName)
        {
            try
            {
                // Try built-in parameters first
                switch (parameterName)
                {
                    case "Sheet Number":
                        return sheet.SheetNumber;
                    case "Sheet Name":
                        return sheet.Name;
                    case "Current Revision":
                        return sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION)?.AsString() ?? "";
                    case "Current Revision Date":
                        return sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION_DATE)?.AsString() ?? "";
                    case "Current Revision Description":
                        return sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION_DESCRIPTION)?.AsString() ?? "";
                    case "Approved By":
                        return sheet.get_Parameter(BuiltInParameter.SHEET_APPROVED_BY)?.AsString() ?? "";
                    case "Checked By":
                        return sheet.get_Parameter(BuiltInParameter.SHEET_CHECKED_BY)?.AsString() ?? "";
                    case "Designed By":
                        return sheet.get_Parameter(BuiltInParameter.SHEET_DESIGNED_BY)?.AsString() ?? "";
                    case "Drawn By":
                        return sheet.get_Parameter(BuiltInParameter.SHEET_DRAWN_BY)?.AsString() ?? "";
                    case "Sheet Issue Date":
                        return sheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE)?.AsString() ?? "";
                }
                
                // Try to find parameter by name
                foreach (Parameter param in sheet.Parameters)
                {
                    if (param.Definition.Name.Equals(parameterName, StringComparison.OrdinalIgnoreCase))
                    {
                        return GetParameterValueAsString(param);
                    }
                }
                
                // Try project information parameters
                var projectInfo = _document.ProjectInformation;
                if (projectInfo != null)
                {
                    foreach (Parameter param in projectInfo.Parameters)
                    {
                        if (param.Definition.Name.Equals(parameterName, StringComparison.OrdinalIgnoreCase))
                        {
                            return GetParameterValueAsString(param);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error getting parameter '{parameterName}': {ex.Message}");
            }
            
            return "";
        }

        /// <summary>
        /// Get parameter value from View
        /// </summary>
        private string GetViewParameterValue(View view, string parameterName)
        {
            try
            {
                // Try built-in parameters first
                switch (parameterName)
                {
                    case "View Name":
                        return view.Name;
                    case "View Template":
                        var templateId = view.ViewTemplateId;
                        if (templateId != ElementId.InvalidElementId)
                        {
                            var template = _document.GetElement(templateId);
                            return template?.Name ?? "";
                        }
                        return "";
                }
                
                // Try to find parameter by name
                foreach (Parameter param in view.Parameters)
                {
                    if (param.Definition.Name.Equals(parameterName, StringComparison.OrdinalIgnoreCase))
                    {
                        return GetParameterValueAsString(param);
                    }
                }
                
                // Try project information parameters
                var projectInfo = _document.ProjectInformation;
                if (projectInfo != null)
                {
                    foreach (Parameter param in projectInfo.Parameters)
                    {
                        if (param.Definition.Name.Equals(parameterName, StringComparison.OrdinalIgnoreCase))
                        {
                            return GetParameterValueAsString(param);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error getting view parameter '{parameterName}': {ex.Message}");
            }
            
            return "";
        }

        /// <summary>
        /// Get parameter value as string regardless of storage type
        /// </summary>
        private string GetParameterValueAsString(Parameter param)
        {
            if (param == null || !param.HasValue)
                return "";
            
            switch (param.StorageType)
            {
                case StorageType.String:
                    return param.AsString() ?? "";
                case StorageType.Integer:
                    return param.AsInteger().ToString();
                case StorageType.Double:
                    return param.AsValueString() ?? param.AsDouble().ToString();
                case StorageType.ElementId:
                    var elemId = param.AsElementId();
                    if (elemId != ElementId.InvalidElementId)
                    {
                        var elem = _document.GetElement(elemId);
                        return elem?.Name ?? "";
                    }
                    return "";
                default:
                    return "";
            }
        }
        
        private string PromptForFilename(string title, string defaultValue)
        {
            // Create a simple WPF dialog
            var dialog = new Window
            {
                Title = title,
                Width = 400,
                Height = 150,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize
            };
            
            var grid = new WpfGrid();
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            
            var textBox = new TextBox
            {
                Text = defaultValue,
                Margin = new Thickness(10),
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 14
            };
            WpfGrid.SetRow(textBox, 0);
            grid.Children.Add(textBox);
            
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(10)
            };
            
            var okButton = new Button
            {
                Content = "OK",
                Width = 80,
                Height = 30,
                Margin = new Thickness(0, 0, 10, 0),
                IsDefault = true
            };
            okButton.Click += (s, e) => { dialog.DialogResult = true; dialog.Close(); };
            
            var cancelButton = new Button
            {
                Content = "Cancel",
                Width = 80,
                Height = 30,
                IsCancel = true
            };
            cancelButton.Click += (s, e) => { dialog.DialogResult = false; dialog.Close(); };
            
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            WpfGrid.SetRow(buttonPanel, 1);
            grid.Children.Add(buttonPanel);
            
            dialog.Content = grid;
            
            textBox.Focus();
            textBox.SelectAll();
            
            return dialog.ShowDialog() == true ? textBox.Text : null;
        }

        private void SetAllCustomFileName_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Set All Custom File Name clicked");
            try
            {
                // Determine which items are currently visible
                var targetItems = new List<object>();
                bool isSheetMode = SheetsRadio?.IsChecked == true;
                
                if (isSheetMode)
                {
                    // Get selected sheets first, if none selected then get all sheets
                    var selectedSheets = Sheets?.Where(s => s.IsSelected).ToList() ?? new List<SheetItem>();
                    if (selectedSheets.Any())
                    {
                        targetItems.AddRange(selectedSheets);
                        WriteDebugLog($"Found {selectedSheets.Count} selected sheets");
                    }
                    else
                    {
                        // No sheets selected, apply to all sheets
                        var allSheets = Sheets?.ToList() ?? new List<SheetItem>();
                        targetItems.AddRange(allSheets);
                        WriteDebugLog($"No sheets selected, applying to all {targetItems.Count} sheets");
                    }
                }
                else if (ViewsRadio?.IsChecked == true)
                {
                    // Get selected views first, if none selected then get all views
                    var selectedViews = Views?.Where(v => v.IsSelected).ToList() ?? new List<ViewItem>();
                    if (selectedViews.Any())
                    {
                        targetItems.AddRange(selectedViews);
                        WriteDebugLog($"Found {selectedViews.Count} selected views");
                    }
                    else
                    {
                        // No views selected, apply to all views
                        var allViews = Views?.ToList() ?? new List<ViewItem>();
                        targetItems.AddRange(allViews);
                        WriteDebugLog($"No views selected, applying to all {targetItems.Count} views");
                    }
                }

                if (!targetItems.Any())
                {
                    MessageBox.Show("No sheets or views available to configure.", 
                                   "No Items", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Use the first item to get available parameters
                var firstItem = targetItems.First();
                var parameterDialog = new ParameterSelectionDialog(_document, firstItem);
                
                string actionDescription = isSheetMode ? "sheets" : "views";
                bool hasSelection = false;
                
                if (isSheetMode)
                {
                    hasSelection = Sheets?.Any(s => s.IsSelected) == true;
                }
                else
                {
                    hasSelection = Views?.Any(v => v.IsSelected) == true;
                }
                
                string message = hasSelection 
                    ? $"Configure custom filename for {targetItems.Count} selected {actionDescription}?"
                    : $"No items selected. Configure custom filename for ALL {targetItems.Count} {actionDescription}?";
                    
                var result = MessageBox.Show(message, "Confirm Action", 
                                           MessageBoxButton.YesNo, MessageBoxImage.Question);
                
                if (result != MessageBoxResult.Yes)
                {
                    return;
                }
                
                if (parameterDialog.ShowDialog() == true)
                {
                    string pattern = parameterDialog.GeneratedFileName;
                    if (!string.IsNullOrEmpty(pattern))
                    {
                        int updatedCount = 0;
                        
                        // Apply the pattern to all target items
                        foreach (var item in targetItems)
                        {
                            try
                            {
                                if (item is SheetItem sheet)
                                {
                                    // Generate filename based on sheet's parameters
                                    var sheetDialog = new ParameterSelectionDialog(_document, sheet);
                                    string fileName = sheetDialog.GenerateFilename(pattern, sheet);
                                    sheet.CustomFileName = fileName;
                                    updatedCount++;
                                    WriteDebugLog($"Set custom filename for sheet {sheet.SheetNumber}: {fileName}");
                                }
                                else if (item is ViewItem view)
                                {
                                    // Generate filename based on view's parameters
                                    var viewDialog = new ParameterSelectionDialog(_document, view);
                                    string fileName = viewDialog.GenerateFilename(pattern, view);
                                    view.CustomFileName = fileName;
                                    updatedCount++;
                                    WriteDebugLog($"Set custom filename for view {view.ViewName}: {fileName}");
                                }
                            }
                            catch (Exception itemEx)
                            {
                                WriteDebugLog($"Error setting filename for item: {itemEx.Message}");
                            }
                        }
                        
                        // Force UI update
                        if (isSheetMode && SheetsDataGrid != null)
                        {
                            SheetsDataGrid.Items.Refresh();
                        }
                        else if (ViewsDataGrid != null)
                        {
                            ViewsDataGrid.Items.Refresh();
                        }
                        
                        WriteDebugLog($"Applied custom filename pattern to {updatedCount} items");
                        MessageBox.Show($"Custom filename pattern applied to {updatedCount} {actionDescription} successfully!\n\nPattern: {pattern}", 
                                       "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in SetAllCustomFileName_Click: {ex.Message}");
                MessageBox.Show($"Error setting custom filename: {ex.Message}", 
                               "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FilterSheetsBySet(string setName)
        {
            WriteDebugLog($"Filtering sheets by set: {setName}");
            
            try
            {
                // Get current sheets from the existing Sheets collection
                var allSheets = Sheets ?? new ObservableRangeCollection<SheetItem>();
                var filteredSheets = new ObservableRangeCollection<SheetItem>();
                
                foreach (var sheet in allSheets)
                {
                    bool includeSheet = false;
                    
                    // Filter based on sheet categorization
                    switch (setName.ToUpper())
                    {
                        case "ARCHITECTURAL":
                            includeSheet = IsArchitecturalSheet(sheet);
                            break;
                        case "STRUCTURAL":
                            includeSheet = IsStructuralSheet(sheet);
                            break;
                        case "MEP":
                            includeSheet = IsMEPSheet(sheet);
                            break;
                        case "ALL SHEETS":
                        case "<NONE>":
                        default:
                            includeSheet = true;
                            break;
                    }
                    
                    if (includeSheet)
                    {
                        filteredSheets.Add(sheet);
                    }
                }
                
                Sheets = filteredSheets;
                WriteDebugLog($"Filtered to {filteredSheets.Count} sheets");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error filtering sheets: {ex.Message}");
            }
        }

        private bool IsArchitecturalSheet(SheetItem sheet)
        {
            // Simple logic based on sheet number or name patterns
            string number = sheet.SheetNumber?.ToUpper() ?? "";
            string name = sheet.SheetName?.ToUpper() ?? "";
            
            return number.StartsWith("A") || 
                   name.Contains("ARCHITECTURAL") || 
                   name.Contains("FLOOR PLAN") ||
                   name.Contains("ELEVATION") ||
                   name.Contains("SECTION");
        }

        private bool IsStructuralSheet(SheetItem sheet)
        {
            string number = sheet.SheetNumber?.ToUpper() ?? "";
            string name = sheet.SheetName?.ToUpper() ?? "";
            
            return number.StartsWith("S") || 
                   name.Contains("STRUCTURAL") || 
                   name.Contains("FOUNDATION") ||
                   name.Contains("FRAMING");
        }

        private bool IsMEPSheet(SheetItem sheet)
        {
            string number = sheet.SheetNumber?.ToUpper() ?? "";
            string name = sheet.SheetName?.ToUpper() ?? "";
            
            return number.StartsWith("M") || 
                   number.StartsWith("E") ||
                   number.StartsWith("P") ||
                   name.Contains("MECHANICAL") || 
                   name.Contains("ELECTRICAL") ||
                   name.Contains("PLUMBING") ||
                   name.Contains("MEP");
        }

        #endregion

    }

    #region Parameter Selection Dialog

    public class ParameterSelectionDialog : Window
    {
        private readonly Document _document;
        private readonly object _item;
        private ComboBox _parameterCombo;
        private TextBox _previewTextBox;
        private CheckBox _includeRevisionCheck;
        private CheckBox _includeSheetNumberCheck;
        private CheckBox _includeSheetNameCheck;
        
        public string GeneratedFileName { get; private set; }

        public ParameterSelectionDialog(Document document, object item)
        {
            _document = document;
            _item = item;
            
            Title = "Set Custom File Name from Parameters";
            Width = 500;
            Height = 400;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            
            InitializeDialog();
        }

        private void InitializeDialog()
        {
            var grid = new WpfGrid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            
            grid.Margin = new Thickness(20, 20, 20, 20);

            // Title
            var titleBlock = new TextBlock
            {
                Text = "Configure Custom File Name",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 15)
            };
            WpfGrid.SetRow(titleBlock, 0);
            grid.Children.Add(titleBlock);

            // Include options
            _includeSheetNumberCheck = new CheckBox
            {
                Content = "Include Sheet Number",
                IsChecked = true,
                Margin = new Thickness(0, 5, 0, 5)
            };
            _includeSheetNumberCheck.Checked += UpdatePreview;
            _includeSheetNumberCheck.Unchecked += UpdatePreview;
            WpfGrid.SetRow(_includeSheetNumberCheck, 1);
            grid.Children.Add(_includeSheetNumberCheck);

            _includeSheetNameCheck = new CheckBox
            {
                Content = "Include Sheet Name",
                IsChecked = true,
                Margin = new Thickness(0, 5, 0, 5)
            };
            _includeSheetNameCheck.Checked += UpdatePreview;
            _includeSheetNameCheck.Unchecked += UpdatePreview;
            WpfGrid.SetRow(_includeSheetNameCheck, 2);
            grid.Children.Add(_includeSheetNameCheck);

            _includeRevisionCheck = new CheckBox
            {
                Content = "Include Revision",
                IsChecked = false,
                Margin = new Thickness(0, 5, 0, 5)
            };
            _includeRevisionCheck.Checked += UpdatePreview;
            _includeRevisionCheck.Unchecked += UpdatePreview;
            WpfGrid.SetRow(_includeRevisionCheck, 3);
            grid.Children.Add(_includeRevisionCheck);

            // Parameter selection
            var paramLabel = new TextBlock
            {
                Text = "Additional Parameter:",
                Margin = new Thickness(0, 15, 0, 5)
            };
            WpfGrid.SetRow(paramLabel, 4);
            grid.Children.Add(paramLabel);

            _parameterCombo = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 15)
            };
            _parameterCombo.SelectionChanged += UpdatePreview;
            LoadAvailableParameters();
            WpfGrid.SetRow(_parameterCombo, 5);
            grid.Children.Add(_parameterCombo);

            // Preview
            var previewLabel = new TextBlock
            {
                Text = "Preview:",
                Margin = new Thickness(0, 10, 0, 5)
            };
            WpfGrid.SetRow(previewLabel, 6);
            grid.Children.Add(previewLabel);

            _previewTextBox = new TextBox
            {
                IsReadOnly = true,
                Background = new SolidColorBrush(WpfColor.FromRgb(248, 248, 248)),
                Margin = new Thickness(0, 0, 0, 15)
            };
            WpfGrid.SetRow(_previewTextBox, 6);
            grid.Children.Add(_previewTextBox);

            // Buttons
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var okButton = new Button
            {
                Content = "OK",
                Width = 80,
                Height = 30,
                Margin = new Thickness(0, 0, 10, 0),
                IsDefault = true
            };
            okButton.Click += OkButton_Click;

            var cancelButton = new Button
            {
                Content = "Cancel",
                Width = 80,
                Height = 30,
                IsCancel = true
            };

            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            WpfGrid.SetRow(buttonPanel, 7);
            grid.Children.Add(buttonPanel);

            Content = grid;
            
            // Initial preview update
            UpdatePreview(null, null);
        }

        private void LoadAvailableParameters()
        {
            _parameterCombo.Items.Add("<None>");
            _parameterCombo.Items.Add("Project Number");
            _parameterCombo.Items.Add("Project Name");
            _parameterCombo.Items.Add("Current Date");
            _parameterCombo.Items.Add("Sheet Issue Date");
            _parameterCombo.SelectedIndex = 0;
        }

        private void UpdatePreview(object sender, RoutedEventArgs e)
        {
            try
            {
                var parts = new List<string>();

                if (_includeSheetNumberCheck?.IsChecked == true && _item is SheetItem sheet)
                {
                    parts.Add(sheet.SheetNumber);
                }

                if (_includeSheetNameCheck?.IsChecked == true && _item is SheetItem sheetForName)
                {
                    // Clean sheet name for filename
                    string cleanName = CleanFileName(sheetForName.SheetName);
                    parts.Add(cleanName);
                }

                if (_includeRevisionCheck?.IsChecked == true && _item is SheetItem sheetForRev)
                {
                    parts.Add($"Rev{sheetForRev.Revision ?? "A"}");
                }

                if (_parameterCombo?.SelectedItem?.ToString() != "<None>")
                {
                    string paramValue = GetParameterValue(_parameterCombo.SelectedItem.ToString());
                    if (!string.IsNullOrEmpty(paramValue))
                    {
                        parts.Add(CleanFileName(paramValue));
                    }
                }

                GeneratedFileName = string.Join("_", parts.Where(p => !string.IsNullOrEmpty(p)));
                
                if (_previewTextBox != null)
                {
                    _previewTextBox.Text = GeneratedFileName;
                }
            }
            catch (Exception ex)
            {
                if (_previewTextBox != null)
                {
                    _previewTextBox.Text = $"Error: {ex.Message}";
                }
            }
        }

        private string GetParameterValue(string parameterName)
        {
            switch (parameterName)
            {
                case "Project Number":
                    return _document.ProjectInformation.Number ?? "";
                case "Project Name":
                    return _document.ProjectInformation.Name ?? "";
                case "Current Date":
                    return DateTime.Now.ToString("yyyyMMdd");
                case "Sheet Issue Date":
                    return DateTime.Now.ToString("yyyyMMdd");
                default:
                    return "";
            }
        }

        private string CleanFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return "";
            
            // Remove invalid characters and clean up
            var invalidChars = Path.GetInvalidFileNameChars();
            string cleaned = new string(fileName.Where(c => !invalidChars.Contains(c)).ToArray());
            
            // Replace spaces with underscores and limit length
            cleaned = cleaned.Replace(" ", "_").Replace("-", "_");
            
            // Remove consecutive underscores
            while (cleaned.Contains("__"))
            {
                cleaned = cleaned.Replace("__", "_");
            }
            
            return cleaned.Trim('_').Substring(0, Math.Min(cleaned.Length, 50));
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(GeneratedFileName))
            {
                MessageBox.Show("Please configure at least one parameter for the file name.", 
                               "Invalid Configuration", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            DialogResult = true;
            Close();
        }

        // Method to generate filename for any item based on current dialog settings
        public string GenerateFilename(string pattern, object item)
        {
            try
            {
                var parts = new List<string>();

                if (_includeSheetNumberCheck?.IsChecked == true && item is SheetItem sheet)
                {
                    parts.Add(sheet.SheetNumber);
                }
                else if (item is ViewItem view)
                {
                    parts.Add(view.ViewType?.Replace(" ", "_"));
                }

                if (_includeSheetNameCheck?.IsChecked == true && item is SheetItem sheetForName)
                {
                    string cleanName = CleanFileName(sheetForName.SheetName);
                    parts.Add(cleanName);
                }
                else if (item is ViewItem viewForName)
                {
                    string cleanName = CleanFileName(viewForName.ViewName);
                    parts.Add(cleanName);
                }

                if (_includeRevisionCheck?.IsChecked == true && item is SheetItem sheetForRev)
                {
                    parts.Add($"Rev{sheetForRev.Revision ?? "A"}");
                }

                if (_parameterCombo?.SelectedItem?.ToString() != "<None>")
                {
                    string paramValue = GetParameterValue(_parameterCombo.SelectedItem.ToString());
                    if (!string.IsNullOrEmpty(paramValue))
                    {
                        parts.Add(CleanFileName(paramValue));
                    }
                }

                return string.Join("_", parts.Where(p => !string.IsNullOrEmpty(p)));
            }
            catch
            {
                return pattern; // Fallback to original pattern
            }
        }

        #endregion

        // ===== IFC SETUP PROFILE MANAGEMENT TEMPORARILY DISABLED =====
        // Reason: WPF temporary assembly build issue
        // The complete profile management system (289 lines) is commented out below
        // due to WPF _wpftmp.csproj compilation errors where temporary assembly
        // cannot access ExportPlusXMLProfile properties and XMLProfileManager methods
        // 
        // SOLUTION OPTIONS:
        // 1. Move to separate ProfileManagementHelper class
        // 2. Implement using Commands/Behaviors pattern
        // 3. Use conditional compilation (#if !XAML_COMPILATION)
        
        #region IFC Setup Profile Management (DISABLED - WPF Build Issue)

        /*

        /// <summary>
        /// Initialize IFC Setup Profiles collection and configuration paths
        /// </summary>
        private void InitializeIFCSetups()
        {
            // Initialize IFC Setups Collection
            IFCCurrentSetups = new ObservableCollection<string>
            {
                "<In-Session Setup>",
                "IFC 2x3 Coordination View 2.0",
                "IFC 2x3 Coordination View",
                "IFC 2x3 GSA Concept Design BIM 2010",
                "IFC 2x3 Basic FM Handover View",
                "IFC 2x2 Coordination View",
                "IFC 2x2 Singapore BCA e-Plan Check",
                "IFC 2x3 COBie 2.4 Design Deliverable View",
                "IFC4 Reference View",
                "IFC4 Design Transfer View",
                "Typical Setup"
            };
            
            // Initialize configuration paths mapping
            _ifcSetupConfigPaths = new Dictionary<string, string>();
            
            // Get IFC profiles directory (in %AppData%\ExportPlus\IFCProfiles)
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string ifcProfilesDir = Path.Combine(appDataPath, "ExportPlus", "IFCProfiles");
            
            // Create directory if not exists
            try
            {
                if (!Directory.Exists(ifcProfilesDir))
                {
                    Directory.CreateDirectory(ifcProfilesDir);
                    WriteDebugLog($"Created IFC profiles directory: {ifcProfilesDir}");
                }
                
                // Map setup names to file paths
                _ifcSetupConfigPaths["IFC 2x3 Coordination View 2.0"] = Path.Combine(ifcProfilesDir, "IFC_2x3_CV2.0.xml");
                _ifcSetupConfigPaths["IFC 2x3 Coordination View"] = Path.Combine(ifcProfilesDir, "IFC_2x3_CV.xml");
                _ifcSetupConfigPaths["IFC 2x3 GSA Concept Design BIM 2010"] = Path.Combine(ifcProfilesDir, "IFC_2x3_GSA.xml");
                _ifcSetupConfigPaths["IFC 2x3 Basic FM Handover View"] = Path.Combine(ifcProfilesDir, "IFC_2x3_FM.xml");
                _ifcSetupConfigPaths["IFC 2x2 Coordination View"] = Path.Combine(ifcProfilesDir, "IFC_2x2_CV.xml");
                _ifcSetupConfigPaths["IFC 2x2 Singapore BCA e-Plan Check"] = Path.Combine(ifcProfilesDir, "IFC_2x2_SG_BCA.xml");
                _ifcSetupConfigPaths["IFC 2x3 COBie 2.4 Design Deliverable View"] = Path.Combine(ifcProfilesDir, "IFC_2x3_COBie.xml");
                _ifcSetupConfigPaths["IFC4 Reference View"] = Path.Combine(ifcProfilesDir, "IFC4_Reference.xml");
                _ifcSetupConfigPaths["IFC4 Design Transfer View"] = Path.Combine(ifcProfilesDir, "IFC4_Design.xml");
                _ifcSetupConfigPaths["Typical Setup"] = Path.Combine(ifcProfilesDir, "Typical_Setup.xml");
                
                WriteDebugLog($"IFC Setup profiles mapped to: {ifcProfilesDir}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR initializing IFC profiles directory: {ex.Message}");
            }
            
            // Set default selected setup
            SelectedIFCSetup = "<In-Session Setup>";
        }

        /// <summary>
        /// Handle IFC Setup selection changed
        /// </summary>
        private void OnIFCSetupChanged()
        {
            try
            {
                WriteDebugLog($"IFC Setup changed to: {SelectedIFCSetup}");
                
                // If In-Session, keep current settings
                if (SelectedIFCSetup == "<In-Session Setup>")
                {
                    WriteDebugLog("Using In-Session setup - keeping current settings");
                    return;
                }
                
                // Try to load setup from file
                if (_ifcSetupConfigPaths != null && 
                    _ifcSetupConfigPaths.TryGetValue(SelectedIFCSetup, out string filePath))
                {
                    if (File.Exists(filePath))
                    {
                        WriteDebugLog($"Loading IFC setup from: {filePath}");
                        ApplyIFCSettingsFromFile(filePath);
                    }
                    else
                    {
                        WriteDebugLog($"IFC setup file not found: {filePath} - creating default");
                        CreateDefaultIFCSetup(SelectedIFCSetup, filePath);
                    }
                }
                else
                {
                    WriteDebugLog($"No file path mapping for setup: {SelectedIFCSetup}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR in OnIFCSetupChanged: {ex.Message}");
                MessageBox.Show($"Error loading IFC setup: {ex.Message}", 
                                "Setup Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Apply IFC settings from XML file
        /// </summary>
        private void ApplyIFCSettingsFromFile(string filePath)
        {
            try
            {
                var profileManager = new XMLProfileManager();
                var profile = profileManager.ImportProfile(filePath);
                
                if (profile != null && profile.IFCSettings != null)
                {
                    // Apply settings to current IFCSettings
                    IFCSettings = profile.IFCSettings;
                    
                    WriteDebugLog($"IFC settings applied from: {filePath}");
                    
                    // Show success message
                    MessageBox.Show($"IFC setup '{SelectedIFCSetup}' loaded successfully!", 
                                    "Setup Loaded", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    WriteDebugLog($"Failed to load IFC settings from: {filePath}");
                    MessageBox.Show("Failed to load IFC setup configuration.", 
                                    "Load Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR applying IFC settings from file: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Create default IFC setup configuration file
        /// </summary>
        private void CreateDefaultIFCSetup(string setupName, string filePath)
        {
            try
            {
                // Create default settings based on setup name
                var defaultSettings = new IFCExportSettings();
                
                // Configure settings based on setup type
                switch (setupName)
                {
                    case "IFC 2x3 Coordination View 2.0":
                        defaultSettings.IFCVersion = "IFC 2x3 Coordination View 2.0";
                        defaultSettings.ExportBaseQuantities = false;
                        break;
                        
                    case "IFC 2x3 Coordination View":
                        defaultSettings.IFCVersion = "IFC 2x3 Coordination View";
                        defaultSettings.ExportBaseQuantities = false;
                        break;
                        
                    case "IFC 2x3 GSA Concept Design BIM 2010":
                        defaultSettings.IFCVersion = "IFC 2x3 GSA Concept Design BIM 2010";
                        defaultSettings.ExportBaseQuantities = true;
                        break;
                        
                    case "IFC 2x3 Basic FM Handover View":
                        defaultSettings.IFCVersion = "IFC 2x3 Basic FM Handover View";
                        defaultSettings.ExportBaseQuantities = true;
                        defaultSettings.SpaceBoundaries = "1st Level";
                        break;
                        
                    case "IFC 2x2 Coordination View":
                        defaultSettings.IFCVersion = "IFC 2x2 Coordination View";
                        defaultSettings.ExportBaseQuantities = false;
                        break;
                        
                    case "IFC 2x2 Singapore BCA e-Plan Check":
                        defaultSettings.IFCVersion = "IFC 2x2 Singapore BCA e-Plan Check";
                        defaultSettings.ExportBaseQuantities = true;
                        break;
                        
                    case "IFC 2x3 COBie 2.4 Design Deliverable View":
                        defaultSettings.IFCVersion = "IFC 2x3 COBie 2.4 Design Deliverable View";
                        defaultSettings.ExportBaseQuantities = true;
                        defaultSettings.SpaceBoundaries = "2nd Level";
                        break;
                        
                    case "IFC4 Reference View":
                        defaultSettings.IFCVersion = "IFC4 Reference View";
                        defaultSettings.ExportBaseQuantities = false;
                        break;
                        
                    case "IFC4 Design Transfer View":
                        defaultSettings.IFCVersion = "IFC4 Design Transfer View";
                        defaultSettings.ExportBaseQuantities = true;
                        break;
                        
                    case "Typical Setup":
                        defaultSettings.IFCVersion = "IFC 2x3 Coordination View 2.0";
                        defaultSettings.ExportBaseQuantities = false;
                        defaultSettings.DetailLevel = "Medium";
                        break;
                }
                
                // Create profile
                var profile = new ExportPlusXMLProfile
                {
                    ProfileName = setupName,
                    CreatedDate = DateTime.Now,
                    IFCSettings = defaultSettings
                };
                
                // Save to file
                var profileManager = new XMLProfileManager();
                bool success = profileManager.ExportProfile(profile, filePath);
                
                if (success)
                {
                    WriteDebugLog($"Created default IFC setup: {filePath}");
                    
                    // Apply the settings
                    IFCSettings = defaultSettings;
                }
                else
                {
                    WriteDebugLog($"Failed to create default IFC setup: {filePath}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR creating default IFC setup: {ex.Message}");
            }
        }

        /// <summary>
        /// Save current IFC settings to selected setup
        /// </summary>
        public void SaveCurrentIFCSetup()
        {
            try
            {
                if (SelectedIFCSetup == "<In-Session Setup>")
                {
                    MessageBox.Show("Cannot save to In-Session setup. Please select or create a named setup.", 
                                    "Save Setup", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                if (_ifcSetupConfigPaths.TryGetValue(SelectedIFCSetup, out string filePath))
                {
                    // Create profile with current settings
                    var profile = new ExportPlusXMLProfile
                    {
                        ProfileName = SelectedIFCSetup,
                        CreatedDate = DateTime.Now,
                        IFCSettings = IFCSettings
                    };
                    
                    // Save to file
                    var profileManager = new XMLProfileManager();
                    bool success = profileManager.ExportProfile(profile, filePath);
                    
                    if (success)
                    {
                        WriteDebugLog($"Saved IFC setup to: {filePath}");
                        MessageBox.Show($"IFC setup '{SelectedIFCSetup}' saved successfully!", 
                                        "Setup Saved", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Failed to save IFC setup.", 
                                        "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR saving IFC setup: {ex.Message}");
                MessageBox.Show($"Error saving IFC setup: {ex.Message}", 
                                "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        */

        #endregion

        // ===== IFC SETTINGS IMPORT/EXPORT TEMPORARILY DISABLED =====
        // Reason: WPF temporary assembly build issue
        // Code available in version control - will be re-enabled after WPF build fix
        // Browse functionality works via BrowseFileBehavior attached property

        #region IFC Settings Event Handlers (DISABLED - WPF Build Issue)

        // NOTE: These methods are commented out due to WPF temporary assembly validation issues
        // The .g.cs file containing x:Name field declarations is not generated until AFTER
        // the temporary assembly compilation succeeds, creating a circular dependency.
        // 
        // SOLUTION: Implement Browse button functionality using:
        // 1. Behaviors/Attached Properties (no code-behind references)
        // 2. MVVM pattern with Commands
        // 3. Post-deployment event wiring (outside WPF build process)

        /*
        /// <summary>
        /// Wire up Browse button Click handlers in constructor
        /// This avoids WPF XAML compilation issues with x:Name controls
        /// </summary>
        private void WireUpIFCBrowseButtons()
        {
            try
            {
                if (BrowseUserPsetsButtonIFC != null)
                {
                    BrowseUserPsetsButtonIFC.Click += BrowseIFCFile_Click;
                }
                
                if (BrowseParamMappingButtonIFC != null)
                {
                    BrowseParamMappingButtonIFC.Click += BrowseIFCFile_Click;
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error wiring up IFC Browse buttons: {ex.Message}");
            }
        }

        /// <summary>
        /// Universal Browse button click handler for IFC file selection
        /// Uses Button.Tag to determine which TextBox to update
        /// </summary>
        private void BrowseIFCFile_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button)) return;

            try
            {
                var dialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Select IFC Configuration File",
                    Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                    FilterIndex = 1,
                    CheckFileExists = false
                };

                // Determine which TextBox to update based on Button.Tag
                TextBox targetTextBox = null;
                string fileType = button.Tag?.ToString() ?? "";

                if (fileType == "UserPsets")
                {
                    dialog.Title = "Select User-Defined Property Sets File";
                    targetTextBox = UserPsetsPathTextBoxIFC;
                }
                else if (fileType == "ParamMapping")
                {
                    dialog.Title = "Select Parameter Mapping Table File";
                    targetTextBox = ParamMappingPathTextBoxIFC;
                }

                if (targetTextBox == null)
                {
                    WriteDebugLog($"ERROR: Could not find target TextBox for button Tag='{fileType}'");
                    return;
                }

                // Set initial directory if path exists
                string currentPath = targetTextBox.Text;
                if (!string.IsNullOrEmpty(currentPath))
                {
                    var directory = System.IO.Path.GetDirectoryName(currentPath);
                    if (!string.IsNullOrEmpty(directory) && System.IO.Directory.Exists(directory))
                    {
                        dialog.InitialDirectory = directory;
                    }
                }

                if (dialog.ShowDialog() == true)
                {
                    targetTextBox.Text = dialog.FileName;
                    WriteDebugLog($"IFC file selected ({fileType}): {dialog.FileName}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in BrowseIFCFile_Click: {ex.Message}");
                System.Windows.MessageBox.Show(
                    $"Error selecting file: {ex.Message}",
                    "Browse Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
        */

        #endregion IFC Settings Event Handlers

        /* TEMPORARILY DISABLED - DWG Export Setup Event Handlers
         * TODO: Re-enable after fixing WPF partial class compiler issues
         * 
        #region DWG Export Setup Event Handlers

        /// <summary>
        /// Load available DWG export setups from Revit document
        /// </summary>
        private void LoadDWGExportSetups()
        {
            try
            {
                WriteDebugLog("Loading DWG Export Setups from Revit...");
                
                if (DWGExportSetupComboBox == null)
                {
                    WriteDebugLog("DWGExportSetupComboBox is null, cannot load setups");
                    return;
                }
                
                DWGExportSetupComboBox.Items.Clear();
                
                // Get predefined setup names from Revit
                IList<string> setupNames = BaseExportOptions.GetPredefinedSetupNames(_document);
                
                WriteDebugLog($"Found {setupNames.Count} DWG export setups");
                
                foreach (string setupName in setupNames)
                {
                    DWGExportSetupComboBox.Items.Add(setupName);
                    WriteDebugLog($"  - {setupName}");
                }
                
                // Add default option if no setups found
                if (DWGExportSetupComboBox.Items.Count == 0)
                {
                    DWGExportSetupComboBox.Items.Add("Default Setup");
                    WriteDebugLog("No setups found, added 'Default Setup'");
                }
                
                // Select first item by default
                DWGExportSetupComboBox.SelectedIndex = 0;
                WriteDebugLog($"Selected default setup: {DWGExportSetupComboBox.SelectedItem}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error loading DWG export setups: {ex.Message}");
                
                // Fallback to default
                if (DWGExportSetupComboBox != null)
                {
                    DWGExportSetupComboBox.Items.Clear();
                    DWGExportSetupComboBox.Items.Add("Default Setup");
                    DWGExportSetupComboBox.SelectedIndex = 0;
                }
            }
        }

        /// <summary>
        /// Handle DWG Export Setup selection changed
        /// </summary>
        private void DWGExportSetupComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (DWGExportSetupComboBox.SelectedItem != null)
                {
                    string selectedSetup = DWGExportSetupComboBox.SelectedItem.ToString();
                    WriteDebugLog($"DWG Export Setup changed to: {selectedSetup}");
                    
                    // Store selected setup in export settings
                    this.ExportSettings.DWGExportSetupName = selectedSetup;
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in DWGExportSetupComboBox_SelectionChanged: {ex.Message}");
            }
        }

        /// <summary>
        /// Open Revit's DWG Export Settings dialog to modify export setups
        /// </summary>
        private void ModifyDWGExportSetup_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WriteDebugLog("Opening Revit's DWG Export Settings dialog...");
                
                // Inform user to use Revit's menu
                MessageBox.Show(
                    "To create or modify DWG Export Setups:\n\n" +
                    "1. Close this dialog\n" +
                    "2. In Revit, go to: File > Export > CAD Formats > DWG\n" +
                    "3. In the export dialog, click 'Modify Setup' button\n" +
                    "4. Create or edit your setup and save it\n" +
                    "5. Reopen this ExportPlus dialog\n" +
                    "6. Your new setup will appear in the dropdown",
                    "Modify DWG Export Setup",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in ModifyDWGExportSetup_Click: {ex.Message}");
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion DWG Export Setup Event Handlers
        */
    }
}
