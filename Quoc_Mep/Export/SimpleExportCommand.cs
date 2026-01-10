using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Quoc_MEP.Export.Views;
using System.Diagnostics;

namespace Quoc_MEP.Export
{
    [Transaction(TransactionMode.Manual)]
    public class SimpleExportCommand : IExternalCommand
    {
        private static ExportPlusMainWindow _cachedWindow = null;
        private static Document _cachedDocument = null;
        private static bool _isExecuting = false;  // ⚡ GUARD: Prevent duplicate execution
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // ⚡⚡⚡ CRITICAL: If already executing, skip duplicate call
            if (_isExecuting)
            {
                Debug.WriteLine("[Export +] ⚠️ DUPLICATE COMMAND - Already executing, ignoring...");
                return Result.Succeeded;
            }
            
            _isExecuting = true;
            
            try
            {
                Debug.WriteLine("[Export +] ===== SimpleExportCommand started =====");
                
                var doc = commandData?.Application?.ActiveUIDocument?.Document;
                if (doc == null)
                {
                    Debug.WriteLine("[Export +] ❌ ERROR: No active document!");
                    TaskDialog.Show("Export+", "Không có document nào được mở!");
                    return Result.Failed;
                }
                
                Debug.WriteLine($"[Export +] Document: {doc.Title}");
                
                // Check if cached window exists and matches document
                if (_cachedWindow != null && _cachedDocument?.Equals(doc) == true)
                {
                    Debug.WriteLine("[Export +] ✅ CACHE HIT - Reusing existing window");
                    
                    try
                    {
                        if (!_cachedWindow.IsVisible)
                        {
                            _cachedWindow.Show();
                            Debug.WriteLine("[Export +] 📖 Cached window shown");
                        }
                        else
                        {
                            _cachedWindow.Activate();
                            Debug.WriteLine("[Export +] 🔝 Cached window activated");
                        }
                        
                        // ⚡ Reset flag immediately after showing cached window
                        _isExecuting = false;
                        Debug.WriteLine("[Export +] 🔓 Flag reset after cache hit");
                        return Result.Succeeded;
                    }
                    catch (Exception cacheEx)
                    {
                        Debug.WriteLine($"[Export +] ⚠️ Cached window error: {cacheEx.Message}");
                        Debug.WriteLine("[Export +] 🔄 Will create new window...");
                        _cachedWindow = null;
                        _cachedDocument = null;
                        // Continue to create new window (flag still set)
                    }
                }
                
                Debug.WriteLine("[Export +] 🆕 Creating NEW ExportPlusMainWindow...");
                
                ExportPlusMainWindow mainWindow = null;
                try
                {
                    mainWindow = new ExportPlusMainWindow(doc, commandData.Application);
                    Debug.WriteLine("[Export +] ✅ Window constructor completed successfully");
                }
                catch (Exception constructorEx)
                {
                    Debug.WriteLine($"[Export +] ❌ CONSTRUCTOR ERROR: {constructorEx.Message}");
                    Debug.WriteLine($"[Export +] Stack trace: {constructorEx.StackTrace}");
                    
                    if (constructorEx.InnerException != null)
                    {
                        Debug.WriteLine($"[Export +] Inner exception: {constructorEx.InnerException.Message}");
                    }
                    
                    TaskDialog.Show("Export+ Error", 
                        $"Lỗi khởi tạo window:\n\n{constructorEx.Message}\n\n" +
                        $"Vui lòng kiểm tra DebugView để xem chi tiết.");
                    
                    _isExecuting = false; // Reset flag on constructor failure
                    return Result.Failed;
                }
                
                if (mainWindow == null)
                {
                    Debug.WriteLine("[Export +] ❌ ERROR: mainWindow is null after constructor!");
                    _isExecuting = false; // Reset flag on null window
                    return Result.Failed;
                }
                
                // Cache the window
                _cachedWindow = mainWindow;
                _cachedDocument = doc;
                Debug.WriteLine("[Export +] 💾 Window cached successfully");
                
                // ⚡⚡⚡ CRITICAL: Reset flag when window ACTUALLY closes (not just hide)
                mainWindow.Closed += (s, e) =>
                {
                    try
                    {
                        Debug.WriteLine("[Export +] 🔓 Window closed - resetting _isExecuting flag");
                        _isExecuting = false;
                        _cachedWindow = null;
                        _cachedDocument = null;
                    }
                    catch (Exception closedEx)
                    {
                        Debug.WriteLine($"[Export +] ⚠️ Error in Closed handler: {closedEx.Message}");
                        _isExecuting = false; // Always reset even if error
                    }
                };
                
                // Setup closing handler (prevent close, just hide)
                mainWindow.Closing += (s, e) =>
                {
                    try
                    {
                        e.Cancel = true;
                        mainWindow.Hide();
                        Debug.WriteLine("[Export +] 🙈 Window hidden (not closed)");
                        
                        // ⚡ IMPORTANT: Reset flag when user hides window
                        // User có thể mở lại ngay lập tức
                        Debug.WriteLine("[Export +] 🔓 Resetting _isExecuting flag after hide");
                        _isExecuting = false;
                    }
                    catch (Exception hideEx)
                    {
                        Debug.WriteLine($"[Export +] ⚠️ Error hiding window: {hideEx.Message}");
                        _isExecuting = false; // Always reset even if error
                    }
                };
                
                // Show the window
                try
                {
                    mainWindow.Show();
                    Debug.WriteLine("[Export +] ✅ Export window shown successfully");
                    Debug.WriteLine("[Export +] 🔒 Flag will be reset by Closing event handler");
                }
                catch (Exception showEx)
                {
                    Debug.WriteLine($"[Export +] ❌ ERROR showing window: {showEx.Message}");
                    TaskDialog.Show("Export+ Error", $"Lỗi hiển thị window:\n\n{showEx.Message}");
                    _isExecuting = false; // Reset immediately on show failure
                    return Result.Failed;
                }
                
                Debug.WriteLine("[Export +] ===== Command completed successfully =====");
                // ⚡ DON'T reset flag here - let Closing event handle it
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Export +] ❌❌❌ UNHANDLED ERROR ❌❌❌");
                Debug.WriteLine($"[Export +] Error: {ex.Message}");
                Debug.WriteLine($"[Export +] Type: {ex.GetType().Name}");
                Debug.WriteLine($"[Export +] Stack trace: {ex.StackTrace}");
                
                if (ex.InnerException != null)
                {
                    Debug.WriteLine($"[Export +] Inner exception: {ex.InnerException.Message}");
                    Debug.WriteLine($"[Export +] Inner stack: {ex.InnerException.StackTrace}");
                }
                
                TaskDialog.Show("Export+ Fatal Error", 
                    $"Lỗi nghiêm trọng:\n\n{ex.Message}\n\n" +
                    $"Type: {ex.GetType().Name}\n\n" +
                    $"Vui lòng báo lỗi với thông tin từ DebugView.");
                
                // ⚡ Reset flag on error since window won't handle it
                _isExecuting = false;
                return Result.Failed;
            }
            // ⚡ NO finally block - flag reset handled by:
            //   - Closing event (normal case)
            //   - catch block (error case)
            //   - show failure (explicit reset)
        }
    }
}
