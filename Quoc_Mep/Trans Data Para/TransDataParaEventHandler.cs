using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ricaun.Revit.UI.StatusBar;
using Nice3point.Revit.Extensions;

namespace Quoc_MEP
{
    /// <summary>
    /// Event Handler để xử lý copy parameters trong ExternalEvent
    /// </summary>
    public class TransDataParaEventHandler : IExternalEventHandler
    {
        private CopyParametersRequestEventArgs _request;
        private TransDataParaWindow _window;
        private CancellationTokenSource _cancellationTokenSource;

        public void SetRequest(CopyParametersRequestEventArgs request, TransDataParaWindow window)
        {
            _request = request;
            _window = window;
            _cancellationTokenSource = new CancellationTokenSource();
        }

        public void CancelOperation()
        {
            _cancellationTokenSource?.Cancel();
            Logger.Warning("User cancelled the operation");
        }

        public void Execute(UIApplication uiApp)
        {
            try
            {
                if (_request == null || _window == null)
                {
                    Logger.Warning("TransDataParaEventHandler: No request or window");
                    return;
                }

                Document doc = uiApp.ActiveUIDocument.Document;

                Logger.StartOperation("Copy Parameter Data");
                Logger.Info($"═══════════════════════════════════════");
                Logger.Info($"COPY CONFIGURATION:");
                Logger.Info($"  Source Group: {_request.SourceGroupName}");
                Logger.Info($"  Source Parameter: {_request.SourceParameterName}");
                Logger.Info($"  ↓");
                Logger.Info($"  Target Group: {_request.TargetGroupName}");
                Logger.Info($"  Target Parameter: {_request.TargetParameterName}");
                Logger.Info($"═══════════════════════════════════════");

                // TỐI ƯU: Get elements by categories nhanh hơn với LINQ
                IEnumerable<Element> elements;

                if (_request.SelectedCategories.Count > 0)
                {
                    Logger.Info($"Selected Categories: {string.Join(", ", _request.SelectedCategories)}");
                    
                    // TỐI ƯU: Dùng Single collector với Where thay vì multiple collectors
                    elements = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .Where(elem => elem.Category != null && 
                                      _request.SelectedCategories.Contains(elem.Category.Name));
                    
                    Logger.Info($"Found {elements.Count()} elements in selected categories");
                }
                else
                {
                    elements = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType();
                    Logger.Info($"Processing ALL elements: {elements.Count()}");
                }

                int totalElements = elements.Count();
                int successCount = 0;
                int skippedCount = 0;
                int errorCount = 0;

                Logger.Info($"Overwrite Mode: {(_request.OverwriteExisting ? "YES - Overwrite all" : "NO - Only empty parameters")}");

                using (Transaction trans = new Transaction(doc, "Copy Parameter Data"))
                {
                    trans.Start();

                    // Tạo progress bar với nút Cancel
                    ProgressBarView progressBar = null;
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        progressBar = new ProgressBarView("Copying Parameters", totalElements);
                        progressBar.CancelRequested += (s, e) => _cancellationTokenSource.Cancel();
                        progressBar.Show();
                    });

                    // Xử lý từng element với hỗ trợ cancel
                    foreach (var elem in elements)
                    {
                        // Check for cancellation
                        if (_cancellationTokenSource.Token.IsCancellationRequested)
                            break;

                        try
                        {
                            // TỐI ƯU: Dùng Nice3point Extensions để lấy parameter nhanh hơn
                            Parameter sourceParam = elem.GetParameter(_request.SourceParameterName);
                            Parameter targetParam = elem.GetParameter(_request.TargetParameterName);

                            // Kiểm tra null và readonly
                            if (sourceParam == null || targetParam == null || targetParam.IsReadOnly)
                            {
                                skippedCount++;
                                continue;
                            }

                            // CHECK: Nếu Source parameter RỖNG → BỎ QUA
                            if (!sourceParam.HasValue)
                            {
                                skippedCount++;
                                continue;
                            }
                            
                            // Kiểm tra có nên copy không
                            bool shouldCopy = _request.OverwriteExisting || !targetParam.HasValue;
                            if (!shouldCopy)
                            {
                                skippedCount++;
                                continue;
                            }

                            // TỐI ƯU: Copy theo StorageType nhanh hơn
                            if (sourceParam.StorageType != targetParam.StorageType)
                            {
                                skippedCount++;
                                continue;
                            }

                            // Copy value theo từng type
                            bool copySuccess = false;
                            switch (sourceParam.StorageType)
                            {
                                case StorageType.Double:
                                    targetParam.Set(sourceParam.AsDouble());
                                    copySuccess = true;
                                    break;
                                    
                                case StorageType.Integer:
                                    targetParam.Set(sourceParam.AsInteger());
                                    copySuccess = true;
                                    break;
                                    
                                case StorageType.String:
                                    string strValue = sourceParam.AsString();
                                    if (!string.IsNullOrEmpty(strValue))
                                    {
                                        targetParam.Set(strValue);
                                        copySuccess = true;
                                    }
                                    break;
                                    
                                case StorageType.ElementId:
                                    targetParam.Set(sourceParam.AsElementId());
                                    copySuccess = true;
                                    break;
                            }

                            if (copySuccess)
                            {
                                successCount++;
                            }
                            else
                            {
                                skippedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            Logger.Error($"  ID {elem.Id.IntegerValue}: Error - {ex.Message}");
                        }

                        // Update progress bar
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (progressBar != null && !progressBar.IsClosed)
                            {
                                progressBar.Update(1);
                            }
                        });
                    }

                    // Đóng progress bar
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        progressBar?.Close();
                    });

                    // Check if operation was cancelled before committing
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        trans.RollBack();
                        Logger.Warning("Transaction rolled back due to cancellation");
                        
                        // Show cancellation message on UI thread
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            System.Windows.MessageBox.Show(
                                "Operation cancelled by user.\nNo changes were made.",
                                "Operation Cancelled",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Information);
                        });
                        
                        return;
                    }
                    
                    trans.Commit();
                }

                Logger.Info($"=== RESULTS ===");
                Logger.Info($"Total Elements: {totalElements}");
                Logger.Info($"Successfully Copied: {successCount}");
                Logger.Info($"Skipped: {skippedCount}");
                Logger.Info($"Errors: {errorCount}");
                Logger.EndOperation("Copy Parameter Data");

                // Show result on UI thread
                _window.Dispatcher.Invoke(() =>
                {
                    System.Windows.MessageBox.Show(
                        $"Copy thành công!\n\n" +
                        $"Tổng elements: {totalElements}\n" +
                        $"Đã copy: {successCount} elements\n" +
                        $"Bỏ qua (đã có giá trị): {skippedCount} elements\n" +
                        $"Lỗi: {errorCount} elements\n\n" +
                        $"Xem DebugView để theo dõi chi tiết!",
                        "Kết quả",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);

                    // Show form again for reuse
                    _window.ShowForReuse();
                });

            }
            catch (Exception ex)
            {
                Logger.Error("TransDataParaEventHandler.Execute failed", ex);
                
                _window?.Dispatcher.Invoke(() =>
                {
                    // Close progress bar on error
                    try
                    {
                        var progressBars = System.Windows.Application.Current?.Windows.OfType<ProgressBarView>();
                        if (progressBars != null)
                        {
                            foreach (var pb in progressBars.ToList())
                            {
                                pb.Close();
                            }
                        }
                    }
                    catch { }
                    
                    System.Windows.MessageBox.Show(
                        $"Lỗi: {ex.Message}",
                        "Lỗi",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Error);
                    
                    _window.ShowForReuse();
                });
            }
            finally
            {
                // Cleanup cancellation token source
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
        }

        public string GetName()
        {
            return "TransDataParaEventHandler";
        }
    }

    /// <summary>
    /// Event args cho copy parameters request
    /// </summary>
    public class CopyParametersRequestEventArgs : EventArgs
    {
        public string SourceGroupName { get; set; }
        public string SourceParameterName { get; set; }
        public string TargetGroupName { get; set; }
        public string TargetParameterName { get; set; }
        public List<string> SelectedCategories { get; set; }
        public bool OverwriteExisting { get; set; }
    }
}
