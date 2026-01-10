using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Quoc_MEP.Export.Models;

namespace Quoc_MEP.Export.Views
{
    /// <summary>
    /// View/Sheet Set Management - Partial class for ExportPlusMainWindow
    /// </summary>
    public partial class ExportPlusMainWindow
    {
        #region View/Sheet Set Management
        
        /// <summary>
        /// Load all View/Sheet Sets into ItemsControl
        /// </summary>
        private void LoadViewSheetSets()
        {
            try
            {
                WriteDebugLog("Loading View/Sheet Sets...");
                
                if (_viewSheetSetManager == null)
                {
                    WriteDebugLog("ViewSheetSetManager not initialized");
                    return;
                }
                
                var sets = _viewSheetSetManager.GetAllViewSheetSets();
                WriteDebugLog($"Found {sets.Count} View/Sheet Sets");
                
                // Create ObservableCollection and subscribe to PropertyChanged
                ViewSheetSets = new ObservableCollection<ViewSheetSetInfo>();
                
                foreach (var set in sets)
                {
                    // Subscribe to IsSelected changes
                    set.PropertyChanged += ViewSheetSet_PropertyChanged;
                    ViewSheetSets.Add(set);
                    WriteDebugLog($"  - {set.Name} ({set.TotalCount} items)");
                }
                
                // Bind to ItemsControl
                ViewSheetSetItems.ItemsSource = ViewSheetSets;
                
                // Select "All Sheets" by default if no filter checkbox
                if (sets.Count > 0 && FilterByVSCheckBox?.IsChecked != true)
                {
                    var allSheets = ViewSheetSets.FirstOrDefault(s => s.Name == "All Sheets");
                    if (allSheets != null)
                    {
                        allSheets.IsSelected = true;
                    }
                }
                
                WriteDebugLog("View/Sheet Sets loaded successfully");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR loading View/Sheet Sets: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Handle ViewSheetSet IsSelected property changes
        /// </summary>
        private void ViewSheetSet_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ViewSheetSetInfo.IsSelected))
            {
                WriteDebugLog($"ViewSheetSet selection changed");
                OnPropertyChanged(nameof(SelectedSetsDisplay));
                
                // Auto-apply filter if checkbox is checked
                if (FilterByVSCheckBox?.IsChecked == true)
                {
                    ApplyMultiSetFilter();
                }
            }
        }
        
        /// <summary>
        /// Apply multi-select filter combining all selected sets
        /// </summary>
        private void ApplyMultiSetFilter()
        {
            if (ViewSheetSets == null)
            {
                WriteDebugLog("ApplyMultiSetFilter: ViewSheetSets is null");
                return;
            }
            
            var selectedSets = ViewSheetSets.Where(s => s.IsSelected).ToList();
            WriteDebugLog($"Applying multi-set filter with {selectedSets.Count} selected sets");
            
            try
            {
                bool isSheetMode = SheetsRadio?.IsChecked == true;
                
                if (isSheetMode)
                {
                    // Combine sheets from all selected sets
                    var combinedSheetIds = new HashSet<ElementId>();
                    
                    if (selectedSets.Count == 0 || selectedSets.Any(s => s.Name.StartsWith("All Sheets")))
                    {
                        // Show all sheets if nothing selected or "All Sheets" selected
                        WriteDebugLog("Showing all sheets (no filter or All Sheets selected)");
                        var allSheets = _viewSheetSetManager.GetSheetsFromSet("All Sheets");
                        foreach (var sheet in allSheets)
                            combinedSheetIds.Add(sheet.Id);
                        WriteDebugLog($"Got {allSheets.Count} sheets from 'All Sheets'");
                    }
                    else
                    {
                        // Combine from selected sets
                        WriteDebugLog($"Combining sheets from {selectedSets.Count} selected sets");
                        foreach (var set in selectedSets)
                        {
                            WriteDebugLog($"Getting sheets from set: '{set.Name}'");
                            var sheets = _viewSheetSetManager.GetSheetsFromSet(set.Name);
                            WriteDebugLog($"  -> Found {sheets.Count} sheets in '{set.Name}'");
                            foreach (var sheet in sheets)
                                combinedSheetIds.Add(sheet.Id);
                        }
                    }
                    
                    WriteDebugLog($"Combined {combinedSheetIds.Count} unique sheets from selected sets");
                    
                    // Get all sheets and filter
                    var allProjectSheets = new FilteredElementCollector(_document)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .Where(s => !s.IsTemplate)
                        .ToList();
                    
                    Sheets.Clear();
                    
                    foreach (var sheet in allProjectSheets.Where(s => combinedSheetIds.Contains(s.Id)))
                    {
                        var sheetItem = CreateSheetItem(sheet);
                        Sheets.Add(sheetItem);
                    }
                    
                    WriteDebugLog($"Displaying {Sheets.Count} sheets after multi-set filter");
                }
                else
                {
                    // Combine views from all selected sets
                    var combinedViewIds = new HashSet<ElementId>();
                    
                    if (selectedSets.Count == 0 || selectedSets.Any(s => s.Name.StartsWith("All Views")))
                    {
                        // Show all views if nothing selected or "All Views" selected
                        WriteDebugLog("Showing all views (no filter or All Views selected)");
                        var allViews = _viewSheetSetManager.GetViewsFromSet("All Views");
                        foreach (var view in allViews)
                            combinedViewIds.Add(view.Id);
                        WriteDebugLog($"Got {allViews.Count} views from 'All Views'");
                    }
                    else
                    {
                        // Combine from selected sets
                        WriteDebugLog($"Combining views from {selectedSets.Count} selected sets");
                        foreach (var set in selectedSets)
                        {
                            WriteDebugLog($"Getting views from set: '{set.Name}'");
                            var views = _viewSheetSetManager.GetViewsFromSet(set.Name);
                            WriteDebugLog($"  -> Found {views.Count} views in '{set.Name}'");
                            foreach (var view in views)
                                combinedViewIds.Add(view.Id);
                        }
                    }
                    
                    WriteDebugLog($"Combined {combinedViewIds.Count} unique views from selected sets");
                    
                    // Get all views and filter
                    var allProjectViews = new FilteredElementCollector(_document)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => !v.IsTemplate && v.CanBePrinted)
                        .ToList();
                    
                    Views.Clear();
                    
                    foreach (var view in allProjectViews.Where(v => combinedViewIds.Contains(v.Id)))
                    {
                        var viewItem = CreateViewItem(view);
                        Views.Add(viewItem);
                    }
                    
                    WriteDebugLog($"Displaying {Views.Count} views after multi-set filter");
                }
                
                UpdateStatusText();
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR applying multi-set filter: {ex.Message}");
                MessageBox.Show(
                    $"Failed to apply filter:\n\n{ex.Message}",
                    "Filter Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// Apply View/Sheet Set filter to current list
        /// </summary>
        private void ApplyViewSheetSetFilter(ViewSheetSetInfo setInfo)
        {
            if (setInfo == null)
            {
                WriteDebugLog("ApplyViewSheetSetFilter: setInfo is null");
                return;
            }
            
            WriteDebugLog($"Applying View/Sheet Set filter: {setInfo.Name}");
            
            try
            {
                bool isSheetMode = SheetsRadio?.IsChecked == true;
                
                if (isSheetMode)
                {
                    // Filter sheets
                    var filteredSheets = _viewSheetSetManager.GetSheetsFromSet(setInfo.Name);
                    var filteredIds = new HashSet<ElementId>(filteredSheets.Select(s => s.Id));
                    
                    WriteDebugLog($"Filtered to {filteredSheets.Count} sheets from set");
                    
                    // Get all sheets and filter
                    var allSheets = new FilteredElementCollector(_document)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .Where(s => !s.IsTemplate)
                        .ToList();
                    
                    Sheets.Clear();
                    
                    foreach (var sheet in allSheets.Where(s => filteredIds.Contains(s.Id)))
                    {
                        var sheetItem = CreateSheetItem(sheet);
                        Sheets.Add(sheetItem);
                    }
                    
                    WriteDebugLog($"Displaying {Sheets.Count} sheets from set '{setInfo.Name}'");
                }
                else
                {
                    // Filter views
                    var filteredViews = _viewSheetSetManager.GetViewsFromSet(setInfo.Name);
                    var filteredIds = new HashSet<ElementId>(filteredViews.Select(v => v.Id));
                    
                    WriteDebugLog($"Filtered to {filteredViews.Count} views from set");
                    
                    // Get all views and filter
                    var allViews = new FilteredElementCollector(_document)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => !v.IsTemplate && v.CanBePrinted)
                        .ToList();
                    
                    Views.Clear();
                    
                    foreach (var view in allViews.Where(v => filteredIds.Contains(v.Id)))
                    {
                        var viewItem = CreateViewItem(view);
                        Views.Add(viewItem);
                    }
                    
                    WriteDebugLog($"Displaying {Views.Count} views from set '{setInfo.Name}'");
                }
                
                UpdateStatusText();
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR applying View/Sheet Set filter: {ex.Message}");
                MessageBox.Show(
                    $"Failed to apply filter:\n\n{ex.Message}",
                    "Filter Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// Helper method to create SheetItem from ViewSheet
        /// </summary>
        private SheetItem CreateSheetItem(ViewSheet sheet)
        {
            return new SheetItem
            {
                Id = sheet.Id,
                SheetNumber = sheet.SheetNumber ?? "",
                SheetName = sheet.Name ?? "",
                IsSelected = false,
                Revision = GetParameterValue(sheet, "Current Revision"),
                Size = sheet.get_Parameter(BuiltInParameter.SHEET_HEIGHT)?.AsValueString() ?? ""
            };
        }
        
        /// <summary>
        /// Helper method to create ViewItem from View
        /// </summary>
        private ViewItem CreateViewItem(View view)
        {
            return new ViewItem
            {
                RevitViewId = view.Id,
                ViewId = view.Id.IntegerValue.ToString(),
                ViewName = view.Name ?? "",
                ViewType = view.ViewType.ToString(),
                Scale = view.Scale.ToString(),
                IsSelected = false
            };
        }
        
        /// <summary>
        /// Helper method to get parameter value
        /// </summary>
        private string GetParameterValue(Element element, string paramName)
        {
            try
            {
                var param = element.LookupParameter(paramName);
                if (param != null && param.HasValue)
                {
                    return param.AsString() ?? "";
                }
            }
            catch { }
            return "";
        }
        
        #endregion View/Sheet Set Management
    }
}
