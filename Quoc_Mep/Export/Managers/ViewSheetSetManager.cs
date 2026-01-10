using Autodesk.Revit.DB;
using Quoc_MEP.Export.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Quoc_MEP.Export.Managers
{
    /// <summary>
    /// Manager for Revit View/Sheet Sets
    /// </summary>
    public class ViewSheetSetManager
    {
        private readonly Document _doc;
        
        public ViewSheetSetManager(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }
        
        /// <summary>
        /// Get all View/Sheet Sets from project
        /// Reads saved ViewSheetSets from the document
        /// </summary>
        public List<ViewSheetSetInfo> GetAllViewSheetSets()
        {
            var sets = new List<ViewSheetSetInfo>();
            
            try
            {
                // Add "All Sheets" built-in option
                var allSheetsSet = new ViewSheetSetInfo("All Sheets")
                {
                    IsBuiltIn = true
                };
                
                var allSheets = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Where(s => !s.IsTemplate)
                    .Select(s => s.Id)
                    .ToList();
                    
                allSheetsSet.SheetIds.AddRange(allSheets);
                sets.Add(allSheetsSet);
                
                // Add "All Views" built-in option
                var allViewsSet = new ViewSheetSetInfo("All Views")
                {
                    IsBuiltIn = true
                };
                
                var allViews = new FilteredElementCollector(_doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate && v.CanBePrinted)
                    .Select(v => v.Id)
                    .ToList();
                    
                allViewsSet.ViewIds.AddRange(allViews);
                sets.Add(allViewsSet);
                
                // Get saved ViewSheetSets from document (created via Print dialog)
                var collector = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewSheetSet));
                    
                System.Diagnostics.Debug.WriteLine($"[ViewSheetSetManager] Found {collector.GetElementCount()} ViewSheetSets in document");
                
                foreach (ViewSheetSet vss in collector)
                {
                    if (vss == null || string.IsNullOrEmpty(vss.Name))
                        continue;
                        
                    System.Diagnostics.Debug.WriteLine($"[ViewSheetSetManager] Processing set: {vss.Name}");
                    
                    var setInfo = new ViewSheetSetInfo(vss.Name)
                    {
                        IsBuiltIn = false
                    };
                    
                    // Get views and sheets in this set using Views property
                    if (vss.Views != null && !vss.Views.IsEmpty)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ViewSheetSetManager]   ViewSet has {vss.Views.Size} items");
                        
                        foreach (View view in vss.Views)
                        {
                            if (view == null)
                                continue;
                            
                            if (view is ViewSheet sheet)
                            {
                                setInfo.SheetIds.Add(sheet.Id);
                                System.Diagnostics.Debug.WriteLine($"[ViewSheetSetManager]     - Sheet: {sheet.SheetNumber} - {sheet.Name}");
                            }
                            else if (view.CanBePrinted && !view.IsTemplate)
                            {
                                setInfo.ViewIds.Add(view.Id);
                                System.Diagnostics.Debug.WriteLine($"[ViewSheetSetManager]     - View: {view.Name}");
                            }
                        }
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"[ViewSheetSetManager]   Total: {setInfo.SheetIds.Count} sheets, {setInfo.ViewIds.Count} views");
                    
                    // Add set even if empty (to show in dropdown)
                    sets.Add(setInfo);
                }
                
                System.Diagnostics.Debug.WriteLine($"[ViewSheetSetManager] Total sets loaded: {sets.Count}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ViewSheetSetManager] ERROR: {ex.Message}\n{ex.StackTrace}");
            }
            
            return sets;
        }
        
        /// <summary>
        /// Create new ViewSheetSet from selected sheets/views
        /// </summary>
        public ViewSheetSet CreateViewSheetSet(string name, List<ElementId> selectedIds)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Set name cannot be empty", nameof(name));
                
            if (selectedIds == null || selectedIds.Count == 0)
                throw new ArgumentException("Must select at least one sheet or view", nameof(selectedIds));
            
            using (var trans = new Transaction(_doc, "Create ViewSheetSet"))
            {
                trans.Start();
                
                try
                {
                    // Check if name already exists
                    var existing = new FilteredElementCollector(_doc)
                        .OfClass(typeof(ViewSheetSet))
                        .Cast<ViewSheetSet>()
                        .FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                        
                    if (existing != null)
                    {
                        throw new InvalidOperationException($"ViewSheetSet '{name}' already exists");
                    }
                    
                    // Create ViewSet and add selected items
                    var printManager = _doc.PrintManager;
                    printManager.PrintRange = PrintRange.Select;
                    
                    var viewSet = new ViewSet();
                    
                    foreach (var id in selectedIds)
                    {
                        var elem = _doc.GetElement(id);
                        if (elem is View view && view.CanBePrinted)
                        {
                            viewSet.Insert(view);
                        }
                    }
                    
                    if (viewSet.IsEmpty)
                    {
                        throw new InvalidOperationException("No printable views selected");
                    }
                    
                    // Create the ViewSheetSet
                    var viewSheetSetting = printManager.ViewSheetSetting;
                    viewSheetSetting.CurrentViewSheetSet.Views = viewSet;
                    viewSheetSetting.SaveAs(name);
                    
                    trans.Commit();
                    
                    // Return the newly created set
                    return new FilteredElementCollector(_doc)
                        .OfClass(typeof(ViewSheetSet))
                        .Cast<ViewSheetSet>()
                        .FirstOrDefault(s => s.Name == name);
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    throw new InvalidOperationException($"Failed to create ViewSheetSet: {ex.Message}", ex);
                }
            }
        }
        
        /// <summary>
        /// Get sheets from ViewSheetSet by name
        /// </summary>
        public List<ViewSheet> GetSheetsFromSet(string setName)
        {
            if (string.IsNullOrWhiteSpace(setName))
                return new List<ViewSheet>();
            
            try
            {
                // Handle built-in "All Sheets"
                if (setName.Equals("All Sheets", StringComparison.OrdinalIgnoreCase) || 
                    setName.StartsWith("All Sheets ("))
                {
                    return new FilteredElementCollector(_doc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .Where(s => !s.IsTemplate)
                        .ToList();
                }
                
                // Find the ViewSheetSet
                var vss = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewSheetSet))
                    .Cast<ViewSheetSet>()
                    .FirstOrDefault(s => s.Name.Equals(setName, StringComparison.OrdinalIgnoreCase) ||
                                        setName.StartsWith(s.Name + " ("));
                    
                if (vss == null)
                    return new List<ViewSheet>();
                
                var sheets = new List<ViewSheet>();
                foreach (ElementId id in vss.Views)
                {
                    if (_doc.GetElement(id) is ViewSheet sheet && !sheet.IsTemplate)
                        sheets.Add(sheet);
                }
                
                return sheets;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting sheets from set: {ex.Message}");
                return new List<ViewSheet>();
            }
        }
        
        /// <summary>
        /// Get views from ViewSheetSet by name
        /// </summary>
        public List<View> GetViewsFromSet(string setName)
        {
            if (string.IsNullOrWhiteSpace(setName))
                return new List<View>();
            
            try
            {
                // Handle built-in "All Views"
                if (setName.Equals("All Views", StringComparison.OrdinalIgnoreCase) || 
                    setName.StartsWith("All Views ("))
                {
                    return new FilteredElementCollector(_doc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => !v.IsTemplate && v.CanBePrinted)
                        .ToList();
                }
                
                // Find the ViewSheetSet
                var vss = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewSheetSet))
                    .Cast<ViewSheetSet>()
                    .FirstOrDefault(s => s.Name.Equals(setName, StringComparison.OrdinalIgnoreCase) ||
                                        setName.StartsWith(s.Name + " ("));
                    
                if (vss == null)
                    return new List<View>();
                
                var views = new List<View>();
                foreach (ElementId id in vss.Views)
                {
                    if (_doc.GetElement(id) is View view && 
                        !view.IsTemplate && 
                        view.CanBePrinted &&
                        !(view is ViewSheet))
                    {
                        views.Add(view);
                    }
                }
                
                return views;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting views from set: {ex.Message}");
                return new List<View>();
            }
        }
        
        /// <summary>
        /// Get ViewSheetSet by name
        /// </summary>
        public ViewSheetSet GetViewSheetSetByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;
                
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheetSet))
                .Cast<ViewSheetSet>()
                .FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }
        
        /// <summary>
        /// Add sheets/views to an existing ViewSheetSet
        /// </summary>
        public bool AddToExistingSet(string setName, List<ElementId> elementIds)
        {
            if (string.IsNullOrWhiteSpace(setName) || elementIds == null || !elementIds.Any())
                return false;
                
            try
            {
                var vss = GetViewSheetSetByName(setName);
                if (vss == null)
                {
                    System.Diagnostics.Debug.WriteLine($"ViewSheetSet '{setName}' not found");
                    return false;
                }
                
                using (var trans = new Transaction(_doc, "Add to Existing ViewSheetSet"))
                {
                    trans.Start();
                    
                    // Get current ViewSet
                    var currentViewSet = vss.Views;
                    var newViewSet = new ViewSet();
                    
                    // Copy existing views to new ViewSet
                    foreach (View existingView in currentViewSet)
                    {
                        newViewSet.Insert(existingView);
                    }
                    
                    // Add new views/sheets
                    int addedCount = 0;
                    foreach (var id in elementIds)
                    {
                        var elem = _doc.GetElement(id);
                        if (elem is View view && view.CanBePrinted)
                        {
                            // Check if not already in set
                            bool alreadyExists = false;
                            foreach (View v in newViewSet)
                            {
                                if (v.Id == view.Id)
                                {
                                    alreadyExists = true;
                                    break;
                                }
                            }
                            
                            if (!alreadyExists)
                            {
                                newViewSet.Insert(view);
                                addedCount++;
                            }
                        }
                    }
                    
                    // Update the ViewSheetSet
                    vss.Views = newViewSet;
                    
                    trans.Commit();
                    
                    System.Diagnostics.Debug.WriteLine($"Added {addedCount} items to set '{setName}'");
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adding to existing set: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Delete ViewSheetSet by name
        /// </summary>
        public bool DeleteViewSheetSet(string setName)
        {
            if (string.IsNullOrWhiteSpace(setName))
                return false;
            
            try
            {
                var vss = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewSheetSet))
                    .Cast<ViewSheetSet>()
                    .FirstOrDefault(s => s.Name.Equals(setName, StringComparison.OrdinalIgnoreCase));
                    
                if (vss == null)
                    return false;
                
                using (var trans = new Transaction(_doc, "Delete ViewSheetSet"))
                {
                    trans.Start();
                    _doc.Delete(vss.Id);
                    trans.Commit();
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error deleting ViewSheetSet: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Check if ViewSheetSet name already exists
        /// </summary>
        public bool SetNameExists(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;
            
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheetSet))
                .Cast<ViewSheetSet>()
                .Any(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }
    }
}
