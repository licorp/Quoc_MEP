using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Windows.Media;
using FastWpfGrid;

namespace Quoc_MEP.Export.Models
{
    /// <summary>
    /// FastWpfGrid model for rendering SheetItem collection with instant performance
    /// </summary>
    public class SheetListGridModel : FastGridModelBase
    {
        private ObservableCollection<SheetItem> _sheets;
        private IFastGridView _attachedView;
        
        // Column indices
        private const int COL_CHECKBOX = 0;
        private const int COL_SHEET_NUMBER = 1;
        private const int COL_SHEET_NAME = 2;
        private const int COL_REVISION = 3;
        private const int COL_SIZE = 4;
        private const int COL_CUSTOM_FILENAME = 5;
        
        public SheetListGridModel(ObservableCollection<SheetItem> sheets)
        {
            _sheets = sheets ?? new ObservableCollection<SheetItem>();
            
            // Listen to collection changes for live updates
            _sheets.CollectionChanged += Sheets_CollectionChanged;
        }

        private void Sheets_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            // Notify grid of data changes
            if (_attachedView != null)
            {
                _attachedView.InvalidateAll();
            }
        }

        public override int ColumnCount => 6;

        public override int RowCount => _sheets?.Count ?? 0;

        public override IFastGridCell GetCell(IFastGridView view, int row, int column)
        {
            if (row < 0 || row >= RowCount) return null;
            
            var sheet = _sheets[row];
            var cell = new FastGridCellImpl();
            
            // Alternate row background
            if (row % 2 == 1)
            {
                cell.BackgroundColor = Color.FromRgb(0xFA, 0xFA, 0xFA);
            }
            
            switch (column)
            {
                case COL_CHECKBOX:
                    // Render checkbox-like indicator
                    var checkText = sheet.IsSelected ? "☑" : "☐";
                    var checkBlock = cell.AddTextBlock(checkText);
                    checkBlock.IsBold = false;
                    break;
                    
                case COL_SHEET_NUMBER:
                    if (!string.IsNullOrEmpty(sheet.SheetNumber))
                    {
                        cell.AddTextBlock(sheet.SheetNumber);
                    }
                    break;
                    
                case COL_SHEET_NAME:
                    if (!string.IsNullOrEmpty(sheet.SheetName))
                    {
                        cell.AddTextBlock(sheet.SheetName);
                    }
                    break;
                    
                case COL_REVISION:
                    if (!string.IsNullOrEmpty(sheet.Revision))
                    {
                        cell.AddTextBlock(sheet.Revision);
                    }
                    break;
                    
                case COL_SIZE:
                    if (!string.IsNullOrEmpty(sheet.PaperSize))
                    {
                        cell.AddTextBlock(sheet.PaperSize);
                    }
                    break;
                    
                case COL_CUSTOM_FILENAME:
                    if (!string.IsNullOrEmpty(sheet.CustomFileName))
                    {
                        cell.AddTextBlock(sheet.CustomFileName);
                    }
                    break;
            }
            
            return cell;
        }

        public override IFastGridCell GetRowHeader(IFastGridView view, int row)
        {
            var cell = new FastGridCellImpl();
            cell.AddTextBlock((row + 1).ToString());
            return cell;
        }

        public override IFastGridCell GetColumnHeader(IFastGridView view, int column)
        {
            var cell = new FastGridCellImpl();
            cell.BackgroundColor = Color.FromRgb(0xE8, 0xE8, 0xE8);
            
            string headerText = "";
            switch (column)
            {
                case COL_CHECKBOX:
                    headerText = "All";
                    break;
                case COL_SHEET_NUMBER:
                    headerText = "Sheet Number";
                    break;
                case COL_SHEET_NAME:
                    headerText = "Sheet Name";
                    break;
                case COL_REVISION:
                    headerText = "Revision";
                    break;
                case COL_SIZE:
                    headerText = "Size";
                    break;
                case COL_CUSTOM_FILENAME:
                    headerText = "Custom File Name";
                    break;
            }
            
            var textBlock = cell.AddTextBlock(headerText);
            textBlock.IsBold = true;
            
            return cell;
        }

        public override IFastGridCell GetGridHeader(IFastGridView view)
        {
            var cell = new FastGridCellImpl();
            cell.BackgroundColor = Color.FromRgb(0xE8, 0xE8, 0xE8);
            return cell;
        }

        public override void AttachView(IFastGridView view)
        {
            base.AttachView(view);
            _attachedView = view;
        }

        public override void DetachView(IFastGridView view)
        {
            base.DetachView(view);
            if (_attachedView == view)
            {
                _attachedView = null;
            }
        }
        
        public override void HandleCommand(IFastGridView view, FastGridCellAddress address, object commandParameter, ref bool handled)
        {
            // Handle cell clicks (e.g., checkbox toggle)
            if (address.Column == COL_CHECKBOX && address.Row.HasValue)
            {
                int row = address.Row.Value;
                if (row >= 0 && row < RowCount)
                {
                    _sheets[row].IsSelected = !_sheets[row].IsSelected;
                    view.InvalidateModelCell(row, COL_CHECKBOX);
                    handled = true;
                }
            }
            
            base.HandleCommand(view, address, commandParameter, ref handled);
        }
    }
}
