using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace ScheduleManager
{
    public class ScheduleManagerViewModel : BaseViewModel
    {
        private readonly UIApplication _uiApp;
        private readonly Document _doc;
        
        private ObservableCollection<ScheduleRow> _scheduleRows;
        private string _statusMessage;
        private string _documentName;
        private string _loadingMessage;
        private bool _isLoading;
        private int _totalSchedules;
        private int _totalRows;

        public ObservableCollection<ScheduleRow> ScheduleRows
        {
            get => _scheduleRows;
            set { _scheduleRows = value; OnPropertyChanged(); }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        public string DocumentName
        {
            get => _documentName;
            set { _documentName = value; OnPropertyChanged(); }
        }

        public string LoadingMessage
        {
            get => _loadingMessage;
            set { _loadingMessage = value; OnPropertyChanged(); }
        }

        public bool IsLoading
        {
            get => _isLoading;
            set { _isLoading = value; OnPropertyChanged(); }
        }

        public int TotalSchedules
        {
            get => _totalSchedules;
            set { _totalSchedules = value; OnPropertyChanged(); }
        }

        public int TotalRows
        {
            get => _totalRows;
            set { _totalRows = value; OnPropertyChanged(); }
        }

        public ICommand LoadSchedulesCommand { get; }
        public ICommand ExportToExcelCommand { get; }
        public ICommand ImportFromExcelCommand { get; }
        public ICommand RefreshCommand { get; }

        public ScheduleManagerViewModel(UIApplication uiApp)
        {
            _uiApp = uiApp;
            _doc = uiApp.ActiveUIDocument.Document;
            
            ScheduleRows = new ObservableCollection<ScheduleRow>();
            DocumentName = _doc.Title;
            StatusMessage = "Ready. Click Load Schedules to begin.";

            LoadSchedulesCommand = new RelayCommand(async () => await LoadSchedulesAsync());
            ExportToExcelCommand = new RelayCommand(ExportToExcel, () => ScheduleRows.Any());
            ImportFromExcelCommand = new RelayCommand(ImportFromExcel);
            RefreshCommand = new RelayCommand(async () => await LoadSchedulesAsync());
        }

        private async Task LoadSchedulesAsync()
        {
            try
            {
                IsLoading = true;
                LoadingMessage = "Loading schedules...";
                StatusMessage = "Loading schedule data from Revit...";
                ScheduleRows.Clear();

                await Task.Run(() =>
                {
                    var schedules = new FilteredElementCollector(_doc)
                        .OfClass(typeof(ViewSchedule))
                        .Cast<ViewSchedule>()
                        .Where(s => !s.IsTemplate)
                        .OrderBy(s => s.Name)
                        .ToList();

                    TotalSchedules = schedules.Count;

                    foreach (var schedule in schedules)
                    {
                        try
                        {
                            TableData tableData = schedule.GetTableData();
                            TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);

                            int rowCount = sectionData.NumberOfRows;
                            int colCount = sectionData.NumberOfColumns;

                            for (int row = 1; row < rowCount; row++)
                            {
                                for (int col = 0; col < colCount; col++)
                                {
                                    string value = schedule.GetCellText(SectionType.Body, row, col);
                                    string fieldName = GetColumnHeader(schedule, col);

                                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                                    {
                                        ScheduleRows.Add(new ScheduleRow
                                        {
                                            ScheduleName = schedule.Name,
                                            ElementId = schedule.Id.IntegerValue.ToString(),
                                            FieldName = fieldName,
                                            Value = value,
                                            RowIndex = row,
                                            ColumnIndex = col
                                        });
                                    });
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error loading schedule {schedule.Name}: {ex.Message}");
                        }
                    }

                    TotalRows = ScheduleRows.Count;
                });

                StatusMessage = $"Loaded {TotalSchedules} schedules with {TotalRows} data cells";
                LoadingMessage = "";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error: {ex.Message}";
                MessageBox.Show($"Error loading schedules:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private string GetColumnHeader(ViewSchedule schedule, int columnIndex)
        {
            try
            {
                ScheduleDefinition definition = schedule.Definition;
                if (columnIndex < definition.GetFieldCount())
                {
                    ScheduleFieldId fieldId = definition.GetFieldId(columnIndex);
                    ScheduleField field = definition.GetField(fieldId);
                    return field.GetName();
                }
            }
            catch { }
            return $"Column {columnIndex}";
        }

        private void ExportToExcel()
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = $"{_doc.Title}_Schedules_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                    Title = "Export Schedules to Excel"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    IsLoading = true;
                    LoadingMessage = "Exporting to Excel...";
                    StatusMessage = "Exporting schedule data to Excel...";

                    using (var workbook = new XLWorkbook())
                    {
                        var scheduleGroups = ScheduleRows.GroupBy(r => r.ScheduleName);

                        foreach (var group in scheduleGroups)
                        {
                            string sheetName = SanitizeSheetName(group.Key);
                            var worksheet = workbook.Worksheets.Add(sheetName);

                            worksheet.Cell(1, 1).Value = "Element ID";
                            worksheet.Cell(1, 2).Value = "Field Name";
                            worksheet.Cell(1, 3).Value = "Value";
                            worksheet.Cell(1, 4).Value = "Row Index";
                            worksheet.Cell(1, 5).Value = "Column Index";

                            var headerRow = worksheet.Range(1, 1, 1, 5);
                            headerRow.Style.Font.Bold = true;
                            headerRow.Style.Fill.BackgroundColor = XLColor.LightBlue;
                            headerRow.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            int row = 2;
                            foreach (var item in group)
                            {
                                worksheet.Cell(row, 1).Value = item.ElementId;
                                worksheet.Cell(row, 2).Value = item.FieldName;
                                worksheet.Cell(row, 3).Value = item.Value;
                                worksheet.Cell(row, 4).Value = item.RowIndex;
                                worksheet.Cell(row, 5).Value = item.ColumnIndex;
                                row++;
                            }

                            worksheet.Columns().AdjustToContents();
                        }

                        workbook.SaveAs(saveFileDialog.FileName);
                    }

                    StatusMessage = $"Exported to: {Path.GetFileName(saveFileDialog.FileName)}";
                    MessageBox.Show($"Export successful!\n\nFile: {saveFileDialog.FileName}\nSchedules: {scheduleGroups.Count()}\nRows: {ScheduleRows.Count}",
                        "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Export failed: {ex.Message}";
                MessageBox.Show($"Error exporting to Excel:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
                LoadingMessage = "";
            }
        }

        private void ImportFromExcel()
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    Title = "Import Schedules from Excel"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    IsLoading = true;
                    LoadingMessage = "Importing from Excel...";
                    StatusMessage = "Reading Excel file...";

                    using (var workbook = new XLWorkbook(openFileDialog.FileName))
                    {
                        int updatedCount = 0;

                        foreach (var worksheet in workbook.Worksheets)
                        {
                            string scheduleName = worksheet.Name;
                            
                            for (int row = 2; row <= worksheet.LastRowUsed().RowNumber(); row++)
                            {
                                try
                                {
                                    string elementId = worksheet.Cell(row, 1).GetString();
                                    string fieldName = worksheet.Cell(row, 2).GetString();
                                    string newValue = worksheet.Cell(row, 3).GetString();
                                    int rowIndex = worksheet.Cell(row, 4).GetValue<int>();
                                    int colIndex = worksheet.Cell(row, 5).GetValue<int>();

                                    var existingRow = ScheduleRows.FirstOrDefault(r =>
                                        r.ScheduleName == scheduleName &&
                                        r.ElementId == elementId &&
                                        r.FieldName == fieldName &&
                                        r.RowIndex == rowIndex &&
                                        r.ColumnIndex == colIndex);

                                    if (existingRow != null && existingRow.Value != newValue)
                                    {
                                        existingRow.Value = newValue;
                                        existingRow.IsModified = true;
                                        updatedCount++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Error importing row {row}: {ex.Message}");
                                }
                            }
                        }

                        StatusMessage = $"Imported {updatedCount} changes from Excel";
                        MessageBox.Show($"Import complete!\n\nFile: {Path.GetFileName(openFileDialog.FileName)}\nUpdated cells: {updatedCount}\n\nModified cells are highlighted in yellow.",
                            "Import Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Import failed: {ex.Message}";
                MessageBox.Show($"Error importing from Excel:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
                LoadingMessage = "";
            }
        }

        private string SanitizeSheetName(string name)
        {
            string sanitized = name;
            char[] invalidChars = new[] { '\\', '/', '?', '*', '[', ']' };
            foreach (char c in invalidChars)
            {
                sanitized = sanitized.Replace(c, '_');
            }
            return sanitized.Length > 31 ? sanitized.Substring(0, 31) : sanitized;
        }
    }
}
