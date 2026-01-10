using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Windows.Interop;

namespace ScheduleManager
{
    /// <summary>
    /// External Command to launch Schedule Manager
    /// Opens WPF window with Excel export/import capabilities
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ScheduleManagerCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;

                // Create ViewModel
                var viewModel = new ScheduleManagerViewModel(uiApp);

                // Create and show window
                var window = new ScheduleManagerWindow
                {
                    DataContext = viewModel
                };

                // Set Revit as parent window
                WindowInteropHelper helper = new WindowInteropHelper(window);
                helper.Owner = commandData.Application.MainWindowHandle;

                // Show as dialog
                window.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                TaskDialog.Show("Schedule Manager Error", $"Error loading Schedule Manager:\n\n{ex.Message}\n\n{ex.StackTrace}");
                return Result.Failed;
            }
        }
    }
}
