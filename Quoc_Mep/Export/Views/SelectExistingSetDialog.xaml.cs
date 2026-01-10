using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Quoc_MEP.Export.Models;

namespace Quoc_MEP.Export.Views
{
    /// <summary>
    /// Dialog to select an existing View/Sheet Set to add items to
    /// </summary>
    public partial class SelectExistingSetDialog : Window
    {
        public string SelectedSetName { get; private set; }

        public SelectExistingSetDialog(IEnumerable<ViewSheetSetInfo> availableSets)
        {
            InitializeComponent();

            // Filter out built-in sets (All Sheets, All Views)
            var editableSets = availableSets.Where(s => !s.IsBuiltIn).ToList();

            if (!editableSets.Any())
            {
                MessageBox.Show(
                    "No custom View/Sheet Sets found. Please create a new set first.",
                    "No Sets Available",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                DialogResult = false;
                Close();
                return;
            }

            ExistingSetsCombo.ItemsSource = editableSets;
            ExistingSetsCombo.SelectedIndex = 0;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (ExistingSetsCombo.SelectedValue is string selectedName)
            {
                SelectedSetName = selectedName;
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show(
                    "Please select a View/Sheet Set.",
                    "Selection Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
