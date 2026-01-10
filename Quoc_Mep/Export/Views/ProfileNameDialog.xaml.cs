using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Win32;

namespace Quoc_MEP.Export.Views
{
    /// <summary>
    /// Dialog for creating new profile with multiple options
    /// </summary>
    public partial class ProfileNameDialog : Window
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
        private static extern void OutputDebugStringA(string lpOutputString);

        private static void WriteDebugLog(string message)
        {
            try
            {
                string logMessage = $"[ExportPlus ProfileDialog] {DateTime.Now:HH:mm:ss.fff} - {message}";
                OutputDebugStringA(logMessage);
                Debug.WriteLine(logMessage);
            }
            catch { }
        }

        public string ProfileName { get; private set; }
        public string ImportFilePath { get; private set; }
        
        public enum ProfileCreationMode 
        { 
            CopyCurrent, 
            UseDefault, 
            ImportFile 
        }
        
        public ProfileCreationMode SelectedMode { get; private set; }

        public ProfileNameDialog()
        {
            InitializeComponent();
            WriteDebugLog("ProfileNameDialog initialized");
            ProfileNameTextBox.Focus();
            
            // Wire up validation events
            ProfileNameTextBox.TextChanged += (s, e) => 
            {
                WriteDebugLog($"Profile name changed: '{ProfileNameTextBox.Text}'");
                ValidateInputs();
            };
            CopyCurrentRadio.Checked += (s, e) => 
            {
                WriteDebugLog("CopyCurrent mode selected");
                ValidateInputs();
            };
            UseDefaultRadio.Checked += (s, e) => 
            {
                WriteDebugLog("UseDefault mode selected");
                ValidateInputs();
            };
            ImportFileRadio.Checked += (s, e) => 
            {
                WriteDebugLog("ImportFile mode selected");
                ValidateInputs();
            };
        }

        private void BrowseProfileFileButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Browse button clicked");
            var dlg = new OpenFileDialog
            {
                Filter = "ExportPlus Profile (*.xml)|*.xml|All files (*.*)|*.*",
                Title = "Import Profile File"
            };
            
            if (dlg.ShowDialog() == true)
            {
                ImportFilePath = dlg.FileName;
                WriteDebugLog($"File selected: {ImportFilePath}");
                
                // Auto-fill profile name from file name if textbox is empty
                if (string.IsNullOrWhiteSpace(ProfileNameTextBox.Text))
                {
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(ImportFilePath);
                    ProfileNameTextBox.Text = fileNameWithoutExtension;
                    WriteDebugLog($"Auto-filled profile name: '{fileNameWithoutExtension}'");
                }
                
                ImportedFilePathTextBlock.Text = $"📄 {Path.GetFileName(ImportFilePath)}";
                ImportFileRadio.IsChecked = true;
                ValidateInputs();
            }
            else
            {
                WriteDebugLog("File selection canceled");
            }
        }

        private void ValidateInputs()
        {
            bool hasName = !string.IsNullOrWhiteSpace(ProfileNameTextBox.Text);
            
            WriteDebugLog($"ValidateInputs - HasName: {hasName}, Create button enabled: {hasName}");
            
            // Enable Create button if name is entered (we'll validate file path on Create click)
            CreateButton.IsEnabled = hasName;
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Create button clicked");
            ProfileName = ProfileNameTextBox.Text?.Trim();

            if (string.IsNullOrWhiteSpace(ProfileName))
            {
                WriteDebugLog("Validation failed: Profile name is empty");
                MessageBox.Show("Please enter a profile name.", "Validation Error", 
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                ProfileNameTextBox.Focus();
                return;
            }

            // Validate profile name (no special characters)
            if (ProfileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                WriteDebugLog("Validation failed: Profile name contains invalid characters");
                MessageBox.Show("Profile name contains invalid characters.", "Validation Error", 
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                ProfileNameTextBox.Focus();
                return;
            }

            // Determine selected mode
            if (CopyCurrentRadio.IsChecked == true)
            {
                SelectedMode = ProfileCreationMode.CopyCurrent;
                WriteDebugLog($"Mode selected: CopyCurrent, Profile name: '{ProfileName}'");
            }
            else if (UseDefaultRadio.IsChecked == true)
            {
                SelectedMode = ProfileCreationMode.UseDefault;
                WriteDebugLog($"Mode selected: UseDefault, Profile name: '{ProfileName}'");
            }
            else if (ImportFileRadio.IsChecked == true)
            {
                SelectedMode = ProfileCreationMode.ImportFile;
                
                // Validate that a file was selected for import mode
                if (string.IsNullOrEmpty(ImportFilePath))
                {
                    WriteDebugLog("Validation failed: Import mode selected but no file chosen");
                    MessageBox.Show("Please select a file to import by clicking the '...' button.", 
                                   "File Required", 
                                   MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                WriteDebugLog($"Mode selected: ImportFile, Profile name: '{ProfileName}', File: '{ImportFilePath}'");
            }

            WriteDebugLog($"Profile creation confirmed - Mode: {SelectedMode}, Name: '{ProfileName}'");
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Cancel button clicked - Profile creation canceled");
            DialogResult = false;
            Close();
        }
    }
}
