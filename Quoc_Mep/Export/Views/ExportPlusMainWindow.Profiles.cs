using System;
using System.Linq;
using System.Windows;
using Quoc_MEP.Export.Models;
using Quoc_MEP.Export.Managers;

namespace Quoc_MEP.Export.Views
{
    /// <summary>
    /// Profile Management functionality for ExportPlusMainWindow
    /// </summary>
    public partial class ExportPlusMainWindow
    {
        /// <summary>
        /// Initialize Profile Manager and load profiles
        /// </summary>
        private void InitializeProfiles()
        {
            WriteDebugLog("Initializing Profile Manager Service");
            try
            {
                _profileManager = new Managers.ProfileManagerService();
                
                // Wire up profile changed event
                _profileManager.ProfileChanged += OnProfileChanged;
                
                // Bind profiles to ComboBox
                ProfileComboBox.ItemsSource = _profileManager.Profiles;
                ProfileComboBox.SelectedItem = _profileManager.CurrentProfile;
                
                WriteDebugLog($"Profile Manager initialized with {_profileManager.Profiles.Count} profiles");
                WriteDebugLog($"Current profile: {_profileManager.CurrentProfile?.Name}");
                
                // Apply current profile to UI
                if (_profileManager.CurrentProfile != null)
                {
                    ApplyProfileToUI(_profileManager.CurrentProfile);
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error initializing Profile Manager: {ex.Message}");
                MessageBox.Show($"Error initializing profiles: {ex.Message}", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Handle profile change event
        /// </summary>
        private void OnProfileChanged(Profile profile)
        {
            WriteDebugLog($"Profile changed event: {profile?.Name}");
            if (profile != null)
            {
                ApplyProfileToUI(profile);
            }
        }

        /// <summary>
        /// Apply profile settings to UI
        /// QUAN TRỌNG: Profile "Default" KHÔNG apply format selection - để user tự do chọn
        /// </summary>
        private void ApplyProfileToUI(Profile profile)
        {
            if (profile?.Settings == null) return;

            WriteDebugLog($"Applying profile '{profile.Name}' to UI");
            try
            {
                var settings = profile.Settings;

                // Apply Create tab settings
                if (!string.IsNullOrEmpty(settings.OutputFolder))
                {
                    OutputFolder = settings.OutputFolder;
                }

                // ✓ SPECIAL CASE: Profile "Default" KHÔNG khóa format selection
                // User có thể tự do tick chọn format mỗi lần
                bool isDefaultProfile = profile.Name.Equals("Default", StringComparison.OrdinalIgnoreCase);
                
                if (isDefaultProfile)
                {
                    WriteDebugLog("⚠️ Profile 'Default' - SKIPPING format selection apply (user decides)");
                }

                // Apply Format settings (EXCEPT for Default profile)
                if (ExportSettings != null)
                {
                    // Chỉ apply formats nếu KHÔNG phải Default profile
                    if (!isDefaultProfile)
                    {
                        WriteDebugLog($"Applying format settings from profile: PDF={settings.PDFEnabled}, DWG={settings.DWGEnabled}");
                        ExportSettings.IsPdfSelected = settings.PDFEnabled;
                        ExportSettings.IsDwgSelected = settings.DWGEnabled;
                        ExportSettings.IsDgnSelected = settings.DGNEnabled;
                        ExportSettings.IsIfcSelected = settings.IFCEnabled;
                        ExportSettings.IsImgSelected = settings.IMGEnabled;
                    }
                    else
                    {
                        WriteDebugLog($"✓ Keeping current format selection: PDF={ExportSettings.IsPdfSelected}, DWG={ExportSettings.IsDwgSelected}");
                    }
                    
                    // Other settings vẫn apply cho cả Default profile
                    ExportSettings.HideCropBoundaries = settings.HideCropBoundaries;
                    ExportSettings.HideScopeBoxes = settings.HideScopeBoxes;
                    ExportSettings.CreateSeparateFolders = !settings.SaveAllInSameFolder;
                }
                
                // Apply custom file names from XML (if this profile was imported from XML)
                if (!string.IsNullOrEmpty(profile.XmlFilePath))
                {
                    WriteDebugLog($"Profile has XML file path: {profile.XmlFilePath}");
                    
                    if (System.IO.File.Exists(profile.XmlFilePath))
                    {
                        WriteDebugLog($"Loading custom file names from XML: {profile.XmlFilePath}");
                        try
                        {
                            var xmlProfile = XMLProfileManager.LoadProfileFromXML(profile.XmlFilePath);
                            if (xmlProfile != null)
                            {
                                ApplyCustomFileNamesFromXML(xmlProfile);
                                ApplyCustomFileNamesFromXML_Views(xmlProfile);
                                
                                // Convert XML parameters to SelectedParameterInfo and save to profile
                                ConvertAndSaveXMLParametersToProfile(xmlProfile, profile);
                                
                                WriteDebugLog($"Custom file names loaded successfully from XML");
                            }
                            else
                            {
                                WriteDebugLog($"WARNING: Failed to load XML profile from {profile.XmlFilePath}");
                            }
                        }
                        catch (Exception xmlEx)
                        {
                            WriteDebugLog($"ERROR loading XML profile: {xmlEx.Message}");
                        }
                    }
                    else
                    {
                        WriteDebugLog($"WARNING: XML file not found: {profile.XmlFilePath}");
                    }
                }
                else
                {
                    WriteDebugLog($"Profile has no XML file path (not imported from XML)");
                    
                    // Check if profile has saved custom file name configuration
                    bool hasCustomConfig = !string.IsNullOrEmpty(profile.Settings?.CustomFileNameConfigJson);
                    
                    if (hasCustomConfig)
                    {
                        WriteDebugLog($"Profile has saved custom file name configuration (from previous edit)");
                        // Configuration will be automatically loaded when user opens Custom File Name dialog
                        // No need to prompt for XML linking
                        return;
                    }
                    
                    // Don't show dialog for "Default" profile - just use default file names
                    if (profile.Name.Equals("Default", StringComparison.OrdinalIgnoreCase))
                    {
                        WriteDebugLog($"Profile 'Default' has no custom config - will use default file names");
                        return;
                    }
                    
                    // Only ask user to link XML if NO custom configuration exists (and not Default profile)
                    WriteDebugLog($"Profile has no custom file name configuration - asking user to link XML");
                    var result = System.Windows.MessageBox.Show(
                        $"Profile '{profile.Name}' does not have custom file name settings.\n\n" +
                        "Would you like to link an XML profile file to load custom file names?\n\n" +
                        "(This is optional - click 'No' to use default file names)",
                        "Link XML Profile?",
                        System.Windows.MessageBoxButton.YesNo,
                        System.Windows.MessageBoxImage.Question);
                    
                    if (result == System.Windows.MessageBoxResult.Yes)
                    {
                        // Open file dialog
                        var openFileDialog = new Microsoft.Win32.OpenFileDialog
                        {
                            Title = "Select ExportPlus XML Profile",
                            Filter = "XML Profile Files (*.xml)|*.xml|All Files (*.*)|*.*",
                            DefaultExt = ".xml"
                        };
                        
                        if (openFileDialog.ShowDialog() == true)
                        {
                            WriteDebugLog($"User selected XML file: {openFileDialog.FileName}");
                            
                            try
                            {
                                // Load and apply custom file names
                                var xmlProfile = XMLProfileManager.LoadProfileFromXML(openFileDialog.FileName);
                                if (xmlProfile != null)
                                {
                                    // Save XML file path to profile
                                    profile.XmlFilePath = openFileDialog.FileName;
                                    _profileManager.SaveProfile(profile);
                                    WriteDebugLog($"XML file path saved to profile: {openFileDialog.FileName}");
                                    
                                    // Apply custom file names
                                    ApplyCustomFileNamesFromXML(xmlProfile);
                                    ApplyCustomFileNamesFromXML_Views(xmlProfile);
                                    WriteDebugLog($"Custom file names loaded from linked XML");
                                    
                                    System.Windows.MessageBox.Show(
                                        $"XML profile linked successfully!\n" +
                                        $"Custom file names have been applied.",
                                        "Success",
                                        System.Windows.MessageBoxButton.OK,
                                        System.Windows.MessageBoxImage.Information);
                                }
                            }
                            catch (Exception linkEx)
                            {
                                WriteDebugLog($"ERROR linking XML file: {linkEx.Message}");
                                System.Windows.MessageBox.Show(
                                    $"Failed to load XML file:\n{linkEx.Message}",
                                    "Error",
                                    System.Windows.MessageBoxButton.OK,
                                    System.Windows.MessageBoxImage.Error);
                            }
                        }
                        else
                        {
                            WriteDebugLog($"User cancelled XML file selection");
                        }
                    }
                }

                WriteDebugLog($"Profile '{profile.Name}' applied successfully");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error applying profile '{profile?.Name}': {ex.Message}");
            }
        }

        /// <summary>
        /// Save current UI settings to profile
        /// QUAN TRỌNG: Profile "Default" KHÔNG lưu format selection - để user tự do chọn
        /// </summary>
        private void SaveCurrentSettingsToProfile(Profile profile)
        {
            if (profile?.Settings == null) return;

            WriteDebugLog($"Saving current settings to profile '{profile.Name}'");
            try
            {
                var settings = profile.Settings;
                bool isDefaultProfile = profile.Name.Equals("Default", StringComparison.OrdinalIgnoreCase);

                // Save Create tab settings
                settings.OutputFolder = OutputFolder ?? "";
                settings.SaveAllInSameFolder = !(ExportSettings?.CreateSeparateFolders ?? false);

                // Save Format settings (EXCEPT for Default profile)
                if (ExportSettings != null)
                {
                    // ✓ KHÔNG save format selection cho Default profile
                    if (!isDefaultProfile)
                    {
                        WriteDebugLog($"Saving format selection: PDF={ExportSettings.IsPdfSelected}, DWG={ExportSettings.IsDwgSelected}");
                        settings.PDFEnabled = ExportSettings.IsPdfSelected;
                        settings.DWGEnabled = ExportSettings.IsDwgSelected;
                        settings.DGNEnabled = ExportSettings.IsDgnSelected;
                        settings.IFCEnabled = ExportSettings.IsIfcSelected;
                        settings.IMGEnabled = ExportSettings.IsImgSelected;
                    }
                    else
                    {
                        WriteDebugLog("⚠️ Profile 'Default' - SKIPPING format selection save (user decides each time)");
                        // Không thay đổi format settings - giữ nguyên default values
                    }
                    
                    // Other settings vẫn save cho cả Default profile
                    settings.HideCropBoundaries = ExportSettings.HideCropBoundaries;
                    settings.HideScopeBoxes = ExportSettings.HideScopeBoxes;
                }

                // Save Selection settings
                settings.SelectedSheetNumbers = Sheets?
                    .Where(s => s.IsSelected)
                    .Select(s => s.SheetNumber)
                    .ToList() ?? new System.Collections.Generic.List<string>();

                _profileManager.SaveProfile(profile);
                WriteDebugLog($"Profile '{profile.Name}' saved successfully");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error saving settings to profile '{profile?.Name}': {ex.Message}");
            }
        }

        /// <summary>
        /// Profile ComboBox selection changed
        /// </summary>
        private void ProfileComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ProfileComboBox.SelectedItem is Profile selectedProfile)
            {
                WriteDebugLog($"Profile selected: {selectedProfile.Name}");
                _selectedProfile = selectedProfile;
                
                // Enable Apply button when different profile is selected
                if (ApplyProfileButton != null)
                {
                    ApplyProfileButton.IsEnabled = true;
                }
                
                // Don't auto-apply, wait for user to click Apply button
                WriteDebugLog($"Profile '{selectedProfile.Name}' selected. Click Apply to load settings.");
            }
        }

        /// <summary>
        /// Apply selected profile button clicked
        /// </summary>
        private void ApplyProfile_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Apply Profile clicked");
            
            if (ProfileComboBox.SelectedItem is Profile selectedProfile)
            {
                WriteDebugLog($"Applying profile: {selectedProfile.Name}");
                
                try
                {
                    // Switch to selected profile (this will trigger ProfileChanged event)
                    _profileManager.SwitchProfile(selectedProfile);
                    
                    // ✅ FIX: Keep the profile selected in ComboBox after applying
                    ProfileComboBox.SelectedItem = selectedProfile;
                    
                    // Disable Apply button after applying
                    if (ApplyProfileButton != null)
                    {
                        ApplyProfileButton.IsEnabled = false;
                    }
                    
                    WriteDebugLog($"Profile '{selectedProfile.Name}' applied successfully");
                    
                    // Show notification
                    System.Windows.MessageBox.Show(
                        $"Profile '{selectedProfile.Name}' has been applied successfully.",
                        "Profile Applied",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"ERROR applying profile '{selectedProfile.Name}': {ex.Message}");
                    System.Windows.MessageBox.Show(
                        $"Failed to apply profile: {ex.Message}",
                        "Error",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Error);
                }
            }
            else
            {
                WriteDebugLog("No profile selected to apply");
                System.Windows.MessageBox.Show(
                    "Please select a profile first.",
                    "No Profile Selected",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Add new profile button clicked
        /// </summary>
        private void AddProfile_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Add Profile clicked");
            try
            {
                var dialog = new ProfileNameDialog
                {
                    Owner = this
                };
                
                if (dialog.ShowDialog() == true)
                {
                    string profileName = dialog.ProfileName;
                    var mode = dialog.SelectedMode;
                    WriteDebugLog($"Creating new profile: {profileName}, Mode: {mode}");
                    
                    Models.Profile newProfile;
                    
                    switch (mode)
                    {
                        case ProfileNameDialog.ProfileCreationMode.CopyCurrent:
                            // Create profile and copy current settings
                            newProfile = _profileManager.CreateNewProfile(profileName);
                            if (newProfile != null)
                            {
                                SaveCurrentSettingsToProfile(newProfile);
                                WriteDebugLog($"Profile '{profileName}' created with current settings");
                            }
                            break;
                            
                        case ProfileNameDialog.ProfileCreationMode.UseDefault:
                            // Create profile with default settings (empty)
                            newProfile = _profileManager.CreateNewProfile(profileName);
                            if (newProfile != null)
                            {
                                _profileManager.SaveProfile(newProfile);
                                WriteDebugLog($"Profile '{profileName}' created with default settings");
                            }
                            break;
                            
                        case ProfileNameDialog.ProfileCreationMode.ImportFile:
                            // Import from XML file
                            WriteDebugLog($"ImportFile mode - Starting import from: {dialog.ImportFilePath}");
                            newProfile = _profileManager.CreateNewProfile(profileName);
                            if (newProfile != null)
                            {
                                WriteDebugLog($"New profile created with ID: {newProfile.Id}");
                                try
                                {
                                    // Load settings from XML file
                                    WriteDebugLog($"Loading XML profile from: {dialog.ImportFilePath}");
                                    var xmlProfile = XMLProfileManager.LoadProfileFromXML(dialog.ImportFilePath);
                                    WriteDebugLog($"XML profile loaded: {(xmlProfile != null ? "Success" : "NULL")}");
                                    
                                    if (xmlProfile != null && xmlProfile.TemplateInfo != null && newProfile.Settings != null)
                                    {
                                        var template = xmlProfile.TemplateInfo;
                                        WriteDebugLog($"=== IMPORTING XML PROFILE SETTINGS ===");
                                        WriteDebugLog($"Profile Name: {xmlProfile.Name}");
                                        WriteDebugLog($"TemplateInfo - PDF:{template.IsPDFChecked}, DWG:{template.IsDWGChecked}, IFC:{template.IsIFCChecked}");
                                        
                                        // ===== APPLY ALL SETTINGS FROM XML =====
                                        
                                        // Format checkboxes
                                        newProfile.Settings.PDFEnabled = template.IsPDFChecked;
                                        newProfile.Settings.DWGEnabled = template.IsDWGChecked;
                                        newProfile.Settings.DGNEnabled = template.IsDGNChecked;
                                        newProfile.Settings.IFCEnabled = template.IsIFCChecked;
                                        newProfile.Settings.IMGEnabled = template.IsIMGChecked;
                                        WriteDebugLog($"Format flags set - PDF:{template.IsPDFChecked}, DWG:{template.IsDWGChecked}, DGN:{template.IsDGNChecked}, IFC:{template.IsIFCChecked}, IMG:{template.IsIMGChecked}");
                                        
                                        // View options
                                        newProfile.Settings.HideCropBoundaries = template.HideCropBoundaries;
                                        newProfile.Settings.HideScopeBoxes = template.HideScopeBox;
                                        WriteDebugLog($"View options set - HideCropBoundaries:{template.HideCropBoundaries}, HideScopeBoxes:{template.HideScopeBox}");
                                        
                                        // File settings
                                        newProfile.Settings.SaveAllInSameFolder = !template.IsSeparateFile;
                                        if (!string.IsNullOrEmpty(template.FilePath))
                                        {
                                            newProfile.Settings.OutputFolder = template.FilePath;
                                        }
                                        WriteDebugLog($"File settings set - SeparateFiles:{template.IsSeparateFile}, OutputFolder:{template.FilePath}");
                                        
                                        // PDF specific settings
                                        newProfile.Settings.PDFVectorProcessing = template.IsVectorProcessing;
                                        newProfile.Settings.PDFRasterQuality = template.RasterQuality;
                                        newProfile.Settings.PDFColorMode = template.Color;
                                        newProfile.Settings.PDFFitToPage = template.IsFitToPage;
                                        newProfile.Settings.PDFIsCenter = template.IsCenter;
                                        newProfile.Settings.PDFMarginType = template.SelectedMarginType;
                                        WriteDebugLog($"PDF settings set - Vector:{template.IsVectorProcessing}, Quality:{template.RasterQuality}, Color:{template.Color}, FitToPage:{template.IsFitToPage}");
                                        
                                        // DWF settings
                                        if (template.DWF != null)
                                        {
                                            newProfile.Settings.DWFImageFormat = template.DWF.OptImageFormat;
                                            newProfile.Settings.DWFImageQuality = template.DWF.OptImageQuality;
                                            newProfile.Settings.DWFExportTextures = template.DWF.OptExportTextures;
                                            WriteDebugLog($"DWF settings set - Format:{template.DWF.OptImageFormat}, Quality:{template.DWF.OptImageQuality}");
                                        }
                                        
                                        // NWC settings
                                        if (template.NWC != null)
                                        {
                                            newProfile.Settings.NWCConvertElementProperties = template.NWC.ConvertElementProperties;
                                            newProfile.Settings.NWCCoordinates = template.NWC.Coordinates;
                                            newProfile.Settings.NWCDivideFileIntoLevels = template.NWC.DivideFileIntoLevels;
                                            newProfile.Settings.NWCExportElementIds = template.NWC.ExportElementIds;
                                            newProfile.Settings.NWCExportParts = template.NWC.ExportParts;
                                            newProfile.Settings.NWCFacetingFactor = template.NWC.FacetingFactor;
                                            WriteDebugLog($"NWC settings set - Coordinates:{template.NWC.Coordinates}, DivideIntoLevels:{template.NWC.DivideFileIntoLevels}");
                                        }
                                        
                                        // IFC settings
                                        if (template.IFC != null)
                                        {
                                            newProfile.Settings.IFCFileVersion = template.IFC.FileVersion;
                                            newProfile.Settings.IFCSpaceBoundaries = template.IFC.SpaceBoundaries;
                                            newProfile.Settings.IFCSitePlacement = template.IFC.SitePlacement;
                                            newProfile.Settings.IFCExportBaseQuantities = template.IFC.ExportBaseQuantities;
                                            newProfile.Settings.IFCExportIFCCommonPropertySets = template.IFC.ExportIFCCommonPropertySets;
                                            newProfile.Settings.IFCTessellationLevelOfDetail = template.IFC.TessellationLevelOfDetail;
                                            WriteDebugLog($"IFC settings set - Version:{template.IFC.FileVersion}, SpaceBoundaries:{template.IFC.SpaceBoundaries}");
                                        }
                                        
                                        // IMG settings
                                        if (template.IMG != null)
                                        {
                                            newProfile.Settings.IMGImageResolution = template.IMG.ImageResolution;
                                            newProfile.Settings.IMGFileType = template.IMG.HLRandWFViewsFileType;
                                            newProfile.Settings.IMGZoomType = template.IMG.ZoomType;
                                            newProfile.Settings.IMGPixelSize = template.IMG.PixelSize;
                                            WriteDebugLog($"IMG settings set - Resolution:{template.IMG.ImageResolution}, FileType:{template.IMG.HLRandWFViewsFileType}");
                                        }
                                        
                                        // Save profile with all settings
                                        WriteDebugLog($"=== SAVING PROFILE WITH ALL SETTINGS ===");
                                        
                                        // Save XML file path for future re-loading
                                        newProfile.XmlFilePath = dialog.ImportFilePath;
                                        WriteDebugLog($"XML file path saved: {dialog.ImportFilePath}");
                                        
                                        _profileManager.SaveProfile(newProfile);
                                        WriteDebugLog($"Profile '{profileName}' saved successfully with all XML settings applied");
                                        
                                        // Apply settings to UI immediately
                                        WriteDebugLog($"=== APPLYING SETTINGS TO UI ===");
                                        ApplyProfileToUI(newProfile);
                                        WriteDebugLog($"Settings applied to UI successfully");
                                        
                                        // Apply custom file names from XML (if available)
                                        ApplyCustomFileNamesFromXML(xmlProfile);
                                        ApplyCustomFileNamesFromXML_Views(xmlProfile);
                                    }
                                    else
                                    {
                                        WriteDebugLog($"Import failed - xmlProfile:{(xmlProfile != null)}, TemplateInfo:{(xmlProfile?.TemplateInfo != null)}, Settings:{(newProfile.Settings != null)}");
                                        MessageBox.Show("Failed to read settings from XML file.", "Import Error",
                                                       MessageBoxButton.OK, MessageBoxImage.Warning);
                                        return;
                                    }
                                }
                                catch (Exception importEx)
                                {
                                    WriteDebugLog($"EXCEPTION importing XML profile: {importEx.Message}");
                                    WriteDebugLog($"Stack trace: {importEx.StackTrace}");
                                    MessageBox.Show($"Error importing profile: {importEx.Message}", "Import Error",
                                                   MessageBoxButton.OK, MessageBoxImage.Error);
                                    return;
                                }
                            }
                            else
                            {
                                WriteDebugLog($"Failed to create new profile '{profileName}'");
                            }
                            break;
                            
                        default:
                            return;
                    }
                    
                    if (newProfile != null)
                    {
                        // Switch to the new profile
                        _profileManager.SwitchProfile(newProfile);
                        ProfileComboBox.SelectedItem = newProfile;
                        
                        WriteDebugLog($"New profile '{profileName}' created and selected");
                        MessageBox.Show($"Profile '{profileName}' created successfully!", 
                                       "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error creating profile: {ex.Message}");
                MessageBox.Show($"Error creating profile: {ex.Message}", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Save profile button clicked
        /// </summary>
        private void SaveProfile_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Save Profile clicked");
            try
            {
                var currentProfile = ProfileComboBox.SelectedItem as Profile;
                if (currentProfile != null)
                {
                    SaveCurrentSettingsToProfile(currentProfile);
                    
                    WriteDebugLog($"Profile '{currentProfile.Name}' saved");
                    MessageBox.Show($"Profile '{currentProfile.Name}' saved successfully!", 
                                   "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Please select a profile first.", "Information",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error saving profile: {ex.Message}");
                MessageBox.Show($"Error saving profile: {ex.Message}", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Delete profile button clicked
        /// </summary>
        private void DeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            WriteDebugLog("Delete Profile clicked");
            try
            {
                var currentProfile = ProfileComboBox.SelectedItem as Profile;
                if (currentProfile != null)
                {
                    var result = MessageBox.Show(
                        $"Are you sure you want to delete profile '{currentProfile.Name}'?",
                        "Confirm Delete",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        _profileManager.DeleteProfile(currentProfile);
                        WriteDebugLog($"Profile '{currentProfile.Name}' deleted");
                        MessageBox.Show($"Profile '{currentProfile.Name}' deleted successfully!",
                                       "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select a profile first.", "Information",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error deleting profile: {ex.Message}");
                MessageBox.Show($"Error deleting profile: {ex.Message}", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #region Options CheckBox Event Handlers

        /// <summary>
        /// Event handler for Format checkboxes (PDF/DWG/etc.) to log and update queue
        /// </summary>
        private void FormatCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox checkBox)
            {
                bool isChecked = checkBox.IsChecked == true;
                string formatName = checkBox.Name.Replace("Check", ""); // PDFCheck → PDF
                
                WriteDebugLog($"🎨 [FORMAT] {formatName} checkbox changed to: {(isChecked ? "CHECKED ✓" : "UNCHECKED ✗")}");
                
                // Log current state of ALL format checkboxes
                if (ExportSettings != null)
                {
                    WriteDebugLog($"📋 [FORMAT STATE] PDF: {ExportSettings.IsPdfSelected}, " +
                                $"DWG: {ExportSettings.IsDwgSelected}, " +
                                $"DGN: {ExportSettings.IsDgnSelected}, " +
                                $"DWF: {ExportSettings.IsDwfSelected}, " +
                                $"NWC: {ExportSettings.IsNwcSelected}, " +
                                $"IFC: {ExportSettings.IsIfcSelected}, " +
                                $"IMG: {ExportSettings.IsImgSelected}");
                    
                    var selectedFormats = ExportSettings.GetSelectedFormatsList();
                    WriteDebugLog($"✅ [SELECTED FORMATS] Count: {selectedFormats.Count}, List: [{string.Join(", ", selectedFormats)}]");
                }
                
                // Update Export Queue to reflect new format selection
                try
                {
                    UpdateExportQueue();
                    WriteDebugLog($"🔄 Export Queue updated after {formatName} checkbox change");
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"❌ ERROR updating Export Queue: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Event handler for all Options checkboxes to log and track changes
        /// </summary>
        private void OptionCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox checkBox)
            {
                bool isChecked = checkBox.IsChecked == true;
                string optionName = checkBox.Content.ToString();
                
                WriteDebugLog($"[OPTIONS] '{optionName}' changed to: {(isChecked ? "CHECKED ✓" : "UNCHECKED ✗")}");
                
                // Log current state of all options
                if (ExportSettings != null)
                {
                    WriteDebugLog($"[OPTIONS STATE] ViewLinksInBlue: {ExportSettings.ViewLinksInBlue}, " +
                                $"HideRefWorkPlanes: {ExportSettings.HideRefWorkPlanes}, " +
                                $"HideUnreferencedViewTags: {ExportSettings.HideUnreferencedViewTags}, " +
                                $"HideScopeBoxes: {ExportSettings.HideScopeBoxes}, " +
                                $"HideCropBoundaries: {ExportSettings.HideCropBoundaries}, " +
                                $"ReplaceHalftone: {ExportSettings.ReplaceHalftone}, " +
                                $"RegionEdgesMask: {ExportSettings.RegionEdgesMask}");
                }
            }
        }

        #endregion Options CheckBox Event Handlers
        
        #region Custom File Name from XML
        
        /// <summary>
        /// Apply custom file names from XML profile to all sheets
        /// </summary>
        private void ApplyCustomFileNamesFromXML(ExportPlusXMLProfile xmlProfile)
        {
            if (xmlProfile == null || xmlProfile.TemplateInfo == null)
            {
                WriteDebugLog("ApplyCustomFileNamesFromXML: xmlProfile or TemplateInfo is null");
                return;
            }
            
            var template = xmlProfile.TemplateInfo;
            
            // Check if custom file name parameters exist in SelectSheetParameters.CombineParameters
            if (template.SelectSheetParameters?.CombineParameters == null || 
                template.SelectSheetParameters.CombineParameters.Count == 0)
            {
                WriteDebugLog("No custom file name parameters found in XML (SelectSheetParameters.CombineParameters is null or empty)");
                return;
            }
            
            WriteDebugLog($"=== APPLYING CUSTOM FILE NAME PARAMETERS FROM XML (SHEETS) ===");
            WriteDebugLog($"Found {template.SelectSheetParameters.CombineParameters.Count} combine parameters in XML");
            WriteDebugLog($"Combine name: {template.SelectSheetParameters.CombineParameterName}");
            
            try
            {
                // Get all ViewSheet elements from document
                var collector = new Autodesk.Revit.DB.FilteredElementCollector(_document);
                var allSheets = collector
                    .OfClass(typeof(Autodesk.Revit.DB.ViewSheet))
                    .Cast<Autodesk.Revit.DB.ViewSheet>()
                    .Where(s => !s.IsTemplate)
                    .ToList();
                
                WriteDebugLog($"Found {allSheets.Count} sheets in document");
                
                // Apply custom names to sheets in UI
                int appliedCount = 0;
                foreach (var revitSheet in allSheets)
                {
                    // Build custom file name from parameters
                    // Each parameter has its own separator in xml:space_x003D_preserve attribute
                    var customFileNameBuilder = new System.Text.StringBuilder();
                    
                    for (int i = 0; i < template.SelectSheetParameters.CombineParameters.Count; i++)
                    {
                        var param = template.SelectSheetParameters.CombineParameters[i];
                        
                        // Convert ParameterId string to int
                        int paramId = 0;
                        int.TryParse(param.ParameterId, out paramId);
                        
                        // Get parameter value from sheet
                        string value = GetSheetParameterValue(revitSheet, param.ParameterName, paramId);
                        
                        bool isLastParam = (i == template.SelectSheetParameters.CombineParameters.Count - 1);
                        
                        if (!string.IsNullOrEmpty(value))
                        {
                            // Add value to filename
                            customFileNameBuilder.Append(value);
                            
                            // Add separator ONLY if not last parameter AND value is not empty
                            if (!isLastParam)
                            {
                                string separator = param.XmlSpaceAttribute;
                                if (string.IsNullOrEmpty(separator))
                                {
                                    separator = "-"; // Default separator
                                }
                                customFileNameBuilder.Append(separator);
                                WriteDebugLog($"  Param {i+1}: '{param.ParameterName}' = '{value}', separator: '{separator}'");
                            }
                            else
                            {
                                WriteDebugLog($"  Param {i+1}: '{param.ParameterName}' = '{value}' (last, no separator)");
                            }
                        }
                        else
                        {
                            WriteDebugLog($"  Param {i+1}: '{param.ParameterName}' = '' (empty, skipped)");
                        }
                    }
                    
                    string customFileName = customFileNameBuilder.ToString();
                    
                    // Find matching sheet in Sheets collection
                    var sheet = Sheets.FirstOrDefault(s => s.Number == revitSheet.SheetNumber);
                    if (sheet != null)
                    {
                        sheet.CustomFileName = customFileName;
                        appliedCount++;
                        WriteDebugLog($"Applied custom name to sheet {sheet.Number}: '{customFileName}'");
                    }
                }
                
                WriteDebugLog($"Custom file names applied to {appliedCount}/{Sheets.Count} sheets");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR applying custom file names from XML: {ex.Message}");
                WriteDebugLog($"Stack trace: {ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// Apply custom file names from XML profile to all views
        /// </summary>
        private void ApplyCustomFileNamesFromXML_Views(ExportPlusXMLProfile xmlProfile)
        {
            WriteDebugLog("=== ApplyCustomFileNamesFromXML_Views CALLED ===");
            
            if (xmlProfile == null || xmlProfile.TemplateInfo == null)
            {
                WriteDebugLog("ApplyCustomFileNamesFromXML_Views: xmlProfile or TemplateInfo is null");
                return;
            }
            
            var template = xmlProfile.TemplateInfo;
            
            // Check if custom file name parameters exist in SelectViewParameters.CombineParameters
            if (template.SelectViewParameters?.CombineParameters == null || 
                template.SelectViewParameters.CombineParameters.Count == 0)
            {
                WriteDebugLog("⚠️ No custom file name parameters found in XML for Views (SelectViewParameters.CombineParameters is null or empty)");
                WriteDebugLog($"SelectViewParameters is null: {template.SelectViewParameters == null}");
                if (template.SelectViewParameters != null)
                {
                    WriteDebugLog($"CombineParameters is null: {template.SelectViewParameters.CombineParameters == null}");
                    WriteDebugLog($"CombineParameters count: {template.SelectViewParameters.CombineParameters?.Count ?? 0}");
                }
                return;
            }
            
            WriteDebugLog($"✓ Found {template.SelectViewParameters.CombineParameters.Count} combine parameters in XML for Views");
            WriteDebugLog($"✓ Combine name: {template.SelectViewParameters.CombineParameterName}");
            
            // Check if Views collection is loaded - if not, load it first
            if (Views == null || Views.Count == 0)
            {
                WriteDebugLog("⚠️ WARNING: Views collection is null or empty! Loading Views now...");
                LoadViews(); // Force load Views collection
                
                if (Views == null || Views.Count == 0)
                {
                    WriteDebugLog("❌ ERROR: Failed to load Views collection. Cannot apply custom file names.");
                    return;
                }
                WriteDebugLog($"✓ Views collection loaded successfully with {Views.Count} items");
            }
            else
            {
                WriteDebugLog($"✓ Views collection already has {Views.Count} items");
            }
            
            try
            {
                // Get all View elements from document (3D views, sections, elevations, etc.)
                var collector = new Autodesk.Revit.DB.FilteredElementCollector(_document);
                var allViews = collector
                    .OfClass(typeof(Autodesk.Revit.DB.View))
                    .Cast<Autodesk.Revit.DB.View>()
                    .Where(v => !v.IsTemplate && v.CanBePrinted) // Only non-template, printable views
                    .ToList();
                
                WriteDebugLog($"✓ Found {allViews.Count} printable views in Revit document");
                
                // Apply custom names to views in UI
                int appliedCount = 0;
                int matchedCount = 0;
                
                foreach (var revitView in allViews)
                {
                    // Build custom file name from parameters
                    var customFileNameBuilder = new System.Text.StringBuilder();
                    
                    WriteDebugLog($"Processing view: '{revitView.Name}'");
                    
                    for (int i = 0; i < template.SelectViewParameters.CombineParameters.Count; i++)
                    {
                        var param = template.SelectViewParameters.CombineParameters[i];
                        
                        // Convert ParameterId string to int
                        int paramId = 0;
                        int.TryParse(param.ParameterId, out paramId);
                        
                        // Get parameter value from view
                        string value = GetViewParameterValue(revitView, param.ParameterName, paramId);
                        
                        bool isLastParam = (i == template.SelectViewParameters.CombineParameters.Count - 1);
                        
                        if (!string.IsNullOrEmpty(value))
                        {
                            // Add value to filename
                            customFileNameBuilder.Append(value);
                            
                            // Add separator ONLY if not last parameter AND value is not empty
                            if (!isLastParam)
                            {
                                string separator = param.XmlSpaceAttribute;
                                if (string.IsNullOrEmpty(separator))
                                {
                                    separator = "_"; // Default separator
                                }
                                customFileNameBuilder.Append(separator);
                                WriteDebugLog($"  Param {i+1}: '{param.ParameterName}' (ID:{paramId}) = '{value}', separator: '{separator}'");
                            }
                            else
                            {
                                WriteDebugLog($"  Param {i+1}: '{param.ParameterName}' (ID:{paramId}) = '{value}' (last, no separator)");
                            }
                        }
                        else
                        {
                            WriteDebugLog($"  Param {i+1}: '{param.ParameterName}' (ID:{paramId}) = '' (empty, skipped)");
                        }
                    }
                    
                    string customFileName = customFileNameBuilder.ToString();
                    WriteDebugLog($"  → Built custom filename: '{customFileName}'");
                    
                    // Find matching view in Views collection
                    var viewItem = Views?.FirstOrDefault(v => v.ViewName == revitView.Name);
                    if (viewItem != null)
                    {
                        matchedCount++;
                        viewItem.CustomFileName = customFileName;
                        appliedCount++;
                        WriteDebugLog($"  ✓ Applied to ViewItem: '{viewItem.ViewName}' → '{customFileName}'");
                    }
                    else
                    {
                        WriteDebugLog($"  ✗ No matching ViewItem found for '{revitView.Name}'");
                    }
                }
                
                WriteDebugLog($"=== SUMMARY ===");
                WriteDebugLog($"Revit views found: {allViews.Count}");
                WriteDebugLog($"Views in UI collection: {Views?.Count ?? 0}");
                WriteDebugLog($"Matched views: {matchedCount}");
                WriteDebugLog($"Custom file names applied: {appliedCount}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"❌ ERROR applying custom file names from XML to Views: {ex.Message}");
                WriteDebugLog($"Stack trace: {ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// Get parameter value from sheet by name or ID
        /// </summary>
        private string GetSheetParameterValue(Autodesk.Revit.DB.ViewSheet sheet, string parameterName, int parameterId)
        {
            try
            {
                // Try by parameter ID first (more reliable)
                var param = sheet.get_Parameter((Autodesk.Revit.DB.BuiltInParameter)parameterId);
                if (param != null && param.HasValue)
                {
                    return param.AsString() ?? param.AsValueString() ?? "";
                }
                
                // Try by parameter name as fallback
                param = sheet.LookupParameter(parameterName);
                if (param != null && param.HasValue)
                {
                    return param.AsString() ?? param.AsValueString() ?? "";
                }
                
                // Try from ProjectInfo for project-level parameters (Project Name, Project Number, etc.)
                var projectInfo = _document.ProjectInformation;
                if (projectInfo != null)
                {
                    // Try by parameter ID
                    if (parameterId != 0)
                    {
                        param = projectInfo.get_Parameter((Autodesk.Revit.DB.BuiltInParameter)parameterId);
                        if (param != null && param.HasValue)
                        {
                            return param.AsString() ?? param.AsValueString() ?? "";
                        }
                    }
                    
                    // Try by parameter name
                    param = projectInfo.LookupParameter(parameterName);
                    if (param != null && param.HasValue)
                    {
                        return param.AsString() ?? param.AsValueString() ?? "";
                    }
                }
                
                return "";
            }
            catch
            {
                return "";
            }
        }
        
        /// <summary>
        /// Get parameter value from view by name or ID
        /// </summary>
        private string GetViewParameterValue(Autodesk.Revit.DB.View view, string parameterName, int parameterId)
        {
            try
            {
                // Try by parameter ID first (more reliable)
                var param = view.get_Parameter((Autodesk.Revit.DB.BuiltInParameter)parameterId);
                if (param != null && param.HasValue)
                {
                    return param.AsString() ?? param.AsValueString() ?? "";
                }
                
                // Try by parameter name as fallback
                param = view.LookupParameter(parameterName);
                if (param != null && param.HasValue)
                {
                    return param.AsString() ?? param.AsValueString() ?? "";
                }
                
                // Try from ProjectInfo for project-level parameters (Project Name, Project Number, etc.)
                var projectInfo = _document.ProjectInformation;
                if (projectInfo != null)
                {
                    // Try by parameter ID
                    if (parameterId != 0)
                    {
                        param = projectInfo.get_Parameter((Autodesk.Revit.DB.BuiltInParameter)parameterId);
                        if (param != null && param.HasValue)
                        {
                            return param.AsString() ?? param.AsValueString() ?? "";
                        }
                    }
                    
                    // Try by parameter name
                    param = projectInfo.LookupParameter(parameterName);
                    if (param != null && param.HasValue)
                    {
                        return param.AsString() ?? param.AsValueString() ?? "";
                    }
                }
                
                return "";
            }
            catch
            {
                return "";
            }
        }
        
        /// <summary>
        /// Convert XML profile parameters to SelectedParameterInfo and save to profile settings
        /// This allows the CustomFileNameDialog to load the correct parameter order
        /// </summary>
        private void ConvertAndSaveXMLParametersToProfile(ExportPlusXMLProfile xmlProfile, Profile profile)
        {
            try
            {
                WriteDebugLog("=== Converting XML Parameters to Profile Configuration ===");
                
                // Check if we have Sheet or View parameters
                bool hasSheetParams = xmlProfile?.TemplateInfo?.SelectSheetParameters?.CombineParameters != null 
                                      && xmlProfile.TemplateInfo.SelectSheetParameters.CombineParameters.Count > 0;
                bool hasViewParams = xmlProfile?.TemplateInfo?.SelectViewParameters?.CombineParameters != null 
                                     && xmlProfile.TemplateInfo.SelectViewParameters.CombineParameters.Count > 0;
                
                if (!hasSheetParams && !hasViewParams)
                {
                    WriteDebugLog("No parameters found in XML profile");
                    return;
                }
                
                if (profile?.Settings == null)
                {
                    WriteDebugLog("⚠️ WARNING: Profile or Settings is null, cannot save configuration");
                    return;
                }
                
                // Convert Sheet parameters if available
                if (hasSheetParams)
                {
                    var sheetParams = new System.Collections.Generic.List<Models.SelectedParameterInfo>();
                    var sourceParams = xmlProfile.TemplateInfo.SelectSheetParameters.CombineParameters;
                    
                    WriteDebugLog($"Converting {sourceParams.Count} Sheet parameters from XML");
                    
                    foreach (var xmlParam in sourceParams)
                    {
                        var selectedParam = new Models.SelectedParameterInfo
                        {
                            ParameterName = xmlParam.ParameterName,
                            Prefix = xmlParam.Prefix ?? "",
                            Suffix = xmlParam.Suffix ?? "",
                            Separator = xmlParam.XmlSpaceAttribute ?? "_", // Default separator
                            SampleValue = "" // Will be filled by dialog preview
                        };
                        
                        sheetParams.Add(selectedParam);
                        WriteDebugLog($"  Sheet Param: {xmlParam.ParameterName} (separator: '{selectedParam.Separator}')");
                    }
                    
                    // Serialize and save Sheet config
                    var sheetConfigJson = Newtonsoft.Json.JsonConvert.SerializeObject(sheetParams);
                    profile.Settings.CustomFileNameConfigJson_Sheets = sheetConfigJson;
                    
                    WriteDebugLog($"✓ Saved {sheetParams.Count} Sheet parameters to profile");
                }
                
                // Convert View parameters if available
                if (hasViewParams)
                {
                    var viewParams = new System.Collections.Generic.List<Models.SelectedParameterInfo>();
                    var sourceParams = xmlProfile.TemplateInfo.SelectViewParameters.CombineParameters;
                    
                    WriteDebugLog($"Converting {sourceParams.Count} View parameters from XML");
                    
                    foreach (var xmlParam in sourceParams)
                    {
                        var selectedParam = new Models.SelectedParameterInfo
                        {
                            ParameterName = xmlParam.ParameterName,
                            Prefix = xmlParam.Prefix ?? "",
                            Suffix = xmlParam.Suffix ?? "",
                            Separator = xmlParam.XmlSpaceAttribute ?? "_", // Default separator
                            SampleValue = "" // Will be filled by dialog preview
                        };
                        
                        viewParams.Add(selectedParam);
                        WriteDebugLog($"  View Param: {xmlParam.ParameterName} (separator: '{selectedParam.Separator}')");
                    }
                    
                    // Serialize and save View config
                    var viewConfigJson = Newtonsoft.Json.JsonConvert.SerializeObject(viewParams);
                    profile.Settings.CustomFileNameConfigJson_Views = viewConfigJson;
                    
                    WriteDebugLog($"✓ Saved {viewParams.Count} View parameters to profile");
                }
                
                // Save profile to disk
                _profileManager.SaveProfile(profile);
                WriteDebugLog("✓ Profile saved to disk with separate Sheet and View configurations");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR converting XML parameters to profile: {ex.Message}");
            }
        }
        
        #endregion Custom File Name from XML
    }
}
