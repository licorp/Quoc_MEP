using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Quoc_MEP.Export.Models;
using Autodesk.Revit.DB;

namespace Quoc_MEP.Export.Managers
{
    /// <summary>
    /// Profile Manager for loading and managing ExportPlus profiles
    /// Compatible with original DiRoots ExportPlus profile format
    /// </summary>
    public class ExportPlusProfileManager
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern void OutputDebugStringA(string message);

        private static void WriteDebugLog(string message)
        {
            string logMessage = $"[ExportPlus ProfileManager] {DateTime.Now:HH:mm:ss.fff} - {message}";
            OutputDebugStringA(logMessage);
        }

        private readonly string _ExportPlusProfileFolder;
        private readonly string _exportPlusProfileFolder;

        public ObservableCollection<ExportPlusProfile> Profiles { get; set; }

        public ExportPlusProfileManager()
        {
            WriteDebugLog("ProfileManager constructor started");

            // Original ExportPlus profile folder
            _ExportPlusProfileFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "DiRoots", "ExportPlus", "Profiles"
            );

            // ExportPlus profile folder (our custom folder)
            _exportPlusProfileFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "DiRoots", "ExportPlus", "Profiles"
            );

            Profiles = new ObservableCollection<ExportPlusProfile>();
            
            WriteDebugLog($"ExportPlus folder: {_ExportPlusProfileFolder}");
            WriteDebugLog($"ExportPlus folder: {_exportPlusProfileFolder}");

            EnsureDirectoriesExist();
            LoadProfiles();
        }

        private void EnsureDirectoriesExist()
        {
            try
            {
                if (!Directory.Exists(_exportPlusProfileFolder))
                {
                    Directory.CreateDirectory(_exportPlusProfileFolder);
                    WriteDebugLog($"Created ExportPlus profiles directory: {_exportPlusProfileFolder}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error creating directories: {ex.Message}");
            }
        }

        /// <summary>
        /// Load profiles from both ExportPlus and ExportPlus directories
        /// </summary>
        public void LoadProfiles()
        {
            WriteDebugLog("LoadProfiles started");
            Profiles.Clear();

            // Load original ExportPlus profiles if available
            LoadExportPlusProfiles();

            // Load existing ExportPlus profiles from user folder
            LoadExistingExportPlusProfiles();

            // Add default profile if no profiles found
            if (Profiles.Count == 0)
            {
                var defaultProfile = CreateDefaultProfile();
                Profiles.Add(defaultProfile);
                WriteDebugLog("Added default profile");
            }

            WriteDebugLog($"Total profiles loaded: {Profiles.Count}");
        }

        public void LoadExportPlusProfile(string jsonFilePath)
        {
            try
            {
                if (File.Exists(jsonFilePath))
                {
                    WriteDebugLog($"Loading ExportPlus profile from: {jsonFilePath}");
                    
                    // Check if it's XML or JSON
                    string extension = Path.GetExtension(jsonFilePath).ToLower();
                    ExportPlusProfile profile = null;
                    
                    if (extension == ".xml")
                    {
                        // Load XML profile and convert to standard profile
                        var xmlProfile = XMLProfileManager.LoadProfileFromXML(jsonFilePath);
                        if (xmlProfile != null)
                        {
                            profile = XMLProfileManager.ConvertXMLToProfile(xmlProfile);
                        }
                    }
                    else if (extension == ".json")
                    {
                        // Load JSON profile
                        string json = File.ReadAllText(jsonFilePath);
                        profile = JsonConvert.DeserializeObject<ExportPlusProfile>(json);
                    }
                    
                    if (profile != null)
                    {
                        if (string.IsNullOrEmpty(profile.ProfileName))
                        {
                            profile.ProfileName = Path.GetFileNameWithoutExtension(jsonFilePath);
                        }
                        
                        // Check if profile already exists
                        var existingProfile = Profiles.FirstOrDefault(p => p.ProfileName == profile.ProfileName);
                        if (existingProfile != null)
                        {
                            Profiles.Remove(existingProfile);
                        }
                        
                        Profiles.Add(profile);
                        WriteDebugLog($"Profile loaded: {profile.ProfileName}");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error loading ExportPlus profile: {ex.Message}");
            }
        }

        /// <summary>
        /// Load XML profile specifically and return SheetFileNameInfo for UI binding
        /// </summary>
        public List<SheetFileNameInfo> LoadXMLProfileWithSheets(string xmlFilePath, List<ViewSheet> sheets)
        {
            try
            {
                WriteDebugLog($"Loading XML profile with sheets: {xmlFilePath}");
                var xmlProfile = XMLProfileManager.LoadProfileFromXML(xmlFilePath);
                if (xmlProfile != null)
                {
                    return XMLProfileManager.GenerateCustomFileNames(xmlProfile, sheets);
                }
                return new List<SheetFileNameInfo>();
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error loading XML profile with sheets: {ex.Message}");
                return new List<SheetFileNameInfo>();
            }
        }

        /// <summary>
        /// Get available XML profiles from DiRoots and our folder
        /// </summary>
        public List<string> GetAvailableXMLProfiles()
        {
            return XMLProfileManager.GetAvailableXMLProfiles();
        }

        private void LoadExportPlusProfiles()
        {
            try
            {
                if (Directory.Exists(_ExportPlusProfileFolder))
                {
                    var jsonFiles = Directory.GetFiles(_ExportPlusProfileFolder, "*.json");
                    WriteDebugLog($"Found {jsonFiles.Length} ExportPlus profile files");

                    foreach (var file in jsonFiles)
                    {
                        try
                        {
                            string json = File.ReadAllText(file);
                            var profile = JsonConvert.DeserializeObject<ExportPlusProfile>(json);
                            
                            if (profile != null)
                            {
                                if (string.IsNullOrEmpty(profile.ProfileName))
                                {
                                    profile.ProfileName = Path.GetFileNameWithoutExtension(file);
                                }
                                
                                Profiles.Add(profile);
                                WriteDebugLog($"Loaded ExportPlus profile: {profile.ProfileName}");
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteDebugLog($"Error loading ExportPlus profile {file}: {ex.Message}");
                        }
                    }
                }
                else
                {
                    WriteDebugLog("ExportPlus profiles directory not found");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error accessing ExportPlus profiles directory: {ex.Message}");
            }
        }

        private void LoadExistingExportPlusProfiles()
        {
            try
            {
                if (Directory.Exists(_exportPlusProfileFolder))
                {
                    var jsonFiles = Directory.GetFiles(_exportPlusProfileFolder, "*.json");
                    WriteDebugLog($"Found {jsonFiles.Length} ExportPlus profile files");

                    foreach (var file in jsonFiles)
                    {
                        try
                        {
                            string json = File.ReadAllText(file);
                            var profile = JsonConvert.DeserializeObject<ExportPlusProfile>(json);
                            
                            if (profile != null)
                            {
                                if (string.IsNullOrEmpty(profile.ProfileName))
                                {
                                    profile.ProfileName = Path.GetFileNameWithoutExtension(file);
                                }
                                
                                Profiles.Add(profile);
                                WriteDebugLog($"Loaded ExportPlus profile: {profile.ProfileName}");
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteDebugLog($"Error loading ExportPlus profile {file}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error accessing ExportPlus profiles directory: {ex.Message}");
            }
        }

        /// <summary>
        /// Save profile to ExportPlus directory
        /// </summary>
        public void SaveProfile(ExportPlusProfile profile)
        {
            try
            {
                if (profile == null || string.IsNullOrEmpty(profile.ProfileName))
                {
                    WriteDebugLog("Cannot save profile: profile is null or name is empty");
                    return;
                }

                string fileName = $"{profile.ProfileName}.json";
                string filePath = Path.Combine(_exportPlusProfileFolder, fileName);
                
                string json = JsonConvert.SerializeObject(profile, Formatting.Indented);
                File.WriteAllText(filePath, json);
                
                WriteDebugLog($"Profile saved: {filePath}");

                // Add to collection if not already present
                if (!Profiles.Any(p => p.ProfileName == profile.ProfileName))
                {
                    Profiles.Add(profile);
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error saving profile: {ex.Message}");
            }
        }

        /// <summary>
        /// Delete profile (only ExportPlus profiles can be deleted)
        /// </summary>
        public void DeleteProfile(ExportPlusProfile profile)
        {
            try
            {
                if (profile == null || string.IsNullOrEmpty(profile.ProfileName))
                {
                    WriteDebugLog("Cannot delete profile: profile is null or name is empty");
                    return;
                }

                string fileName = $"{profile.ProfileName}.json";
                string filePath = Path.Combine(_exportPlusProfileFolder, fileName);
                
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Profiles.Remove(profile);
                    WriteDebugLog($"Profile deleted: {profile.ProfileName}");
                }
                else
                {
                    WriteDebugLog($"Profile file not found for deletion: {filePath}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error deleting profile: {ex.Message}");
            }
        }

        private ExportPlusProfile CreateDefaultProfile()
        {
            return new ExportPlusProfile
            {
                ProfileName = "Default ExportPlus",
                OutputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                SelectedFormats = new List<string> { "DWG" },
                CreateSeparateFolders = false,
                PaperSize = "Auto",
                Orientation = "Auto",
                PlaceCenterDrawing = true,
                ZoomTo100 = false,
                HideCropRegions = true,
                HideScopeBoxes = true
            };
        }

        /// <summary>
        /// Get profile by name
        /// </summary>
        public ExportPlusProfile GetProfile(string profileName)
        {
            return Profiles.FirstOrDefault(p => p.ProfileName == profileName);
        }

        /// <summary>
        /// Create new profile from current settings
        /// </summary>
        public ExportPlusProfile CreateProfileFromSettings(ExportSettings settings, string profileName)
        {
            var profile = new ExportPlusProfile
            {
                ProfileName = profileName,
                OutputFolder = settings?.OutputFolder ?? "",
                SelectedFormats = settings?.GetSelectedFormatsList() ?? new List<string>(),
                CreateSeparateFolders = settings?.CreateSeparateFolders ?? false,
                HideCropRegions = settings?.HideCropBoundaries ?? true,
                HideScopeBoxes = settings?.HideScopeBoxes ?? true
            };

            return profile;
        }

        /// <summary>
        /// Load profile from external JSON file (import)
        /// </summary>
        public ExportPlusProfile LoadProfileFromFile(string jsonFilePath)
        {
            try
            {
                if (!File.Exists(jsonFilePath))
                {
                    WriteDebugLog($"Profile file not found: {jsonFilePath}");
                    return null;
                }

                string json = File.ReadAllText(jsonFilePath);
                var profile = JsonConvert.DeserializeObject<ExportPlusProfile>(json);
                
                if (profile != null)
                {
                    WriteDebugLog($"Profile loaded from file: {profile.ProfileName}");
                    return profile;
                }
                
                return null;
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error loading profile from file: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Export profile to external JSON file
        /// </summary>
        public void ExportProfileToFile(ExportPlusProfile profile, string jsonFilePath)
        {
            try
            {
                if (profile == null)
                {
                    WriteDebugLog("Cannot export profile: profile is null");
                    return;
                }

                string json = JsonConvert.SerializeObject(profile, Formatting.Indented);
                File.WriteAllText(jsonFilePath, json);
                
                WriteDebugLog($"Profile exported to: {jsonFilePath}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error exporting profile to file: {ex.Message}");
                throw;
            }
        }
    }
}