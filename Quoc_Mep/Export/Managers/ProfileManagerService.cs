using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using System.Windows;
using Quoc_MEP.Export.Models;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Quoc_MEP.Export.Managers
{
    /// <summary>
    /// Manages profile creation, loading, saving, and switching
    /// </summary>
    public class ProfileManagerService
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
        private static extern void OutputDebugStringA(string lpOutputString);

        private readonly string _profilesFolder;
        private const string PROFILES_FOLDER = "ExportPlusProfiles";
        private const string DEFAULT_PROFILE = "Default";

        public ObservableCollection<Profile> Profiles { get; private set; }
        public Profile CurrentProfile { get; private set; }

        public event Action<Profile> ProfileChanged;

        public ProfileManagerService()
        {
            _profilesFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                PROFILES_FOLDER);

            Directory.CreateDirectory(_profilesFolder);
            Profiles = new ObservableCollection<Profile>();
            LoadProfiles();
            
            WriteDebugLog($"ProfileManager initialized - Profiles folder: {_profilesFolder}");
        }

        /// <summary>
        /// Load all profiles from disk
        /// </summary>
        public void LoadProfiles()
        {
            Profiles.Clear();

            try
            {
                var profileFiles = Directory.GetFiles(_profilesFolder, "*.json");
                WriteDebugLog($"Found {profileFiles.Length} profile files");
                
                foreach (var file in profileFiles)
                {
                    try
                    {
                        var json = File.ReadAllText(file);
                        var profile = JsonConvert.DeserializeObject<Profile>(json);
                        if (profile != null)
                        {
                            Profiles.Add(profile);
                            WriteDebugLog($"Loaded profile: {profile.Name}");
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteDebugLog($"Error loading profile {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                // Create default profile if none exist
                if (!Profiles.Any())
                {
                    WriteDebugLog("No profiles found - creating default profile");
                    CreateDefaultProfile();
                }

                // Set current profile to Default or first available
                CurrentProfile = Profiles.FirstOrDefault(p => p.Name == DEFAULT_PROFILE) 
                               ?? Profiles.FirstOrDefault();
                
                WriteDebugLog($"Current profile: {CurrentProfile?.Name}");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error in LoadProfiles: {ex.Message}");
                MessageBox.Show($"Error loading profiles: {ex.Message}", "Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
                CreateDefaultProfile();
            }
        }

        /// <summary>
        /// Save profile to disk
        /// </summary>
        public void SaveProfile(Profile profile)
        {
            try
            {
                profile.LastModified = DateTime.Now;
                var json = JsonConvert.SerializeObject(profile, Formatting.Indented);

                var filePath = Path.Combine(_profilesFolder, $"{profile.Name}.json");
                File.WriteAllText(filePath, json);
                
                WriteDebugLog($"Profile saved: {profile.Name} at {filePath}");

                // Update in collection if exists, otherwise add
                var existing = Profiles.FirstOrDefault(p => p.Id == profile.Id);
                if (existing != null)
                {
                    var index = Profiles.IndexOf(existing);
                    Profiles[index] = profile;
                    WriteDebugLog($"Updated existing profile in collection: {profile.Name}");
                }
                else
                {
                    Profiles.Add(profile);
                    WriteDebugLog($"Added new profile to collection: {profile.Name}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error saving profile {profile?.Name}: {ex.Message}");
                MessageBox.Show($"Error saving profile: {ex.Message}", "Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Delete profile from disk and collection
        /// </summary>
        public void DeleteProfile(Profile profile)
        {
            if (profile.Name == DEFAULT_PROFILE)
            {
                MessageBox.Show("Cannot delete the default profile.", "Warning", 
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var filePath = Path.Combine(_profilesFolder, $"{profile.Name}.json");
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    WriteDebugLog($"Deleted profile file: {filePath}");
                }

                Profiles.Remove(profile);
                WriteDebugLog($"Removed profile from collection: {profile.Name}");

                // Switch to default if deleted profile was current
                if (CurrentProfile?.Id == profile.Id)
                {
                    var defaultProfile = Profiles.FirstOrDefault(p => p.Name == DEFAULT_PROFILE) 
                                      ?? Profiles.FirstOrDefault();
                    SwitchProfile(defaultProfile);
                    WriteDebugLog($"Switched to profile: {defaultProfile?.Name}");
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error deleting profile {profile?.Name}: {ex.Message}");
                MessageBox.Show($"Error deleting profile: {ex.Message}", "Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Switch to different profile
        /// </summary>
        public void SwitchProfile(Profile profile)
        {
            if (profile != null)
            {
                CurrentProfile = profile;
                WriteDebugLog($"Switched to profile: {profile.Name}");
                ProfileChanged?.Invoke(profile);
            }
        }

        /// <summary>
        /// Create new profile with given name
        /// </summary>
        public Profile CreateNewProfile(string name)
        {
            // Check if name already exists
            if (Profiles.Any(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show($"Profile '{name}' already exists.", "Warning", 
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return null;
            }

            var newProfile = new Profile
            {
                Name = name,
                Description = $"Profile created on {DateTime.Now:yyyy-MM-dd HH:mm}",
                Settings = new ProfileSettings() // Default settings
            };

            SaveProfile(newProfile);
            WriteDebugLog($"Created new profile: {name}");
            return newProfile;
        }

        /// <summary>
        /// Create default profile with standard settings
        /// </summary>
        private void CreateDefaultProfile()
        {
            var defaultProfile = new Profile
            {
                Name = DEFAULT_PROFILE,
                Description = "Default ExportPlus profile with standard settings",
                CreatedDate = DateTime.Now,
                Settings = new ProfileSettings
                {
                    PDFEnabled = false,  // Changed: Only DWG by default
                    PDFPrinterName = "PDF24",
                    PaperPlacementCenter = true,
                    FitToPage = false,
                    ZoomPercent = 100,
                    VectorProcessing = true,
                    ColorMode = "Color",
                    RasterQuality = "High",
                    CreateSeparateFiles = true,
                    HideCropBoundaries = true,
                    HideScopeBoxes = true,
                    OutputFolder = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop), 
                        "Export +"),
                    SaveAllInSameFolder = true,
                    ReportType = "Don't Save Report"
                }
            };

            SaveProfile(defaultProfile);
            CurrentProfile = defaultProfile;
            WriteDebugLog("Default profile created and set as current");
        }

        /// <summary>
        /// Write debug log
        /// </summary>
        private void WriteDebugLog(string message)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                string fullMessage = $"[ProfileManager] {timestamp} - {message}";
                Debug.WriteLine(fullMessage);
                OutputDebugStringA(fullMessage + "\r\n");
            }
            catch { }
        }
    }
}
