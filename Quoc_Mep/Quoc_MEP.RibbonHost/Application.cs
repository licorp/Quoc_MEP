using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace Quoc_MEP.RibbonHost
{
    /// <summary>
    /// Shared ribbon host that loads commands from multiple add-in DLLs.
    /// Creates a single ribbon panel and dynamically discovers and adds buttons.
    /// </summary>
    public class Application : IExternalApplication
    {
        private const string TAB_NAME = "Quoc MEP";
        private const string PANEL_NAME = "MEP Tools";

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create ribbon tab (ignore if exists)
                try
                {
                    application.CreateRibbonTab(TAB_NAME);
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    // Tab already exists - continue
                }

                // Create ribbon panel
                RibbonPanel panel = application.CreateRibbonPanel(TAB_NAME, PANEL_NAME);

                // Load commands from DLLs in same folder
                LoadCommands(panel);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RibbonHost Error", $"Failed to initialize ribbon:\n{ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        /// <summary>
        /// Discovers and loads commands from DLLs in the same folder as this assembly.
        /// </summary>
        private void LoadCommands(RibbonPanel panel)
        {
            try
            {
                // Get the directory where this DLL is located
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string directory = Path.GetDirectoryName(assemblyPath);

                // Find all DLLs except this one
                var dllFiles = Directory.GetFiles(directory, "*.dll")
                    .Where(f => !f.Equals(assemblyPath, StringComparison.OrdinalIgnoreCase))
                    .Where(f => !Path.GetFileName(f).StartsWith("Autodesk."))
                    .Where(f => !Path.GetFileName(f).StartsWith("RevitAPI"))
                    .Where(f => !Path.GetFileName(f).StartsWith("System."))
                    .Where(f => !Path.GetFileName(f).StartsWith("Microsoft."))
                    .ToList();

                // Load commands from each DLL
                foreach (var dllPath in dllFiles)
                {
                    LoadCommandsFromDll(panel, dllPath);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RibbonHost Warning", $"Error loading commands:\n{ex.Message}");
            }
        }

        /// <summary>
        /// Loads commands from a specific DLL using CommandInfo classes or attributes.
        /// </summary>
        private void LoadCommandsFromDll(RibbonPanel panel, string dllPath)
        {
            try
            {
                Assembly assembly = Assembly.LoadFrom(dllPath);
                string dllName = Path.GetFileNameWithoutExtension(dllPath);

                // Look for CommandInfo classes (naming pattern: *CommandInfo)
                var commandInfoTypes = assembly.GetTypes()
                    .Where(t => t.Name.EndsWith("CommandInfo") && !t.IsAbstract && t.IsClass)
                    .ToList();

                foreach (var infoType in commandInfoTypes)
                {
                    try
                    {
                        // Try to get static properties for command metadata
                        var nameProperty = infoType.GetProperty("Name", BindingFlags.Public | BindingFlags.Static);
                        var textProperty = infoType.GetProperty("Text", BindingFlags.Public | BindingFlags.Static);
                        var tooltipProperty = infoType.GetProperty("Tooltip", BindingFlags.Public | BindingFlags.Static);
                        var commandClassProperty = infoType.GetProperty("CommandClass", BindingFlags.Public | BindingFlags.Static);

                        if (nameProperty != null && textProperty != null && commandClassProperty != null)
                        {
                            string commandName = nameProperty.GetValue(null) as string;
                            string buttonText = textProperty.GetValue(null) as string;
                            string tooltip = tooltipProperty?.GetValue(null) as string ?? buttonText;
                            string commandClass = commandClassProperty.GetValue(null) as string;

                            AddCommandButton(panel, commandName, buttonText, tooltip, dllPath, commandClass);
                        }
                    }
                    catch
                    {
                        // Skip this command info if metadata extraction fails
                        continue;
                    }
                }

                // Fallback: If no CommandInfo found, look for IExternalCommand implementations
                if (!commandInfoTypes.Any())
                {
                    var commandTypes = assembly.GetTypes()
                        .Where(t => typeof(IExternalCommand).IsAssignableFrom(t) && !t.IsAbstract && t.IsClass)
                        .ToList();

                    foreach (var cmdType in commandTypes)
                    {
                        string commandName = $"{dllName}.{cmdType.Name}";
                        string buttonText = cmdType.Name.Replace("Command", "").Replace("Cmd", "");
                        AddCommandButton(panel, commandName, buttonText, buttonText, dllPath, cmdType.FullName);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but continue loading other DLLs
                System.Diagnostics.Debug.WriteLine($"Failed to load commands from {dllPath}: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds a button to the ribbon panel.
        /// </summary>
        private void AddCommandButton(RibbonPanel panel, string commandName, string buttonText, string tooltip, string assemblyPath, string className)
        {
            try
            {
                var buttonData = new PushButtonData(
                    commandName,
                    buttonText,
                    assemblyPath,
                    className)
                {
                    ToolTip = tooltip
                };

                // Try to get icon from embedded resource
                try
                {
                    var assembly = Assembly.LoadFrom(assemblyPath);
                    var iconResourceName = $"{assembly.GetName().Name}.Resources.Icon.png";
                    var resourceStream = assembly.GetManifestResourceStream(iconResourceName);
                    
                    if (resourceStream != null)
                    {
                        BitmapImage bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.StreamSource = resourceStream;
                        bitmap.EndInit();
                        buttonData.LargeImage = bitmap;
                    }
                    else
                    {
                        // Use default icon if no embedded resource
                        buttonData.LargeImage = GetDefaultIcon();
                    }
                }
                catch
                {
                    // Use default icon if loading fails
                    buttonData.LargeImage = GetDefaultIcon();
                }

                panel.AddItem(buttonData);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to add button {commandName}: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a default icon for buttons without custom icons.
        /// </summary>
        private ImageSource GetDefaultIcon()
        {
            int size = 32;
            var bitmap = new System.Drawing.Bitmap(size, size);
            using (var g = System.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(System.Drawing.Color.LightBlue);
                g.DrawRectangle(System.Drawing.Pens.Navy, 0, 0, size - 1, size - 1);
            }

            var bitmapImage = new BitmapImage();
            using (var memory = new System.IO.MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                memory.Position = 0;
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
            }

            return bitmapImage;
        }
    }
}
