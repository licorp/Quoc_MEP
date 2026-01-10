using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using System.Linq;
using Autodesk.Revit.DB;
using Quoc_MEP.Export.Models;

namespace Quoc_MEP.Export.Managers
{
    public class XMLProfileManager
    {
        private static string ProfilesFolder => 
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                         "ExportPlusAddin", "Profiles");

        private static string DiRootsProfilesFolder => 
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                         "DiRoots", "ExportPlus");

        static XMLProfileManager()
        {
            if (!Directory.Exists(ProfilesFolder))
                Directory.CreateDirectory(ProfilesFolder);
        }

        public static ExportPlusXMLProfile LoadProfileFromXML(string filePath)
        {
            try
            {
                WriteDebugLog($"Loading XML profile from: {filePath}");
                var serializer = new XmlSerializer(typeof(ExportPlusProfileList));
                using (var reader = new StreamReader(filePath))
                {
                    var profileList = (ExportPlusProfileList)serializer.Deserialize(reader);
                    var profile = profileList.Profiles.FirstOrDefault();
                    if (profile != null)
                    {
                        WriteDebugLog($"XML profile loaded: {profile.Name}");
                    }
                    return profile;
                }
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error loading XML profile: {ex.Message}");
                throw new Exception($"Error loading profile: {ex.Message}");
            }
        }

        public static void SaveProfileToXML(ExportPlusXMLProfile profile, string filePath)
        {
            try
            {
                WriteDebugLog($"Saving XML profile to: {filePath}");
                var profileList = new ExportPlusProfileList();
                profileList.Profiles.Add(profile);

                var serializer = new XmlSerializer(typeof(ExportPlusProfileList));
                using (var writer = new StreamWriter(filePath))
                {
                    serializer.Serialize(writer, profileList);
                }
                WriteDebugLog($"XML profile saved successfully");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"Error saving XML profile: {ex.Message}");
                throw new Exception($"Error saving profile: {ex.Message}");
            }
        }

        public static List<string> GetAvailableXMLProfiles()
        {
            var profiles = new List<string>();
            
            // Load from our folder
            if (Directory.Exists(ProfilesFolder))
            {
                var ourProfiles = Directory.GetFiles(ProfilesFolder, "*.xml")
                                           .Select(Path.GetFileNameWithoutExtension);
                profiles.AddRange(ourProfiles);
            }

            // Load from DiRoots ExportPlus folder
            if (Directory.Exists(DiRootsProfilesFolder))
            {
                var diRootsProfiles = Directory.GetFiles(DiRootsProfilesFolder, "*.xml")
                                              .Select(f => "DiRoots: " + Path.GetFileNameWithoutExtension(f));
                profiles.AddRange(diRootsProfiles);
            }

            WriteDebugLog($"Found {profiles.Count} XML profiles");
            return profiles;
        }

        public static List<SheetFileNameInfo> GenerateCustomFileNames(
            ExportPlusXMLProfile profile, 
            List<ViewSheet> sheets)
        {
            WriteDebugLog($"Generating custom file names for {sheets.Count} sheets");
            var result = new List<SheetFileNameInfo>();
            
            // Get selected parameters from SelectedParams_Virtual
            List<SelectionParameter> parameters = null;
            string separator = "-";
            
            if (profile.TemplateInfo.SelectionSheets?.SelectedParamsVirtual?.SelectionParameters != null)
            {
                parameters = profile.TemplateInfo.SelectionSheets.SelectedParamsVirtual.SelectionParameters
                    .Where(p => p.IsSelected)
                    .ToList();
                    
                separator = profile.TemplateInfo.SelectionSheets.FieldSeparator ?? "-";
                WriteDebugLog($"Found {parameters.Count} selected parameters, separator: '{separator}'");
            }
            
            if (parameters == null || !parameters.Any())
            {
                WriteDebugLog("No custom filename parameters found, using sheet numbers");
                // Fallback to sheet number only
                foreach (var sheet in sheets)
                {
                    result.Add(new SheetFileNameInfo
                    {
                        SheetId = sheet.Id.IntegerValue.ToString(),
                        SheetNumber = sheet.SheetNumber,
                        SheetName = sheet.Name,
                        Revision = GetSheetRevision(sheet),
                        Size = GetSheetPaperSize(sheet),
                        CustomFileName = sheet.SheetNumber,
                        IsSelected = true
                    });
                }
            }
            else
            {
                // Build custom filenames using selected parameters
                foreach (var sheet in sheets)
                {
                    var fileName = BuildCustomFileNameFromSelectionParams(sheet, parameters, separator);
                    
                    result.Add(new SheetFileNameInfo
                    {
                        SheetId = sheet.Id.IntegerValue.ToString(),
                        SheetNumber = sheet.SheetNumber,
                        SheetName = sheet.Name,
                        Revision = GetSheetRevision(sheet),
                        Size = GetSheetPaperSize(sheet),
                        CustomFileName = fileName,
                        IsSelected = true
                    });
                }
            }
            
            WriteDebugLog($"Generated {result.Count} custom file names");
            return result;
        }
        
        private static string BuildCustomFileNameFromSelectionParams(
            ViewSheet sheet, 
            List<SelectionParameter> parameters, 
            string separator)
        {
            var parts = new List<string>();
            
            foreach (var param in parameters)
            {
                string value = "";
                
                // Handle custom separators
                if (param.Type == "CustemSeparator")
                {
                    // Use the DisplayName as the separator
                    if (!string.IsNullOrEmpty(param.DisplayName))
                    {
                        parts.Add(param.DisplayName.Trim());
                    }
                    continue;
                }
                
                // Get parameter value based on DisplayName
                string paramName = param.DisplayName?.Trim() ?? "";
                
                switch (paramName)
                {
                    case "Sheet Number":
                        value = sheet.SheetNumber;
                        break;
                    case "Sheet Number Prefix":
                        // Extract prefix from sheet number (e.g., "A-101" -> "A")
                        var sheetNumber = sheet.SheetNumber;
                        var dashIndex = sheetNumber.IndexOf('-');
                        value = dashIndex > 0 ? sheetNumber.Substring(0, dashIndex) : "";
                        break;
                    case "Sheet Name":
                        value = sheet.Name;
                        break;
                    case "Current Revision":
                        value = GetSheetRevision(sheet);
                        break;
                    default:
                        // Try to get parameter from sheet
                        var sheetParam = sheet.LookupParameter(paramName);
                        if (sheetParam != null)
                        {
                            value = GetParameterValueAsString(sheetParam);
                        }
                        break;
                }
                
                if (!string.IsNullOrEmpty(value))
                {
                    parts.Add(value);
                }
            }
            
            // Join parts with separator, but don't add separator between custom separators
            var fileName = string.Join("", parts);
            
            // Fallback to sheet number if no custom name generated
            if (string.IsNullOrWhiteSpace(fileName))
            {
                fileName = sheet.SheetNumber;
            }
            
            return fileName;
        }
        
        private static string GetParameterValueAsString(Parameter param)
        {
            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        return param.AsString() ?? "";
                    case StorageType.Integer:
                        return param.AsInteger().ToString();
                    case StorageType.Double:
                        return param.AsDouble().ToString("F2");
                    case StorageType.ElementId:
                        return param.AsValueString() ?? "";
                    default:
                        return "";
                }
            }
            catch
            {
                return "";
            }
        }

        private static string GetParameterValue(ViewSheet sheet, string parameterName)
        {
            try
            {
                switch (parameterName)
                {
                    case "Sheet Number":
                        return sheet.SheetNumber;
                    case "Sheet Name":
                        return sheet.Name;
                    case "Current Revision":
                        return GetSheetRevision(sheet);
                    default:
                        var param = sheet.LookupParameter(parameterName);
                        if (param != null)
                        {
                            switch (param.StorageType)
                            {
                                case StorageType.String:
                                    return param.AsString() ?? "";
                                case StorageType.Integer:
                                    return param.AsInteger().ToString();
                                case StorageType.Double:
                                    return param.AsDouble().ToString();
                                default:
                                    return param.AsValueString() ?? "";
                            }
                        }
                        return "";
                }
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static string GetSheetRevision(ViewSheet sheet)
        {
            try
            {
                var revParam = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION);
                return revParam?.AsString() ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static string GetSheetPaperSize(ViewSheet sheet)
        {
            try
            {
                // Try to get paper size from title block
                var titleBlocks = new FilteredElementCollector(sheet.Document, sheet.Id)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .WhereElementIsNotElementType()
                    .ToElements();

                if (titleBlocks.Any())
                {
                    var titleBlock = titleBlocks.First();
                    var sizeParam = titleBlock.LookupParameter("Sheet Size");
                    if (sizeParam != null)
                    {
                        return sizeParam.AsString() ?? "A3";
                    }
                }

                // Fallback: detect from sheet dimensions
                var outline = sheet.Outline;
                var width = outline.Max.U - outline.Min.U;
                var height = outline.Max.V - outline.Min.V;
                
                // Convert to mm (assuming feet)
                var widthMm = width * 304.8;
                var heightMm = height * 304.8;
                
                // Common paper sizes in mm
                if (Math.Abs(widthMm - 420) < 50 && Math.Abs(heightMm - 297) < 50) return "A3";
                if (Math.Abs(widthMm - 297) < 50 && Math.Abs(heightMm - 210) < 50) return "A4";
                if (Math.Abs(widthMm - 594) < 50 && Math.Abs(heightMm - 420) < 50) return "A2";
                if (Math.Abs(widthMm - 841) < 50 && Math.Abs(heightMm - 594) < 50) return "A1";
                if (Math.Abs(widthMm - 1189) < 50 && Math.Abs(heightMm - 841) < 50) return "A0";
                
                return "A3"; // Default
            }
            catch (Exception)
            {
                return "A3";
            }
        }

        public static ExportPlusProfile ConvertXMLToProfile(ExportPlusXMLProfile xmlProfile)
        {
            WriteDebugLog($"Converting XML profile to standard profile: {xmlProfile.Name}");
            var profile = new ExportPlusProfile
            {
                ProfileName = xmlProfile.Name,
                OutputFolder = xmlProfile.FilePath ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                CreateSeparateFolders = xmlProfile.TemplateInfo.IsSeparateFile,
                HideCropRegions = xmlProfile.TemplateInfo.HideCropBoundaries,
                HideScopeBoxes = xmlProfile.TemplateInfo.HideScopeBox,
                PaperSize = xmlProfile.TemplateInfo.PaperSize,
                SelectedFormats = new List<string>()
            };

            // Convert format flags
            if (xmlProfile.TemplateInfo.IsPDFChecked) profile.SelectedFormats.Add("PDF");
            if (xmlProfile.TemplateInfo.IsDWGChecked) profile.SelectedFormats.Add("DWG");
            if (xmlProfile.TemplateInfo.IsDGNChecked) profile.SelectedFormats.Add("DGN");
            if (xmlProfile.TemplateInfo.IsIFCChecked) profile.SelectedFormats.Add("IFC");
            if (xmlProfile.TemplateInfo.IsIMGChecked) profile.SelectedFormats.Add("JPG");
            if (xmlProfile.TemplateInfo.IsNWCChecked) profile.SelectedFormats.Add("NWC");
            if (xmlProfile.TemplateInfo.IsDWFChecked) profile.SelectedFormats.Add("DWF");

            WriteDebugLog($"Converted profile with {profile.SelectedFormats.Count} formats");
            return profile;
        }
        
        /// <summary>
        /// Apply XML profile settings to UI/ViewModel
        /// This method maps all XML settings to the current UI state
        /// </summary>
        public static void ApplyXMLProfileToUI(ExportPlusXMLProfile xmlProfile, 
            Action<string, object> setUIProperty)
        {
            if (xmlProfile == null || setUIProperty == null)
            {
                WriteDebugLog("ERROR: Cannot apply XML profile - null parameters");
                return;
            }

            WriteDebugLog($"=== APPLYING XML PROFILE TO UI: {xmlProfile.Name} ===");
            var template = xmlProfile.TemplateInfo;

            try
            {
                // PDF/Export Format Settings
                WriteDebugLog("Applying PDF/Export format settings...");
                setUIProperty("IsVectorProcessing", template.IsVectorProcessing);
                setUIProperty("RasterQuality", template.RasterQuality);
                setUIProperty("ColorMode", template.Color);
                setUIProperty("IsFitToPage", template.IsFitToPage);
                
                // Paper Placement Settings
                WriteDebugLog("Applying paper placement settings...");
                setUIProperty("IsCenter", template.IsCenter);
                setUIProperty("SelectedMarginType", template.SelectedMarginType);
                setUIProperty("PaperSize", template.PaperSize);
                
                // View Options
                WriteDebugLog("Applying view options...");
                setUIProperty("ViewLinksInBlue", template.ViewLink);
                setUIProperty("HideRefWorkPlanes", template.HidePlanes);
                setUIProperty("HideScopeBoxes", template.HideScopeBox);
                setUIProperty("HideUnreferencedViewTags", template.HideUnreferencedTags);
                setUIProperty("HideCropBoundaries", template.HideCropBoundaries);
                setUIProperty("ReplaceHalftone", template.ReplaceHalftone);
                setUIProperty("MaskCoincidentLines", template.MaskCoincidentLines);
                
                // File Settings
                WriteDebugLog("Applying file settings...");
                setUIProperty("CreateSeparateFiles", template.IsSeparateFile);
                setUIProperty("OutputFolder", template.FilePath ?? "");
                
                // DWF Settings
                if (template.DWF != null)
                {
                    WriteDebugLog("Applying DWF settings...");
                    setUIProperty("DWF_ImageFormat", template.DWF.OptImageFormat);
                    setUIProperty("DWF_ImageQuality", template.DWF.OptImageQuality);
                    setUIProperty("DWF_ExportTextures", template.DWF.OptExportTextures);
                }
                
                // NWC Settings
                if (template.NWC != null)
                {
                    WriteDebugLog("Applying NWC settings...");
                    // Core settings
                    setUIProperty("NWC_ConvertConstructionParts", template.NWC.ConvertConstructionParts);
                    setUIProperty("NWC_ConvertElementIds", template.NWC.ConvertElementIds);
                    setUIProperty("NWC_ConvertElementParameters", template.NWC.ConvertElementParameters);
                    setUIProperty("NWC_ConvertElementProperties", template.NWC.ConvertElementProperties);
                    setUIProperty("NWC_ConvertLinkedFiles", template.NWC.ConvertLinkedFiles);
                    setUIProperty("NWC_ConvertRoomAsAttribute", template.NWC.ConvertRoomAsAttribute);
                    setUIProperty("NWC_ConvertURLs", template.NWC.ConvertURLs);
                    setUIProperty("NWC_Coordinates", template.NWC.Coordinates);
                    setUIProperty("NWC_DivideFileIntoLevels", template.NWC.DivideFileIntoLevels);
                    
                    // New settings (added 2024)
                    setUIProperty("NWC_EmbedTextures", template.NWC.EmbedTextures);
                    setUIProperty("NWC_ExportScope", template.NWC.ExportScope);
                    setUIProperty("NWC_SeparateCustomProperties", template.NWC.SeparateCustomProperties);
                    setUIProperty("NWC_StrictSectioning", template.NWC.StrictSectioning);
                    setUIProperty("NWC_TypePropertiesOnElements", template.NWC.TypePropertiesOnElements);
                    
                    // Additional settings
                    setUIProperty("NWC_ExportRoomGeometry", template.NWC.ExportRoomGeometry);
                    setUIProperty("NWC_TryAndFindMissingMaterials", template.NWC.TryAndFindMissingMaterials);
                    setUIProperty("NWC_ConvertLinkedCADFormats", template.NWC.ConvertLinkedCADFormats);
                    setUIProperty("NWC_ConvertLights", template.NWC.ConvertLights);
                    setUIProperty("NWC_FacetingFactor", template.NWC.FacetingFactor);
                }
                
                // IFC Settings
                if (template.IFC != null)
                {
                    WriteDebugLog("Applying IFC settings...");
                    setUIProperty("IFC_FileVersion", template.IFC.FileVersion);
                    setUIProperty("IFC_SpaceBoundaries", template.IFC.SpaceBoundaries);
                    setUIProperty("IFC_SitePlacement", template.IFC.SitePlacement);
                    setUIProperty("IFC_ExportBaseQuantities", template.IFC.ExportBaseQuantities);
                    setUIProperty("IFC_ExportIFCCommonPropertySets", template.IFC.ExportIFCCommonPropertySets);
                    setUIProperty("IFC_TessellationLevelOfDetail", template.IFC.TessellationLevelOfDetail);
                    setUIProperty("IFC_VisibleElementsOfCurrentView", template.IFC.VisibleElementsOfCurrentView);
                }
                
                // IMG Settings
                if (template.IMG != null)
                {
                    WriteDebugLog("Applying IMG settings...");
                    setUIProperty("IMG_ImageResolution", template.IMG.ImageResolution);
                    setUIProperty("IMG_FileType", template.IMG.HLRandWFViewsFileType);
                    setUIProperty("IMG_ZoomType", template.IMG.ZoomType);
                    setUIProperty("IMG_PixelSize", template.IMG.PixelSize);
                }
                
                // Format Checkboxes
                WriteDebugLog("Applying format checkboxes...");
                setUIProperty("IsPDFChecked", template.IsPDFChecked);
                setUIProperty("IsDWGChecked", template.IsDWGChecked);
                setUIProperty("IsDGNChecked", template.IsDGNChecked);
                setUIProperty("IsIFCChecked", template.IsIFCChecked);
                setUIProperty("IsIMGChecked", template.IsIMGChecked);
                setUIProperty("IsNWCChecked", template.IsNWCChecked);
                setUIProperty("IsDWFChecked", template.IsDWFChecked);
                
                // Custom File Name Parameters
                if (template.SelectionSheets?.SelectedParamsVirtual?.SelectionParameters != null)
                {
                    WriteDebugLog($"Applying custom filename parameters ({template.SelectionSheets.SelectedParamsVirtual.SelectionParameters.Count} params)...");
                    var selectedParams = template.SelectionSheets.SelectedParamsVirtual.SelectionParameters
                        .Where(p => p.IsSelected)
                        .Select(p => p.DisplayName?.Trim())
                        .Where(n => !string.IsNullOrEmpty(n))
                        .ToList();
                    
                    setUIProperty("CustomFileNameParameters", selectedParams);
                    WriteDebugLog($"Selected parameters: {string.Join(", ", selectedParams)}");
                }
                
                WriteDebugLog($"=== XML PROFILE APPLIED SUCCESSFULLY: {xmlProfile.Name} ===");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"ERROR applying XML profile: {ex.Message}");
                WriteDebugLog($"Stack trace: {ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// Get detailed export settings from XML profile for specific format
        /// </summary>
        public static Dictionary<string, object> GetFormatSettings(ExportPlusXMLProfile xmlProfile, string format)
        {
            var settings = new Dictionary<string, object>();
            
            if (xmlProfile?.TemplateInfo == null) return settings;
            
            var template = xmlProfile.TemplateInfo;
            
            switch (format.ToUpper())
            {
                case "PDF":
                    settings["VectorProcessing"] = template.IsVectorProcessing;
                    settings["RasterQuality"] = template.RasterQuality;
                    settings["ColorMode"] = template.Color;
                    settings["FitToPage"] = template.IsFitToPage;
                    settings["IsCenter"] = template.IsCenter;
                    settings["MarginType"] = template.SelectedMarginType;
                    break;
                    
                case "DWF":
                    if (template.DWF != null)
                    {
                        settings["IsDwfx"] = template.DWF.IsDwfx;
                        settings["ImageFormat"] = template.DWF.OptImageFormat;
                        settings["ImageQuality"] = template.DWF.OptImageQuality;
                        settings["ExportTextures"] = template.DWF.OptExportTextures;
                        settings["FitToPage"] = template.DWF.IsFitToPage;
                        settings["RasterQuality"] = template.DWF.RasterQuality;
                    }
                    break;
                    
                case "NWC":
                    if (template.NWC != null)
                    {
                        settings["ConvertElementProperties"] = template.NWC.ConvertElementProperties;
                        settings["Coordinates"] = template.NWC.Coordinates;
                        settings["DivideFileIntoLevels"] = template.NWC.DivideFileIntoLevels;
                        settings["ExportElementIds"] = template.NWC.ExportElementIds;
                        settings["ExportParts"] = template.NWC.ExportParts;
                        settings["ExportRoomAsAttribute"] = template.NWC.ExportRoomAsAttribute;
                        settings["FacetingFactor"] = template.NWC.FacetingFactor;
                    }
                    break;
                    
                case "IFC":
                    if (template.IFC != null)
                    {
                        settings["FileVersion"] = template.IFC.FileVersion;
                        settings["SpaceBoundaries"] = template.IFC.SpaceBoundaries;
                        settings["SitePlacement"] = template.IFC.SitePlacement;
                        settings["ExportBaseQuantities"] = template.IFC.ExportBaseQuantities;
                        settings["ExportIFCCommonPropertySets"] = template.IFC.ExportIFCCommonPropertySets;
                        settings["TessellationLevelOfDetail"] = template.IFC.TessellationLevelOfDetail;
                        settings["VisibleElementsOfCurrentView"] = template.IFC.VisibleElementsOfCurrentView;
                    }
                    break;
                    
                case "IMG":
                case "JPG":
                case "PNG":
                    if (template.IMG != null)
                    {
                        settings["ImageResolution"] = template.IMG.ImageResolution;
                        settings["FileType"] = template.IMG.HLRandWFViewsFileType;
                        settings["ZoomType"] = template.IMG.ZoomType;
                        settings["PixelSize"] = template.IMG.PixelSize;
                    }
                    break;
            }
            
            WriteDebugLog($"Retrieved {settings.Count} settings for format: {format}");
            return settings;
        }

        private static void WriteDebugLog(string message)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                string fullMessage = $"[XMLProfileManager] {timestamp} - {message}";
                
                System.Diagnostics.Debug.WriteLine(fullMessage);
                Console.WriteLine(fullMessage);
                
                // Output for DebugView
                OutputDebugStringA(fullMessage + "\r\n");
            }
            catch (Exception)
            {
                // Ignore logging errors
            }
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Ansi)]
        private static extern void OutputDebugStringA(string lpOutputString);
    }
}