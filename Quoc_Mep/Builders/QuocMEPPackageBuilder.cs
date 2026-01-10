using Autodesk.PackageBuilder;

namespace Quoc_MEP.Builders
{
    /// <summary>
    /// Builder for creating PackageContents.xml for Quoc_MEP Universal Package.
    /// Supports Revit versions 2020-2026.
    /// </summary>
    /// <remarks>
    /// Uses Autodesk.PackageBuilder fluent API for type-safe manifest generation.
    /// See: https://github.com/ricaun-io/Autodesk.PackageBuilder
    /// </remarks>
    public class QuocMEPPackageBuilder : PackageContentsBuilder
    {
        private const string PackageName = "Quoc_MEP";
        private const string CompanyName = "Quoc Nguyen";
        private const string CompanyEmail = "quoc.nguyen@company.com";
        private const string CompanyUrl = "https://github.com/licorp/Quoc_MEP";
        private const string VendorId = "QNVN";
        private const int MinRevitVersion = 2020;
        private const int MaxRevitVersion = 2026;
        
        /// <summary>
        /// Initializes a new instance of QuocMEPPackageBuilder.
        /// </summary>
        /// <param name="version">Package version (e.g., "1.0.0")</param>
        public QuocMEPPackageBuilder(string version = "1.0.0")
        {
            // Use RevitApplication() extension for cleaner code (Nice3point pattern)
            ApplicationPackage
                .Create()
                .RevitApplication()
                .Name(PackageName)
                .AppVersion(version)
                .Description($"{PackageName} - MEP Tools for Revit {MinRevitVersion}-{MaxRevitVersion}");

            // Company Details
            CompanyDetails
                .Create(CompanyName)
                .Email(CompanyEmail)
                .Url(CompanyUrl);

            // Generate Components for each supported Revit version
            for (int year = MinRevitVersion; year <= MaxRevitVersion; year++)
            {
                CreateRevitComponent(year);
            }
        }

        /// <summary>
        /// Creates a component entry for a specific Revit version.
        /// </summary>
        /// <param name="year">Revit version year (e.g., 2024)</param>
        private void CreateRevitComponent(int year)
        {
            Components
                .CreateEntry($"Revit {year}")
                .RevitPlatform(year)  // Uses extension method for automatic SeriesMin/SeriesMax
                .AppName(PackageName)
                .ModuleName($"./Contents/{year}/{PackageName}.addin");
        }
    }
}
