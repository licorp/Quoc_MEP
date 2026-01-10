using Autodesk.PackageBuilder;

namespace Quoc_MEP.Builders
{
    /// <summary>
    /// Builder for creating PackageContents.xml for Quoc_MEP Universal Package
    /// Supports Revit 2020-2026
    /// </summary>
    public class QuocMEPPackageBuilder : PackageContentsBuilder
    {
        private const string PackageName = "Quoc_MEP";
        private const string CompanyName = "Quoc Nguyen";
        private const string VendorId = "QNVN";
        
        public QuocMEPPackageBuilder(string version = "1.0.0")
        {
            // Application Package metadata
            ApplicationPackage
                .Create()
                .ProductType(ProductTypes.Application)
                .AutodeskProduct(AutodeskProducts.Revit)
                .Name(PackageName)
                .AppVersion(version)
                .Description("Quoc MEP Tools - Universal package for Revit 2020-2026");

            // Company Details
            CompanyDetails
                .Create(CompanyName)
                .Email("quoc.nguyen@company.com")
                .Url("https://github.com/licorp/Quoc_MEP");

            // Generate Components for each Revit version (2020-2026)
            CreateRevitComponent(2020);
            CreateRevitComponent(2021);
            CreateRevitComponent(2022);
            CreateRevitComponent(2023);
            CreateRevitComponent(2024);
            CreateRevitComponent(2025);
            CreateRevitComponent(2026);
        }

        private void CreateRevitComponent(int year)
        {
            Components
                .CreateEntry($"Revit {year}")
                .RevitPlatform(year)
                .AppName(PackageName)
                .ModuleName($"./Contents/{year}/Quoc_MEP.addin");
        }
    }
}
