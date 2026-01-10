using System;
using Autodesk.PackageBuilder;

namespace Quoc_MEP.Builders
{
    /// <summary>
    /// Builder for creating Quoc_MEP.addin manifest file for Revit.
    /// </summary>
    /// <remarks>
    /// Generates Revit add-in manifest with proper GUID and metadata.
    /// Compatible with Revit 2020-2026 API.
    /// </remarks>
    public class QuocMEPAddinBuilder : RevitAddInsBuilder
    {
        private const string AddinName = "Quoc_MEP";
        private const string AddinAssembly = "Quoc_MEP.dll";
        private const string AddinClassName = "Quoc_MEP.Application";
        private const string AddinGuid = "AD8D446D-A204-48BF-A983-C931DD3C066F";
        private const string VendorId = "QNVN";
        private const string VendorDescription = "Quoc Nguyen - MEP Tools";
        
        /// <summary>
        /// Initializes a new instance of QuocMEPAddinBuilder.
        /// </summary>
        public QuocMEPAddinBuilder()
        {
            // Create Application AddIn entry with metadata
            AddIn.CreateEntry("Application")
                .Name(AddinName)
                .Assembly(AddinAssembly)
                .AddInId(new Guid(AddinGuid))
                .FullClassName(AddinClassName)
                .VendorId(VendorId)
                .VendorDescription(VendorDescription);
            
            // Note: ManifestSettings for Revit 2026+ are automatically handled
            // by the Revit.Sdk build process (PatchManifest task)
            // See: https://github.com/Nice3point/RevitTemplates/wiki
        }
    }
}
