using System;
using Autodesk.PackageBuilder;

namespace Quoc_MEP.Builders
{
    /// <summary>
    /// Builder for creating Quoc_MEP.addin file for Revit
    /// </summary>
    public class QuocMEPAddinBuilder : RevitAddInsBuilder
    {
        private const string AddinName = "Quoc_MEP";
        private const string VendorId = "QNVN";
        private const string VendorDescription = "Quoc Nguyen - MEP Tools";
        
        public QuocMEPAddinBuilder()
        {
            // Create Application AddIn entry
            AddIn.CreateEntry("Application")
                .Name(AddinName)
                .Assembly($"{AddinName}.dll")
                .AddInId(new Guid("AD8D446D-A204-48BF-A983-C931DD3C066F"))
                .FullClassName($"{AddinName}.Application")
                .VendorId(VendorId)
                .VendorDescription(VendorDescription);
        }
    }
}
