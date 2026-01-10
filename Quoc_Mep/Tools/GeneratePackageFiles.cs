using System;
using Autodesk.PackageBuilder;
using Quoc_MEP.Builders;

namespace Quoc_MEP.Tools
{
    /// <summary>
    /// Console tool to generate package files
    /// Usage: dotnet run --project GeneratePackageFiles.csproj -- [output-dir] [version]
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            string outputDir = args.Length > 0 ? args[0] : "./PackageFiles";
            string version = args.Length > 1 ? args[1] : "1.0.0";
            
            Console.WriteLine("================================================");
            Console.WriteLine("  Quoc MEP Package Generator");
            Console.WriteLine("  Using Autodesk.PackageBuilder");
            Console.WriteLine("================================================");
            Console.WriteLine();
            
            try
            {
                // Ensure output directory exists
                if (!System.IO.Directory.Exists(outputDir))
                {
                    System.IO.Directory.CreateDirectory(outputDir);
                    Console.WriteLine($"📁 Created output directory: {outputDir}");
                }
                
                Console.WriteLine($"📋 Configuration:");
                Console.WriteLine($"   Output Dir  : {outputDir}");
                Console.WriteLine($"   Version     : {version}");
                Console.WriteLine();
                
                // Generate PackageContents.xml
                Console.WriteLine("🎨 Generating PackageContents.xml...");
                var packageBuilder = new QuocMEPPackageBuilder(version);
                var packageXmlPath = System.IO.Path.Combine(outputDir, "PackageContents.xml");
                packageBuilder.Build(packageXmlPath);
                
                if (System.IO.File.Exists(packageXmlPath))
                {
                    Console.WriteLine($"✅ PackageContents.xml created: {packageXmlPath}");
                    Console.WriteLine();
                    Console.WriteLine("📄 Preview:");
                    Console.WriteLine(System.IO.File.ReadAllText(packageXmlPath));
                }
                
                Console.WriteLine();
                
                // Generate Quoc_MEP.addin
                Console.WriteLine("🎨 Generating Quoc_MEP.addin...");
                var addinBuilder = new QuocMEPAddinBuilder();
                var addinPath = System.IO.Path.Combine(outputDir, "Quoc_MEP.addin");
                addinBuilder.Build(addinPath);
                
                if (System.IO.File.Exists(addinPath))
                {
                    Console.WriteLine($"✅ Quoc_MEP.addin created: {addinPath}");
                    Console.WriteLine();
                    Console.WriteLine("📄 Preview:");
                    Console.WriteLine(System.IO.File.ReadAllText(addinPath));
                }
                
                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine("  ✅ All files generated successfully!");
                Console.WriteLine("================================================");
                Console.WriteLine();
                Console.WriteLine("Generated files:");
                Console.WriteLine($"  - {packageXmlPath}");
                Console.WriteLine($"  - {addinPath}");
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                Environment.Exit(1);
            }
        }
    }
}
