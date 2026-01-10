# Script build project cho nhieu phien ban Revit
# Chay: .\build-all-versions.ps1

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("2020", "2021", "2022", "2023", "2024", "2025", "2026")]
    [string[]]$Versions = @("2020", "2021", "2022", "2023", "2024", "2025", "2026"),
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    
    [Parameter(Mandatory=$false)]
    [switch]$CleanBefore = $false
)

$ErrorActionPreference = "Continue"
$projectFile = "Quoc_MEP.csproj"
$successCount = 0
$failCount = 0
$results = @()

# Find MSBuild
$msbuildPath = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe | Select-Object -First 1
if (-not $msbuildPath) {
    Write-Host "ERROR: MSBuild not found!" -ForegroundColor Red
    Write-Host "Please install Visual Studio or Build Tools" -ForegroundColor Red
    exit 1
}
Write-Host "Using MSBuild: $msbuildPath" -ForegroundColor Gray
Write-Host ""

# Framework mapping
$frameworkInfo = @{
    "2020" = @{ Framework = ".NET 4.8"; UsePackages = $true }
    "2021" = @{ Framework = ".NET 4.8"; UsePackages = $true }
    "2022" = @{ Framework = ".NET 4.8"; UsePackages = $true }
    "2023" = @{ Framework = ".NET 4.8"; UsePackages = $true }
    "2024" = @{ Framework = ".NET 4.8"; UsePackages = $true }
    "2025" = @{ Framework = ".NET 8.0"; UsePackages = $false }
    "2026" = @{ Framework = ".NET 8.0"; UsePackages = $false }
}

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Multi-Version Revit Build Script" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration: $Configuration" -ForegroundColor Yellow
Write-Host "Versions to build: $($Versions -join ', ')" -ForegroundColor Yellow
Write-Host ""

foreach ($version in $Versions) {
    $info = $frameworkInfo[$version]
    
    Write-Host "--------------------------------------------------" -ForegroundColor Gray
    Write-Host "Building for Revit $version ($($info.Framework))..." -ForegroundColor Green
    
    # Check if Revit 2025/2026 requires .NET 8.0 SDK (optional warning)
    if ($info.Framework -eq ".NET 8.0") {
        Write-Host "  Note: Using hardcoded Revit DLL paths for 2025/2026" -ForegroundColor Yellow
        Write-Host "        (.NET 8.0 packages not compatible with .NET 4.8 project)" -ForegroundColor Yellow
        
        # Check if Revit installation exists
        $revitPath = "C:\Program Files\Autodesk\Revit $version"
        if (-not (Test-Path $revitPath)) {
            Write-Host "  WARNING: Revit $version not found at $revitPath" -ForegroundColor Red
            Write-Host "           Skipping build..." -ForegroundColor Red
            
            $results += [PSCustomObject]@{
                Version = $version
                Framework = $info.Framework
                Status = "Skipped (Revit not installed)"
                Duration = "0s"
            }
            $failCount++
            Write-Host ""
            continue
        }
    }
    
    Write-Host ""
    
    # Clean if requested
    if ($CleanBefore) {
        Write-Host "  Cleaning..." -ForegroundColor Yellow
        & $msbuildPath $projectFile `
            /t:Clean `
            /p:RevitVersion=$version `
            /p:Configuration=$Configuration `
            /nologo `
            /verbosity:minimal
    }
    
    # Build
    Write-Host "  Building..." -ForegroundColor Yellow
    $buildStart = Get-Date
    
    & $msbuildPath $projectFile `
        /t:Rebuild `
        /p:RevitVersion=$version `
        /p:Configuration=$Configuration `
        /nologo `
        /verbosity:minimal `
        /maxcpucount
    
    $buildEnd = Get-Date
    $duration = ($buildEnd - $buildStart).TotalSeconds
    
    if ($LASTEXITCODE -eq 0) {
        $successCount++
        Write-Host ""
        Write-Host "  Success - Revit $version build successful! ($([math]::Round($duration, 2))s)" -ForegroundColor Green
        
        $outputPath = "bin\$Configuration\Revit$version"
        if (Test-Path $outputPath) {
            Write-Host "  Output: $outputPath" -ForegroundColor Gray
        }
        
        $results += [PSCustomObject]@{
            Version = $version
            Framework = $info.Framework
            Status = "Success"
            Duration = "$([math]::Round($duration, 2))s"
        }
    }
    else {
        $failCount++
        Write-Host ""
        Write-Host "  Failed - Revit $version build failed!" -ForegroundColor Red
        
        $results += [PSCustomObject]@{
            Version = $version
            Framework = $info.Framework
            Status = "Failed"
            Duration = "$([math]::Round($duration, 2))s"
        }
    }
    
    Write-Host ""
}

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Build Summary" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Display results table
$results | Format-Table -AutoSize

Write-Host ""
Write-Host "Total: $($Versions.Count) builds" -ForegroundColor White
Write-Host "Success: $successCount" -ForegroundColor Green
Write-Host "Failed: $failCount" -ForegroundColor Red
Write-Host ""

if ($failCount -eq 0) {
    Write-Host "All builds completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Output locations:" -ForegroundColor Cyan
    Write-Host "  - Revit 2020-2024: bin\$Configuration\Revit{Version}\" -ForegroundColor White
    Write-Host "  - Revit 2025-2026: bin\$Configuration\Revit{Version}\" -ForegroundColor White
    Write-Host ""
    Write-Host "Note: Revit 2025-2026 use hardcoded DLL references" -ForegroundColor Yellow
    Write-Host "      For full .NET 8.0 support, create separate project" -ForegroundColor Yellow
    exit 0
}
else {
    Write-Host "Some builds failed. Check the output above for details." -ForegroundColor Yellow
    exit 1
}
