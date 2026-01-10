# Script cài đặt Nice3point.Revit.Api packages cho nhiều phiên bản Revit
# Chạy: .\install-nice3point-packages.ps1 -RevitVersion 2023

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("2020", "2021", "2022", "2023", "2024", "2025", "2026")]
    [string]$RevitVersion = "2023",
    
    [Parameter(Mandatory=$false)]
    [switch]$AllVersions = $false
)

$ErrorActionPreference = "Stop"
$projectPath = $PSScriptRoot
$packagesPath = Join-Path $projectPath "packages"

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Nice3point.Revit.Api Package Installer" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Các packages cần cài
$packages = @(
    "Nice3point.Revit.Api.RevitAPI",
    "Nice3point.Revit.Api.RevitAPIUI",
    "Nice3point.Revit.Api.AdWindows"
)

function Install-RevitPackages {
    param([string]$Version)
    
    Write-Host "📦 Installing packages for Revit $Version..." -ForegroundColor Green
    Write-Host ""
    
    foreach ($package in $packages) {
        Write-Host "  → Installing $package..." -ForegroundColor Yellow
        
        try {
            # Tìm version mới nhất cho Revit version đó
            $nugetCmd = ".\nuget.exe"
            
            # Install package
            & $nugetCmd install $package `
                -OutputDirectory $packagesPath `
                -NonInteractive `
                -PreRelease
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "    ✓ $package installed successfully" -ForegroundColor Green
            } else {
                Write-Host "    ✗ Failed to install $package" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "    ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        Write-Host ""
    }
}

# Cài cho tất cả versions
if ($AllVersions) {
    $allVersions = @("2020", "2021", "2022", "2023", "2024", "2025", "2026")
    
    foreach ($ver in $allVersions) {
        Install-RevitPackages -Version $ver
    }
}
else {
    # Cài cho version cụ thể
    Install-RevitPackages -Version $RevitVersion
}

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Update your .csproj file to use PackageReference" -ForegroundColor White
Write-Host "2. Remove old hardcoded Revit DLL references" -ForegroundColor White
Write-Host "3. Rebuild your project" -ForegroundColor White
Write-Host ""
Write-Host "See NICE3POINT_SETUP_GUIDE.md for detailed instructions" -ForegroundColor Cyan
