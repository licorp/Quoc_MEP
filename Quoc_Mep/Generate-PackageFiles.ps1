#!/usr/bin/env pwsh
#Requires -Version 7.0

<#
.SYNOPSIS
    Generates PackageContents.xml and .addin files using Autodesk.PackageBuilder
    
.DESCRIPTION
    This script uses the QuocMEPPackageBuilder and QuocMEPAddinBuilder classes
    to generate the required XML and .addin files for the Revit add-in package.
    
.PARAMETER OutputDir
    The directory where generated files will be saved
    
.PARAMETER Version
    The version string for the package (default: 1.0.0)
    
.EXAMPLE
    .\Generate-PackageFiles.ps1 -OutputDir ".\artifacts" -Version "1.2.3"
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputDir = ".\PackageFiles",
    
    [Parameter(Mandatory=$false)]
    [string]$Version = "1.0.0",
    
    [Parameter(Mandatory=$false)]
    [string]$ProjectPath = "Quoc_MEP.csproj"
)

$ErrorActionPreference = "Stop"

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Quoc MEP Package Generator" -ForegroundColor Cyan
Write-Host "  Using Autodesk.PackageBuilder" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Check if output directory exists, create if not
if (-not (Test-Path $OutputDir)) {
    Write-Host "📁 Creating output directory: $OutputDir" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# Absolute paths
$OutputDir = (Resolve-Path $OutputDir -ErrorAction SilentlyContinue)
if (-not $OutputDir) {
    $OutputDir = Join-Path (Get-Location) $OutputDir
    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }
    $OutputDir = (Resolve-Path $OutputDir).Path
} else {
    $OutputDir = $OutputDir.Path
}

# Get project path relative to script location
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectPath = Join-Path $ScriptDir $ProjectPath
if (-not (Test-Path $ProjectPath)) {
    Write-Host "❌ Project file not found: $ProjectPath" -ForegroundColor Red
    exit 1
}
$ProjectPath = (Resolve-Path $ProjectPath).Path
$ProjectDir = Split-Path $ProjectPath -Parent

Write-Host "📋 Configuration:" -ForegroundColor Green
Write-Host "   Output Dir  : $OutputDir" -ForegroundColor Gray
Write-Host "   Version     : $Version" -ForegroundColor Gray
Write-Host "   Project     : $ProjectPath" -ForegroundColor Gray
Write-Host ""

# Build the project first to ensure we have the latest assemblies
Write-Host "🔨 Building project to get builder assemblies..." -ForegroundColor Yellow
try {
    $buildOutput = dotnet build "$ProjectPath" `
        -c Release `
        -p:RevitVersion=2024 `
        --verbosity quiet `
        2>&1
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ Build failed!" -ForegroundColor Red
        Write-Host $buildOutput
        exit 1
    }
    Write-Host "✅ Build successful" -ForegroundColor Green
} catch {
    Write-Host "❌ Error during build: $_" -ForegroundColor Red
    exit 1
}

# Find the built assembly
$assemblyPath = Join-Path $ProjectDir "bin\Release\Revit2024\Quoc_MEP.dll"
if (-not (Test-Path $assemblyPath)) {
    Write-Host "❌ Assembly not found: $assemblyPath" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "📦 Loading Autodesk.PackageBuilder..." -ForegroundColor Yellow

# Load the assembly and its dependencies
Add-Type -Path $assemblyPath

# Load PackageBuilder (should be in the same directory or GAC)
$packageBuilderPath = Join-Path $ProjectDir "bin\Release\Revit2024\Autodesk.PackageBuilder.dll"
if (Test-Path $packageBuilderPath) {
    Add-Type -Path $packageBuilderPath
} else {
    Write-Host "⚠️  PackageBuilder not found at expected location, assuming it's in GAC" -ForegroundColor Yellow
}

Write-Host "✅ Assemblies loaded" -ForegroundColor Green
Write-Host ""

# Generate PackageContents.xml
Write-Host "🎨 Generating PackageContents.xml..." -ForegroundColor Yellow
try {
    $packageBuilder = New-Object Quoc_MEP.Builders.QuocMEPPackageBuilder($Version)
    $packageXmlPath = Join-Path $OutputDir "PackageContents.xml"
    $null = $packageBuilder.Build($packageXmlPath)
    
    if (Test-Path $packageXmlPath) {
        Write-Host "✅ PackageContents.xml created: $packageXmlPath" -ForegroundColor Green
        
        # Display preview
        Write-Host ""
        Write-Host "📄 Preview:" -ForegroundColor Cyan
        $content = Get-Content $packageXmlPath -Raw
        Write-Host $content -ForegroundColor Gray
    } else {
        Write-Host "❌ Failed to create PackageContents.xml" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "❌ Error generating PackageContents.xml: $_" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🎨 Generating Quoc_MEP.addin..." -ForegroundColor Yellow
try {
    $addinBuilder = New-Object Quoc_MEP.Builders.QuocMEPAddinBuilder
    $addinPath = Join-Path $OutputDir "Quoc_MEP.addin"
    $null = $addinBuilder.Build($addinPath)
    
    if (Test-Path $addinPath) {
        Write-Host "✅ Quoc_MEP.addin created: $addinPath" -ForegroundColor Green
        
        # Display preview
        Write-Host ""
        Write-Host "📄 Preview:" -ForegroundColor Cyan
        $content = Get-Content $addinPath -Raw
        Write-Host $content -ForegroundColor Gray
    } else {
        Write-Host "❌ Failed to create Quoc_MEP.addin" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "❌ Error generating Quoc_MEP.addin: $_" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  ✅ All files generated successfully!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Generated files:" -ForegroundColor Green
Write-Host "  - $packageXmlPath" -ForegroundColor Gray
Write-Host "  - $addinPath" -ForegroundColor Gray
Write-Host ""
