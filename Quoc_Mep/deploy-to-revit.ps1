# Deploy Quoc MEP RibbonHost to Revit
# This script copies the RibbonHost and all command DLLs to Revit's Addins folder

param(
    [string]$RevitVersion = "2023"
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Deploying Quoc MEP to Revit $RevitVersion" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Paths
$sourcePath = "$PSScriptRoot\bin\Release\Revit$RevitVersion"
$addinsPath = "$env:APPDATA\Autodesk\Revit\Addins\$RevitVersion"
$targetFolder = "$addinsPath\QuocMEP"

# Check if source exists
if (!(Test-Path $sourcePath)) {
    Write-Host "ERROR: Source folder not found: $sourcePath" -ForegroundColor Red
    Write-Host "Please build the project first using: .\build-ribbonhost.ps1" -ForegroundColor Yellow
    exit 1
}

# Create target folder
Write-Host "`nCreating target folder..." -ForegroundColor Yellow
if (!(Test-Path $targetFolder)) {
    New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
    Write-Host "Created: $targetFolder" -ForegroundColor Green
} else {
    Write-Host "Folder exists: $targetFolder" -ForegroundColor Green
}

# Copy DLLs
Write-Host "`nCopying DLL files..." -ForegroundColor Yellow
$dllFiles = Get-ChildItem -Path $sourcePath -Filter "*.dll" | Where-Object {
    $_.Name -notlike "Autodesk.*" -and
    $_.Name -notlike "RevitAPI*" -and
    $_.Name -notlike "System.*" -and
    $_.Name -notlike "Microsoft.*"
}

$copiedCount = 0
foreach ($dll in $dllFiles) {
    Copy-Item $dll.FullName -Destination $targetFolder -Force
    Write-Host "  Copied: $($dll.Name)" -ForegroundColor Gray
    $copiedCount++
}
Write-Host "Total DLLs copied: $copiedCount" -ForegroundColor Green

# Copy .addin file
Write-Host "`nCopying .addin manifest..." -ForegroundColor Yellow
$addinSource = "$PSScriptRoot\Quoc_MEP_SharedRibbon.addin"
$addinTarget = "$addinsPath\Quoc_MEP_SharedRibbon.addin"

if (Test-Path $addinSource) {
    # Update .addin file to point to correct location
    $addinContent = Get-Content $addinSource -Raw
    $addinContent = $addinContent -replace '<Assembly>.*?</Assembly>', "<Assembly>$targetFolder\Quoc_MEP.RibbonHost.dll</Assembly>"
    
    Set-Content -Path $addinTarget -Value $addinContent -Force
    Write-Host "  Created: Quoc_MEP_SharedRibbon.addin" -ForegroundColor Green
} else {
    Write-Host "  WARNING: .addin file not found at $addinSource" -ForegroundColor Yellow
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Deployment Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Target folder: $targetFolder" -ForegroundColor White
Write-Host "Files deployed: $copiedCount DLLs + 1 .addin" -ForegroundColor White
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Close Revit if running" -ForegroundColor White
Write-Host "2. Start Revit $RevitVersion" -ForegroundColor White
Write-Host "3. Look for 'Quoc MEP' tab in ribbon" -ForegroundColor White
Write-Host "`nNote: If tab doesn't appear, check:" -ForegroundColor Yellow
Write-Host "  - Revit Addins Manager (should show 'Quoc MEP - Shared Ribbon')" -ForegroundColor Gray
Write-Host "  - File path in .addin matches deployment location" -ForegroundColor Gray
