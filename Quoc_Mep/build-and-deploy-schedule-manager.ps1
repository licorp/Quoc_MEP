# ScheduleManager Full Build and Deploy Script
# This script builds Schedule Manager with full UI functionality and deploys to Revit

param(
    [string]$RevitVersion = "2020"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Building Schedule Manager with Full UI" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Step 1: Build RibbonHost
Write-Host "`nStep 1: Building RibbonHost..." -ForegroundColor Yellow
cd "Quoc_MEP.RibbonHost"
dotnet build -c Release /p:RevitVersion=$RevitVersion
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ RibbonHost build failed!" -ForegroundColor Red
    exit 1
}
cd ..

# Step 2: Build ScheduleManager
Write-Host "`nStep 2: Building ScheduleManager..." -ForegroundColor Yellow
cd "ScheduleManager"
dotnet build -c Release /p:RevitVersion=$RevitVersion
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ ScheduleManager build failed!" -ForegroundColor Red
    exit 1
}
cd ..

# Step 3: Deploy to Revit
Write-Host "`nStep 3: Deploying to Revit $RevitVersion..." -ForegroundColor Yellow
.\deploy-to-revit.ps1 -RevitVersion $RevitVersion

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "✅ Build and Deploy Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
