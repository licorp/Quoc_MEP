# Build ScheduleManager for all Revit versions
param(
    [string[]]$Versions = @("2020", "2021", "2022", "2023", "2024"),
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$solution = "ScheduleManager.sln"

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Building ScheduleManager for Revit Versions: $($Versions -join ', ')" -ForegroundColor Cyan
Write-Host "Configuration: $Configuration" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan

$successCount = 0
$failCount = 0
$results = @()

foreach ($version in $Versions) {
    Write-Host ""
    Write-Host "----- Building for Revit $version -----" -ForegroundColor Yellow
    
    try {
        Write-Host "Restoring packages..." -ForegroundColor Gray
        & dotnet restore $solution /p:RevitVersion=$version
        
        Write-Host "Building..." -ForegroundColor Gray
        & dotnet msbuild $solution /p:RevitVersion=$version /p:Configuration=$Configuration /v:minimal
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "SUCCESS: Revit $version build succeeded" -ForegroundColor Green
            $successCount++
            $results += @{Version=$version; Status="Success"}
        } else {
            Write-Host "FAILED: Revit $version build failed with exit code $LASTEXITCODE" -ForegroundColor Red
            $failCount++
            $results += @{Version=$version; Status="Failed"}
        }
    }
    catch {
        Write-Host "ERROR: Revit $version build failed: $_" -ForegroundColor Red
        $failCount++
        $results += @{Version=$version; Status="Error: $_"}
    }
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Build Summary:" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
foreach ($result in $results) {
    if ($result.Status -eq "Success") {
        Write-Host "Revit $($result.Version): $($result.Status)" -ForegroundColor Green
    } else {
        Write-Host "Revit $($result.Version): $($result.Status)" -ForegroundColor Red
    }
}
Write-Host ""
Write-Host "Total: $successCount succeeded, $failCount failed" -ForegroundColor Cyan

if ($failCount -gt 0) {
    exit 1
}
