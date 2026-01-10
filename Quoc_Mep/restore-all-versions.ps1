# Script restore Nice3point packages cho tat ca phien ban Revit
# Chay: .\restore-all-versions.ps1

param(
    [Parameter(Mandatory=$false)]
    [string[]]$Versions = @("2020", "2021", "2022", "2023", "2024", "2025", "2026")
)

$ErrorActionPreference = "Continue"
$projectFile = "Quoc_MEP.csproj"
$solutionFile = "Quoc_MEP.sln"

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Nice3point Packages Restore for All Versions" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Doc noi dung .csproj hien tai
$csprojContent = Get-Content $projectFile -Raw
$originalRevitVersion = ""

# Tim RevitVersion hien tai
if ($csprojContent -match '<RevitVersion>(\d+)</RevitVersion>') {
    $originalRevitVersion = $Matches[1]
    Write-Host "Original RevitVersion: $originalRevitVersion" -ForegroundColor Yellow
}

$successCount = 0
$failCount = 0
$results = @()

foreach ($version in $Versions) {
    Write-Host ""
    Write-Host "--------------------------------------------------" -ForegroundColor Gray
    Write-Host "Restoring packages for Revit $version..." -ForegroundColor Green
    
    # Cap nhat RevitVersion trong .csproj
    $csprojContent = Get-Content $projectFile -Raw
    $csprojContent = $csprojContent -replace '<RevitVersion>\d+</RevitVersion>', "<RevitVersion>$version</RevitVersion>"
    Set-Content $projectFile -Value $csprojContent -NoNewline
    
    Write-Host "  Updated RevitVersion to $version" -ForegroundColor Gray
    
    # Restore packages
    $restoreStart = Get-Date
    
    & .\nuget.exe restore $solutionFile -NonInteractive | Out-Null
    
    $restoreEnd = Get-Date
    $duration = ($restoreEnd - $restoreStart).TotalSeconds
    
    if ($LASTEXITCODE -eq 0) {
        $successCount++
        Write-Host "  Success - Revit $version packages restored! ($([math]::Round($duration, 2))s)" -ForegroundColor Green
        
        # Kiem tra packages da tai
        $packagePath = "$env:USERPROFILE\.nuget\packages\nice3point.revit.api.revitapi\$version.*"
        $installedPackage = Get-ChildItem $packagePath -Directory -ErrorAction SilentlyContinue | Select-Object -First 1
        
        if ($installedPackage) {
            Write-Host "  Installed: $($installedPackage.Name)" -ForegroundColor Gray
        }
        
        $results += [PSCustomObject]@{
            Version = $version
            Status = "Success"
            Duration = "$([math]::Round($duration, 2))s"
        }
    }
    else {
        $failCount++
        Write-Host "  Failed - Revit $version restore failed!" -ForegroundColor Red
        
        $results += [PSCustomObject]@{
            Version = $version
            Status = "Failed"
            Duration = "$([math]::Round($duration, 2))s"
        }
    }
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Restore Summary" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Display results table
$results | Format-Table -AutoSize

Write-Host ""
Write-Host "Total: $($Versions.Count) versions" -ForegroundColor White
Write-Host "Success: $successCount" -ForegroundColor Green
Write-Host "Failed: $failCount" -ForegroundColor Red
Write-Host ""

# Khoi phuc lai RevitVersion ban dau
if ($originalRevitVersion) {
    Write-Host "Restoring original RevitVersion to $originalRevitVersion..." -ForegroundColor Yellow
    $csprojContent = Get-Content $projectFile -Raw
    $csprojContent = $csprojContent -replace '<RevitVersion>\d+</RevitVersion>', "<RevitVersion>$originalRevitVersion</RevitVersion>"
    Set-Content $projectFile -Value $csprojContent -NoNewline
    Write-Host "RevitVersion restored to $originalRevitVersion" -ForegroundColor Green
}

Write-Host ""

if ($failCount -eq 0) {
    Write-Host "All packages restored successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "  1. Run: .\build-all-versions.ps1" -ForegroundColor White
    Write-Host "  2. Check outputs in: bin\Release\Revit{Version}\" -ForegroundColor White
    exit 0
}
else {
    Write-Host "Some restores failed. Check the output above for details." -ForegroundColor Yellow
    exit 1
}
