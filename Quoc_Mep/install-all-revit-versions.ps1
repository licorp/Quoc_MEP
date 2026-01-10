# Install Nice3point.Revit.Api for all Revit versions (2020-2026)
$versions = @("2020", "2021", "2022", "2023", "2024", "2025", "2026")
$packages = @("Nice3point.Revit.Api.RevitAPI", "Nice3point.Revit.Api.RevitAPIUI", "Nice3point.Revit.Api.AdWindows")

Write-Host "Installing Nice3point packages for Revit 2020-2026..." -ForegroundColor Green
Write-Host ""

$success = 0
$failed = 0

foreach ($ver in $versions) {
    Write-Host "Revit $ver" -ForegroundColor Cyan
    foreach ($pkg in $packages) {
        $name = $pkg.Split('.')[-1]
        Write-Host "  $name..." -NoNewline
        $out = & .\nuget.exe install $pkg -Version "$ver.*" -OutputDirectory packages -NonInteractive 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Host " OK" -ForegroundColor Green
            $success++
        } else {
            Write-Host " FAIL" -ForegroundColor Red  
            $failed++
        }
    }
}

Write-Host ""
Write-Host "Done: $success OK, $failed Failed" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Yellow" })
