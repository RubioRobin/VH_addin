# ParaManager Build Script
# Builds the add-in for Revit 2023, 2024, and 2025

param(
    [string]$Configuration = "Release"
)

Write-Host "Building ParaManager for all Revit versions..." -ForegroundColor Cyan

# Build for Revit 2023
Write-Host "`nBuilding for Revit 2023..." -ForegroundColor Yellow
dotnet build ParaManager.csproj -c "$($Configuration)2023"

# Build for Revit 2024
Write-Host "`nBuilding for Revit 2024..." -ForegroundColor Yellow
dotnet build ParaManager.csproj -c "$($Configuration)2024"

# Build for Revit 2025
Write-Host "`nBuilding for Revit 2025..." -ForegroundColor Yellow
dotnet build ParaManager.csproj -c "$($Configuration)2025"

Write-Host "`nBuild completed!" -ForegroundColor Green
Write-Host "`nOutput locations:" -ForegroundColor Cyan
Write-Host "  Revit 2023: bin\2023\" -ForegroundColor White
Write-Host "  Revit 2024: bin\2024\" -ForegroundColor White
Write-Host "  Revit 2025: bin\2025\" -ForegroundColor White

# Check if builds were successful
$success = $true
if (-not (Test-Path "bin\2023\ParaManager.dll")) {
    Write-Host "`nWarning: Revit 2023 build not found" -ForegroundColor Red
    $success = $false
}
if (-not (Test-Path "bin\2024\ParaManager.dll")) {
    Write-Host "`nWarning: Revit 2024 build not found" -ForegroundColor Red
    $success = $false
}
if (-not (Test-Path "bin\2025\ParaManager.dll")) {
    Write-Host "`nWarning: Revit 2025 build not found" -ForegroundColor Red
    $success = $false
}

if ($success) {
    Write-Host "`nAll builds successful!" -ForegroundColor Green
    
    # Ask if user wants to deploy
    $deploy = Read-Host "`nDo you want to deploy to Revit add-ins folder? (y/n)"
    if ($deploy -eq 'y' -or $deploy -eq 'Y') {
        & "$PSScriptRoot\deploy.ps1"
    }
}
