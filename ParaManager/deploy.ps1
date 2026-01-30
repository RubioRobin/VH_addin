# ParaManager Deployment Script
# Deploys the add-in to Revit add-ins folders

Write-Host "Deploying ParaManager..." -ForegroundColor Cyan

$userName = $env:USERNAME
$addinsBasePath = "C:\Users\$userName\AppData\Roaming\Autodesk\Revit\Addins"

# Deploy for Revit 2023
if (Test-Path "bin\2023\ParaManager.dll") {
    $revit2023Path = "$addinsBasePath\2023"
    if (Test-Path $revit2023Path) {
        Write-Host "`nDeploying to Revit 2023..." -ForegroundColor Yellow
        Copy-Item "bin\2023\*" -Destination $revit2023Path -Recurse -Force
        Copy-Item "ParaManager.addin" -Destination $revit2023Path -Force
        Write-Host "  Deployed to: $revit2023Path" -ForegroundColor Green
    }
    else {
        Write-Host "`nRevit 2023 add-ins folder not found" -ForegroundColor Red
    }
}

# Deploy for Revit 2024
if (Test-Path "bin\2024\ParaManager.dll") {
    $revit2024Path = "$addinsBasePath\2024"
    if (Test-Path $revit2024Path) {
        Write-Host "`nDeploying to Revit 2024..." -ForegroundColor Yellow
        Copy-Item "bin\2024\*" -Destination $revit2024Path -Recurse -Force
        Copy-Item "ParaManager.addin" -Destination $revit2024Path -Force
        Write-Host "  Deployed to: $revit2024Path" -ForegroundColor Green
    }
    else {
        Write-Host "`nRevit 2024 add-ins folder not found" -ForegroundColor Red
    }
}

# Deploy for Revit 2025
if (Test-Path "bin\2025\ParaManager.dll") {
    $revit2025Path = "$addinsBasePath\2025"
    if (Test-Path $revit2025Path) {
        Write-Host "`nDeploying to Revit 2025..." -ForegroundColor Yellow
        Copy-Item "bin\2025\*" -Destination $revit2025Path -Recurse -Force
        Copy-Item "ParaManager.addin" -Destination $revit2025Path -Force
        Write-Host "  Deployed to: $revit2025Path" -ForegroundColor Green
    }
    else {
        Write-Host "`nRevit 2025 add-ins folder not found" -ForegroundColor Red
    }
}

Write-Host "`nDeployment completed!" -ForegroundColor Green
Write-Host "`nRestart Revit to load the add-in." -ForegroundColor Cyan
