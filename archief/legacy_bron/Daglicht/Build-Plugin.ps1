# VH Daglicht Plugin Build Script
Write-Host "`n=== VH Daglicht Plugin Build ===" -ForegroundColor Cyan

# Find MSBuild
$msbuildPaths = @(
    "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
)

$msbuild = $null
foreach ($path in $msbuildPaths) {
    if (Test-Path $path) {
        $msbuild = $path
        Write-Host "Found MSBuild: $msbuild" -ForegroundColor Green
        break
    }
}

if (-not $msbuild) {
    Write-Host "ERROR: MSBuild not found!" -ForegroundColor Red
    Write-Host "Please build manually in Visual Studio (Ctrl+Shift+B)" -ForegroundColor Yellow
    pause
    exit 1
}

# Build the project
Write-Host "`nBuilding Release configuration..." -ForegroundColor Cyan
& $msbuild "VH_DaglichtPlugin.csproj" /p:Configuration=Release /t:Rebuild /v:minimal /nologo

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n=== BUILD SUCCESSFUL ===" -ForegroundColor Green
    Write-Host "`nPlugin deployed to:" -ForegroundColor Cyan
    Write-Host "$env:APPDATA\Autodesk\Revit\Addins\2025\VH_DaglichtPlugin\" -ForegroundColor White
    Write-Host "`nYou can now test in Revit 2025!" -ForegroundColor Green
} else {
    Write-Host "`n=== BUILD FAILED ===" -ForegroundColor Red
}

pause
