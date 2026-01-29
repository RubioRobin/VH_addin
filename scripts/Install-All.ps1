# Install VH Engineering plugins
$ErrorActionPreference = "Stop"
Write-Host "=== VH Engineering Plugins Installer ===" -ForegroundColor Cyan

$scriptDir = $PSScriptRoot
$projectDir = Join-Path $scriptDir "..\src\VH_Addin"
$revitVersions = @("2023", "2024", "2025")

foreach ($version in $revitVersions) {
    $addinPath = "C:\ProgramData\Autodesk\Revit\Addins\$version"
    if (-not (Test-Path $addinPath)) {
        Write-Host "Revit $version addin folder niet gevonden, overslaan." -ForegroundColor Gray
        continue
    }
    Write-Host "`nInstalleren voor Revit $version..." -ForegroundColor Yellow

    # Remove legacy addin/manifests
    $legacyFiles = @("VH_ViewDuplicator.dll", "VH_ViewDuplicator.addin")
    foreach ($f in $legacyFiles) {
        $lp = Join-Path $addinPath $f
        if (Test-Path $lp) { Remove-Item $lp -Force -ErrorAction SilentlyContinue }
    }

    # Copy Built DLLs
    $buildDir = Join-Path $projectDir "bin\Debug$version"
    if (Test-Path $buildDir) {
        $dlls = Get-ChildItem -Path $buildDir -Filter "*.dll" -File
        foreach ($dll in $dlls) {
            Copy-Item $dll.FullName $addinPath -Force
            Write-Host "  Copy: $($dll.Name)" -ForegroundColor Green
        }
        
        # Also copy XAMLs
        $viewDir = Join-Path $projectDir "Views"
        if (Test-Path $viewDir) {
            Get-ChildItem $viewDir -Filter "*.xaml" | ForEach-Object {
                Copy-Item $_.FullName $addinPath -Force
                Write-Host "  Copy: $($_.Name)" -ForegroundColor Green
            }
        }
    }
    else {
        Write-Host "  WARNING: Build dir not found: $buildDir" -ForegroundColor Red
    }

    # Copy Manifest
    $manifest = Join-Path $projectDir "VH_Tools.addin"
    if (Test-Path $manifest) {
        Copy-Item $manifest $addinPath -Force
        Write-Host "  Copy manifest: VH_Tools.addin" -ForegroundColor Green
    }
}
Write-Host "`n=== Installation Complete ===" -ForegroundColor Cyan
Write-Host "Restart Revit to load changes."
