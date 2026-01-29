$versions = @("2023", "2024", "2025")
$scriptDir = $PSScriptRoot
Set-Location $scriptDir

Write-Host "=== VH Daglicht Plugin - Build & Deploy ===" -ForegroundColor Cyan

foreach ($v in $versions) {
    Write-Host "`n--- Revit $v ---" -ForegroundColor Yellow
    
    # Build
    $config = "Release"
    Write-Host "Building $config..." -ForegroundColor Gray
    dotnet build VH_DaglichtPlugin.csproj -c $config -v minimal /p:RevitVersion=$v
    
    if ($LASTEXITCODE -eq 0) {
        $targetDir = "$env:APPDATA\Autodesk\Revit\Addins\$v\VH_DaglichtPlugin"
        
        # Create target directory if it doesn't exist
        if (-not (Test-Path $targetDir)) {
            New-Item $targetDir -ItemType Directory -Force | Out-Null
        }
        
        Write-Host "Deploying to $targetDir..." -ForegroundColor Gray
        
        # Copy everything from bin (including DLLs, PDBs, deps.json)
        $binDir = "bin\$config"
        if (Test-Path $binDir) {
            Get-ChildItem "$binDir\*" -Include *.dll, *.pdb, *.deps.json, *.runtimeconfig.json -File | ForEach-Object {
                Copy-Item $_.FullName $targetDir -Force
                Unblock-File (Join-Path $targetDir $_.Name)
            }
        }

        # Copy Addin Manifest to parent directory
        $addinTargetDir = "$env:APPDATA\Autodesk\Revit\Addins\$v"
        Copy-Item "VH_DaglichtPlugin.addin" $addinTargetDir -Force
        Unblock-File (Join-Path $addinTargetDir "VH_DaglichtPlugin.addin")
        
        Write-Host "SUCCESS: $v deployed" -ForegroundColor Green
    }
    else {
        Write-Host "FAILED: Build error for $v" -ForegroundColor Red
    }
}

Write-Host "`n=== COMPLETE ===" -ForegroundColor Cyan
Write-Host "Plugin deployed to Revit Addins folders." -ForegroundColor White
Write-Host "Restart Revit versions to load the plugin.`n" -ForegroundColor Yellow
