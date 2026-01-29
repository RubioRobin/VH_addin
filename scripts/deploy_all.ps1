$versions = @("2023", "2024", "2025")
$scriptDir = $PSScriptRoot
Set-Location $scriptDir

Write-Host "=== VH Tools - Build & Deploy ===" -ForegroundColor Cyan

foreach ($v in $versions) {
    Write-Host "`n--- Revit $v ---" -ForegroundColor Yellow
    
    # Build
    $config = "Debug$v"
    Write-Host "Building $config..." -ForegroundColor Gray
    $projectPath = Join-Path $scriptDir "..\src\VH_Addin\VH_ViewDuplicator.csproj"
    dotnet build $projectPath -c $config -v minimal
    
    if ($LASTEXITCODE -eq 0) {
        $targetDir = "C:\ProgramData\Autodesk\Revit\Addins\$v"
        
        if (Test-Path $targetDir) {
            Write-Host "Deploying to $targetDir..." -ForegroundColor Gray
            
            # Copy everything from bin (including DLLs, PDBs, deps.json)
            $srcBinDir = Join-Path $scriptDir "..\src\VH_Addin\bin\$config"
            
            # Using a more robust way to get files instead of -Include with a path
            $extensions = @(".dll", ".pdb", ".json", ".xaml")
            Get-ChildItem -Path $srcBinDir -File | Where-Object { $extensions -contains $_.Extension.ToLower() } | ForEach-Object {
                Copy-Item $_.FullName $targetDir -Force
                Unblock-File (Join-Path $targetDir $_.Name)
            }

            # Copy Addin Manifest
            $addinManifest = Join-Path $scriptDir "..\src\VH_Addin\VH_Tools.addin"
            Copy-Item $addinManifest $targetDir -Force
            Unblock-File (Join-Path $targetDir "VH_Tools.addin")

            # Copy all XAML from Views (if they are needed as loose files)
            $viewsDir = Join-Path $scriptDir "..\src\VH_Addin\Views"
            Get-ChildItem (Join-Path $viewsDir "*.xaml") -ErrorAction SilentlyContinue | ForEach-Object {
                Copy-Item $_.FullName $targetDir -Force
                Unblock-File (Join-Path $targetDir $_.Name)
            }

            # Copy Resources folder
            $resDir = Join-Path $scriptDir "..\src\VH_Addin\Resources"
            if (Test-Path $resDir) {
                $targetRes = Join-Path $targetDir "Resources"
                if (-not (Test-Path $targetRes)) { New-Item $targetRes -ItemType Directory | Out-Null }
                Copy-Item (Join-Path $resDir "*") $targetRes -Force -Recurse
                Get-ChildItem $targetRes -Recurse -File | ForEach-Object { Unblock-File $_.FullName }
            }
            
            Write-Host "SUCCESS: $v deployed" -ForegroundColor Green
        }
        else {
            Write-Host "SKIP: Target directory not found" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "FAILED: Build error" -ForegroundColor Red
    }
}

Write-Host "`n=== COMPLETE ===" -ForegroundColor Cyan
Write-Host "Restart Revit to load the updated plugin.`n"
