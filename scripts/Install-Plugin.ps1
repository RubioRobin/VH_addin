$revitVersions = @("2023", "2024", "2025")
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Join-Path $scriptDir "..\src\VH_Addin"

# Ensure Revit is closed to avoid file locks
Write-Host "Closing Revit if open..." -ForegroundColor Yellow
Stop-Process -Name Revit -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

foreach ($version in $revitVersions) {
    $addinsPath = "C:\ProgramData\Autodesk\Revit\Addins\$version"
    $dllPath = Join-Path $projectDir "bin\Debug$version\VH_Tools.dll"
    $addinPath = Join-Path $projectDir "VH_Tools.addin"
    
    if (Test-Path $addinsPath) {
        Write-Host "`n--- Processing Revit $version ---" -ForegroundColor Cyan
        
        if (Test-Path $dllPath) {
            # Try to delete existing DLL first to detect locks
            $targetDll = Join-Path $addinsPath "VH_Tools.dll"
            if (Test-Path $targetDll) {
                try {
                    Remove-Item $targetDll -Force -ErrorAction Stop
                    Write-Host "  Successfully removed old DLL" -ForegroundColor Gray
                }
                catch {
                    Write-Host "  CRITICAL ERROR: Could not remove $targetDll. Is Revit still open?" -ForegroundColor Red
                    continue
                }
            }

            Write-Host "  Copying plugin files to $addinsPath"
            
            # Copy specific files from bin
            $buildDir = Join-Path $projectDir "bin\Debug$version"
            $filesToCopy = @("VH_Tools.dll", "VH_Tools.pdb", "VH_Tools.deps.json")
            foreach ($f in $filesToCopy) {
                $src = Join-Path $buildDir $f
                if (Test-Path $src) {
                    Copy-Item $src $addinsPath -Force
                }
            }
            
            # Copy .addin file
            if (Test-Path $addinPath) {
                try {
                    Copy-Item $addinPath $addinsPath -Force -ErrorAction SilentlyContinue
                }
                catch {}
            }


            # Force copy XAML files from Views folder
            if (Test-Path (Join-Path $projectDir "Views")) {
                Get-ChildItem (Join-Path $projectDir "Views") -Filter "*.xaml" | ForEach-Object {
                    Copy-Item $_.FullName $addinsPath -Force
                }
            }
            
            # Copy ViewDuplicatorMainContent.xaml specifically
            $mainContentXaml = Join-Path $projectDir "ViewDuplicatorMainContent.xaml"
            if (Test-Path $mainContentXaml) {
                Copy-Item $mainContentXaml $addinsPath -Force
            }
            
            Write-Host "  INSTALL SUCCESSFUL for $version" -ForegroundColor Green
        }
        else {
            Write-Host "  SKIPPING: Source DLL not found: $dllPath" -ForegroundColor Yellow
        }
    }
}
Write-Host "`nInstall Complete."
if ($Host.Name -eq "ConsoleHost") {
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
