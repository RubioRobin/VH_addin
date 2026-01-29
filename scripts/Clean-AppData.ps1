
Write-Host "Cleaning AppData duplicates..." -ForegroundColor Cyan

$versions = @("2023", "2024", "2025")
$path = "$env:APPDATA\Autodesk\Revit\Addins"

foreach ($v in $versions) {
    if (Test-Path "$path\$v\VH_ViewDuplicator.addin") {
        Write-Host "Removing from AppData\$v" -ForegroundColor Yellow
        Remove-Item "$path\$v\VH_ViewDuplicator.addin" -Force
    }
    if (Test-Path "$path\$v\VH_ViewDuplicator.dll") {
        Write-Host "Removing DLL from AppData\$v" -ForegroundColor Yellow
        Remove-Item "$path\$v\VH_ViewDuplicator.dll" -Force
    }
}
Write-Host "Done. Restart Revit." -ForegroundColor Green

if ($Host.Name -eq "ConsoleHost") {
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
