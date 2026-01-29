$projectFile = "..\src\VH_Addin\VH_ViewDuplicator.csproj"
$configurations = @("Debug2023", "Debug2024", "Debug2025")

foreach ($config in $configurations) {
    Write-Host "`n--- Building $config ---" -ForegroundColor Cyan
    dotnet build $projectFile -c $config
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Success $config" -ForegroundColor Green
    }
    else {
        Write-Host "Failed $config" -ForegroundColor Red
    }
}
Write-Host "`nBuild Process Complete" -ForegroundColor Cyan
