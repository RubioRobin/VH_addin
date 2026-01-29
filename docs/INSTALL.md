
# VH Tools - Installatie Instructies

## 🎉 Compilatie Succesvol!


Alle Revit versies zijn succesvol gecompileerd:
- ✅ Revit 2023: `bin\Debug2023\VH_ViewDuplicator.dll`
- ✅ Revit 2024: `bin\Debug2024\VH_ViewDuplicator.dll`
- ✅ Revit 2025: `bin\Debug2025\VH_ViewDuplicator.dll`

## Installatie


### Stap 1: Kopieer de DLL naar Revit Addins folder (VH Tools)

Voor **Revit 2023**:
```powershell
Copy-Item "bin\Debug2023\VH_ViewDuplicator.dll" "C:\ProgramData\Autodesk\Revit\Addins\2023\"
Copy-Item "VH_ViewDuplicator.addin" "C:\ProgramData\Autodesk\Revit\Addins\2023\"
```

Voor **Revit 2024**:
```powershell
Copy-Item "bin\Debug2024\VH_ViewDuplicator.dll" "C:\ProgramData\Autodesk\Revit\Addins\2024\"
Copy-Item "VH_ViewDuplicator.addin" "C:\ProgramData\Autodesk\Revit\Addins\2024\"
```

Voor **Revit 2025**:
```powershell
Copy-Item "bin\Debug2025\VH_ViewDuplicator.dll" "C:\ProgramData\Autodesk\Revit\Addins\2025\"
Copy-Item "VH_ViewDuplicator.addin" "C:\ProgramData\Autodesk\Revit\Addins\2025\"
```

### Stap 2: Pas het .addin bestand aan (indien nodig)


Open `VH_ViewDuplicator.addin` of `VH_DetailSheets.addin` in Kladblok en controleer of het pad klopt:

```xml
<Assembly>VH_ViewDuplicator.dll</Assembly>
```

Als de DLL in een andere folder staat, pas dan het volledige pad aan:

De add-in verschijnt in Revit als 'VH Tools'.
```xml
<Assembly>C:\Path\To\VH_ViewDuplicator.dll</Assembly>
```

### Stap 3: Herstart Revit

Start Revit opnieuw op. De plugin wordt automatisch geladen.

## Gebruik

1. Open Revit (2023, 2024, of 2025)
2. Ga naar de tab **"VH Engineering"** in het lint
3. Klik op **"View Duplicator"**
4. Het View Duplicator venster opent met alle functionaliteit van het Dynamo script!

## Snelle Installatie Script

Run dit PowerShell script om automatisch te installeren voor alle Revit versies:

```powershell
# Installeer voor alle geïnstalleerde Revit versies
$revitVersions = @("2023", "2024", "2025")

foreach ($version in $revitVersions) {
    $addinsPath = "C:\ProgramData\Autodesk\Revit\Addins\$version"
    
    if (Test-Path $addinsPath) {
        Write-Host "Installing for Revit $version..." -ForegroundColor Green
        
        # Kopieer DLL
        Copy-Item "bin\Debug$version\VH_ViewDuplicator.dll" $addinsPath -Force
        
        # Kopieer .addin file
        Copy-Item "VH_ViewDuplicator.addin" $addinsPath -Force
        
        Write-Host "✓ Installed to $addinsPath" -ForegroundColor Green
    } else {
        Write-Host "Revit $version not found, skipping..." -ForegroundColor Yellow
    }
}

Write-Host "`n✓ Installation complete! Restart Revit to load the plugin." -ForegroundColor Cyan
```

## Probleem oplossen

### Plugin wordt niet geladen

1. Controleer of beide bestanden in de Addins folder staan:
   - `VH_ViewDuplicator.dll`
   - `VH_ViewDuplicator.addin`

2. Open Revit en ga naar: Add-ins → External Tools
   - Kijk of er error messages zijn

3. Controleer het pad in het .addin bestand

### Button is niet zichtbaar

- De button staat in de tab **"VH Engineering"** > panel **"View Tools"**
- Als de tab niet zichtbaar is, restart Revit opnieuw

### Windows blocked the DLL

Als Windows de DLL blokkeert:
1. Rechtermuisklik op `VH_ViewDuplicator.dll`
2. Kies Properties
3. Vink "Unblock" aan onderaan
4. Klik OK
5. Restart Revit

## Updates

Na wijzigingen in de code:
1. Rebuild: `dotnet build VH_ViewDuplicator.csproj -c Debug20XX`
2. Kopieer nieuwe DLL naar Addins folder
3. Restart Revit

## Productie Build

Voor Release builds (geoptimaliseerd, geen debug symbols):

```powershell
dotnet build VH_ViewDuplicator.csproj -c Release2023
dotnet build VH_ViewDuplicator.csproj -c Release2024
dotnet build VH_ViewDuplicator.csproj -c Release2025
```

DLL's staan dan in: `bin\Release20XX\`
