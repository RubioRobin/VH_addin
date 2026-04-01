# VH_addin — Claude Code Context

## Project
Suite van Revit add-ins voor Van Heugten. Bevat parameter-beheer en view-duplicatie tools.

## Structuur
```
vh-addin/
├── src/
│   ├── VH_Addin/              → VH_ViewDuplicator add-in
│   │   ├── VH_ViewDuplicator.csproj
│   │   ├── Commands/          → Revit command handlers
│   │   ├── Models/            → Datamodellen
│   │   ├── Views/             → WPF XAML vensters
│   │   └── ViewModels/        → MVVM ViewModels
│   ├── ParaManager/           → Parameter-beheer add-in (.csproj)
│   └── Installer_Source/      → NSIS installer configuratie
├── archief/                   → Archief van oude versies (niet meer actief)
├── docs/                      → Documentatie
├── scripts/                   → Build/deploy scripts
└── git_push_helper.bat        → Git helper script
```

## Tech Stack
- **Taal:** C# (.NET)
- **UI:** WPF (Windows Presentation Foundation), XAML
- **Patroon:** MVVM (Model-View-ViewModel)
- **Build:** Visual Studio, MSBuild (.sln / .csproj)
- **Installer:** NSIS

## Build Targets per Revit versie
```
Debug2023 / Release2023   → Revit 2023 (net48 / .NET Framework 4.8)
Debug2024 / Release2024   → Revit 2024 (net48 / .NET Framework 4.8)
Debug2025 / Release2025   → Revit 2025 (net8.0-windows / .NET 8.0)
```

Assembly name: `VH_Tools`

## Hoe bouwen
- Open de `.sln` in Visual Studio
- Selecteer de juiste build configuratie (bijv. `Debug2025`)
- Build via `Ctrl+Shift+B`
- DLL komt in de `bin/` map van het betreffende project

## Add-ins
1. **VH_ViewDuplicator** — dupliceert Revit views met aangepaste instellingen
2. **ParaManager** — beheert parameters van Revit-elementen

## Revit integratie
- Add-ins worden geladen via `.addin` manifest bestanden
- Manifest locatie: `%APPDATA%\Autodesk\Revit\Addins\20XX\`

## Veiligheidsregels
- NOOIT bestanden verwijderen zonder bevestiging
- Wees voorzichtig met build targets — verkeerde target kan Revit crashen
- Test altijd in een Revit-testomgeving, niet direct in productie
