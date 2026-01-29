# VH Tools - Revit Add-in Suite

Dit project bevat een verzameling Revit Add-ins ontwikkeld door VH Engineering, waaronder de **View Duplicator**, **Daglichttool** (NEN 2057:2011), en diverse andere hulpmiddelen.

## Projectstructuur

De repository is als volgt georganiseerd:

*   **`/src`**: Bevat de broncode van de Revit Add-in (`VH_Addin`).
*   **`/docs`**: Gedetailleerde documentatie, handleidingen en ontwerpbeslissingen.
*   **`/scripts`**: PowerShell scripts voor het bouwen, installeren en beheren van de plugin.
*   **`/archief`**: Verouderde bronbestanden, scripts en testgegevens die niet meer actief worden gebruikt.

## Ondersteunde Revit Versies

De add-in is compatibel met de volgende versies:
- Revit 2023
- Revit 2024
- Revit 2025

## Snelle Start

### 1. Bouwen en Installeren
Gebruik de meegeleverde scripts in de `/scripts` map voor een eenvoudige installatie:

1. Open PowerShell als Administrator.
2. Navigeer naar de `/scripts` map.
3. Voer het volgende commando uit:
   ```powershell
   .\Install-All.ps1
   ```
   *Dit bouwt de plugin voor alle ondersteunde versies en installeert de manifest-bestanden in de juiste Revit mappen.*

### 2. Gebruik in Revit
Na installatie verschijnt het tabblad **VH Tools** in Revit. Hier vindt u de verschillende tools zoals de Daglichtberekening en de View Duplicator.

## Ontwikkeling

### Vereisten
- Visual Studio 2022
- .NET Framework 4.8 (voor Revit 2023/2024)
- .NET 8 (voor Revit 2025)
- Revit API access (Revit geïnstalleerd)

### Project Openen
U kunt de solution direct openen via:
- `/src/VH_Addin/VH_Addin.sln`
OF gebruik het handige script:
- `/scripts/Open-In-VS.ps1`

## Documentatie
Zie de `/docs` map voor meer specifieke handleidingen:
- [Installatie handleiding](docs/INSTALL.md)
- [Daglichtberekening Details](docs/KLAAR_VOOR_COMPILATIE.md)

## Licentie
Copyright © VH Engineering. Alle rechten voorbehouden.
