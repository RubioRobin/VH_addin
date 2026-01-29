# VH Daglichttool - Revit Plugin

## Overzicht

De VH Daglichttool is een complete Revit add-in voor daglichtberekeningen volgens NEN-normen. Het combineert α-waaier berekeningen, β-hoek analyse, glasoppervlakte berekeningen en NEN 55% toetsing in één geïntegreerde tool.

## Functionaliteit

### 1. α-Berekening (Alpha Waaier)
- 11 stralen met hoeken van -50° tot +50° (stappen van 10°)
- Ray casting naar obstakels (muren)
- Berekening van α-hoek op basis van X-afstand en hoogteverschil
- 3D visualisatie met X-lijnen en A-lijnen
- Automatische parameter `VH_kozijn_α` update

### 2. β-Berekening (Beta Hoek)
- Echte 3D geometrie-analyse
- Detectie van belemmeringen in doorsnedevlak
- Overstek en wanden boven kozijn
- Maximale zoekafstand: 5 meter
- Automatische parameter `VH_kozijn_β/ε` update
- Visualisatie met β-lijnen

### 3. Glasoppervlakte Berekening
- Ondersteuning voor HOUT en ALU kozijnen
- Automatische detectie van kozijntype
- Vakverdeling (kolommen en rijen)
- Operabele vakken detectie (raam, draai, kiep, etc.)
- Aftrek voor raamhout bij operabele vakken
- 600mm afkapgrens boven level
- Parameters: `VH_kozijn_Ad` of `berekende_Ad`

### 4. NEN 55% Daglichttoets
- Cb-waarde lookup uit NEN-tabel (α, β)
- Ae berekening per kozijn: Ae = Ad × Cb × Cu × CLTA
- Koppeling kozijnen aan VG-gebieden (Areas)
- Point-in-polygon detectie
- 55% toetsing per verblijfsgebied
- Cbi berekening: Cbi = Ae_VG / (0.55 × A_VG)
- Automatische CSV export naar bureaublad

## Installatie

### Automatische Deployment
Het project is geconfigureerd voor automatische deployment naar Revit 2025:

```bash
# Build het project in Visual Studio
# De plugin wordt automatisch gekopieerd naar:
# %AppData%\Autodesk\Revit\Addins\2025\VH_DaglichtPlugin\
```

### Handmatige Installatie
1. Kopieer `VH_DaglichtPlugin.dll` naar een map, bijv:
   `C:\ProgramData\Autodesk\Revit\Addins\2025\VH_DaglichtPlugin\`

2. Kopieer `VH_DaglichtPlugin.addin` naar:
   `%AppData%\Autodesk\Revit\Addins\2025\`

3. Herstart Revit

## Gebruik

### UI Opties

**Selectie:**
- Kozijnen in actieve view
- Kozijnen op huidig level
- Alleen huidige selectie
- Alle kozijnen in model

**Berekeningen:**
- ☑ Lijnen tekenen (waaier + β- en X/A-lijnen)
- ☑ α berekenen
- ☑ β berekenen
- ☑ Glasoppervlakte berekenen
- ☐ Alleen kozijnen die grenzen aan VG-gebieden (Areas)

### Vereiste Parameters

#### Kozijn Instance Parameters:
- `VH_kozijn_α` (Angle) - Output
- `VH_kozijn_β/ε` (Angle) - Output
- `VH_kozijn_Ad` (Area) - Input/Output
- `VH_kozijn_Ae` (Area) - Output
- `VH_kozijn_Cb` (Number) - Output
- `VH_kozijn_Cbi` (Number) - Output

#### Hout Kozijnen:
- `dikte_bovendorpel`
- `dikte_eindstijlen`
- `dikte_onderdorpel`
- `dikte_tussenstijl`
- `dikte_tussendorpel`
- `aantal_tussenstijlen`
- `aantal_tussendorpels`
- `breedte_raamhout`

#### Alu Kozijnen:
- `aanzicht_raamprofiel`
- `aanzicht_stijl`
- `offset_stijl`
- `extra_stelruimte_onder`
- `aanzicht_onderdorpel`
- `aanzicht_tussenstijl`
- `aanzicht_tussendorpel`

#### VG Areas:
- `VG_groep` (optioneel)

## Output

### Revit Parameters
Alle berekende waarden worden automatisch teruggeschreven naar de kozijn parameters.

### CSV Export
Gedetailleerd rapport per verblijfsgebied met:
- Kozijnoverzicht per VG
- α, β, Cb, Ae waarden per kozijn
- Totaal Ae per VG
- NEN 55% toetsresultaat
- Benodigde VG-reductie (indien niet voldaan)

Bestandsnaam: `NEN_daglicht_YYYYMMDD_HHMMSS_detail.csv`
Locatie: Bureaublad

### Visualisatie
Model curves met custom line styles:
- **VH_Alpha** (rood): α-waaier stralen
- **VH_Beta** (blauw): β-hoek lijnen
- **VH_Alpha_X** (rood): 3D X-lijnen en A-lijnen

## Technische Details

### Architectuur
```
VH_DaglichtPlugin/
├── Application.cs              # Ribbon integration
├── DaglichtCommand.cs          # Main command orchestration
├── DaglichtWindow.xaml         # WPF UI
├── DaglichtWindow.xaml.cs      # UI code-behind
├── Models.cs                   # Data models
├── DaglichtCalculator.cs       # Alpha calculation
├── DaglichtCalculatorBeta.cs   # Beta calculation
├── CalculatorHelpers.cs        # Utility methods
├── GlassAreaCalculator.cs      # Glass area logic
├── CbTable.cs                  # NEN Cb lookup table
└── NENCalculator.cs            # NEN 55% compliance
```

### Constanten
- Ray count: 11
- Angle range: -50° tot +50°
- Ray length: 5000mm
- Max β distance: 5000mm
- 600mm afkapgrens
- NEN eis: 55% (0.55)

## Revit Versies

Getest met:
- Revit 2025

Voor andere versies: pas `<RevitVersion>` aan in `.csproj`

## Licentie

© VH Engineering

## Contact

Voor vragen of ondersteuning, neem contact op met VH Engineering.
