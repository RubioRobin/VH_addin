# VH Daglichttool - Build Instructies

## ✅ Code is Klaar!

Alle 14 bestanden zijn succesvol aangemaakt:
- Application.cs (Ribbon)
- DaglichtCommand.cs (Main command)
- DaglichtWindow.xaml + .cs (UI)
- DaglichtCalculator.cs (Alpha)
- DaglichtCalculatorBeta.cs (Beta)
- CalculatorHelpers.cs (Utilities)
- GlassAreaCalculator.cs (Glass area)
- CbTable.cs (NEN table)
- NENCalculator.cs (NEN 55%)
- Models.cs (Data models)
- VH_DaglichtPlugin.csproj (Project)
- VH_DaglichtPlugin.addin (Manifest)
- README.md (Documentatie)

## 🔨 Build Stappen

### Optie 1: Visual Studio (Aanbevolen)

1. **Open het project:**
   ```
   Dubbelklik op: VH_DaglichtPlugin.csproj
   ```
   Of: File → Open → Project/Solution → Selecteer `VH_DaglichtPlugin.csproj`

2. **Build het project:**
   - Druk op `Ctrl+Shift+B`
   - Of: Build → Build Solution

3. **Automatische deployment:**
   De plugin wordt automatisch gekopieerd naar:
   ```
   %AppData%\Autodesk\Revit\Addins\2025\VH_DaglichtPlugin\
   ```

4. **Herstart Revit 2025**
   - "VH Tools" tab verschijnt in ribbon
   - Klik op "Daglicht Berekening"

### Optie 2: Developer Command Prompt

Als Visual Studio niet werkt, gebruik Developer Command Prompt:

1. Start → "Developer Command Prompt for VS 2022"

2. Navigeer naar project:
   ```cmd
   cd C:\Users\Stage-VHEngineering\.gemini\antigravity\scratch\VH_DaglichtPlugin
   ```

3. Build:
   ```cmd
   msbuild VH_DaglichtPlugin.csproj /p:Configuration=Debug /t:Rebuild
   ```

## ⚠️ Mogelijke Issues

### Issue: "RevitAPI.dll not found"
**Oplossing:** Pas het pad aan in `.csproj` als Revit 2025 op een andere locatie staat:
```xml
<HintPath>C:\Program Files\Autodesk\Revit 2025\RevitAPI.dll</HintPath>
```

### Issue: XAML compilation errors
**Oplossing:** Zorg dat .NET Framework 4.8 SDK is geïnstalleerd.

### Issue: Plugin laadt niet in Revit
**Controle checklist:**
1. ✓ DLL aanwezig in `%AppData%\Autodesk\Revit\Addins\2025\VH_DaglichtPlugin\`
2. ✓ .addin file aanwezig in `%AppData%\Autodesk\Revit\Addins\2025\`
3. ✓ Revit volledig herstart
4. ✓ Kijk in Revit → Add-Ins → External Tools voor foutmeldingen

## 📋 Na Succesvolle Build

### Test Checklist:
- [ ] Plugin laadt zonder errors
- [ ] "VH Tools" tab zichtbaar
- [ ] UI opent correct
- [ ] α-berekening werkt
- [ ] β-berekening werkt
- [ ] Glasoppervlakte berekening werkt
- [ ] NEN toets werkt
- [ ] CSV wordt geëxporteerd
- [ ] Parameters worden gezet

### Vereiste Revit Parameters:
Maak deze shared parameters aan in je Revit families:
```
VH_kozijn_α         → Angle
VH_kozijn_β/ε       → Angle  
VH_kozijn_Ad        → Area
VH_kozijn_Ae        → Area
VH_kozijn_Cb        → Number
VH_kozijn_Cbi       → Number
```

## 🎯 Project Locatie

```
C:\Users\Stage-VHEngineering\.gemini\antigravity\scratch\VH_DaglichtPlugin\
```

## 📚 Documentatie

Zie `README.md` voor:
- Volledige functionaliteit overzicht
- Alle vereiste parameters (Hout/Alu)
- Gebruik instructies
- Technische details

Zie `walkthrough.md` voor:
- Code architectuur
- Implementatie details
- Belangrijke code highlights
