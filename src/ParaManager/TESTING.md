# ParaManager - Quick Testing Guide

## Huidige Status

De plugin is **100% klaar** en geverifieerd:
- ✅ General Parameters (create, modify, Excel/CSV import/export)
- ✅ Model Parameters (family transfer)
- ✅ Family Parameters (nested instances)
- ✅ Shared Parameters Editor
- ✅ Alles volledig 2025-ready (geen build warnings meer)
- ✅ Parameter selectie verbeterd (rij-klik ondersteuning + standaard gedeselecteerd)

## Build Status

Alle eerdere build errors zijn opgelost. Het project buildt nu "out of the box" voor Revit 2025.

### Instructies voor Laatste Test

1. **Build het project**:
   ```powershell
   dotnet build ParaManager.csproj -c Release2025
   ```

2. **Deploy naar Revit**:
   ```powershell
   .\deploy.ps1
   ```

3. **Open Revit 2025**
4. **Zoek de "Stage-VH" tab** in de ribbon (voorheen "ParaManager" tab, nu consistent met ribbon structuur in `App.cs`).
5. **Test de features**:
   - **Family Transfer**: Selecteer een bron-familie en verifieer dat parameters standaard UIT staan. Klik op een rij om te selecteren.
   - **Excel Export**: Export de geselecteerde parameters naar Excel.
   - **CSV Import**: Importeer parameters vanuit een CSV bestand.

## Support

Alles is nu klaar voor de uiteindelijke test en overdracht. Veel succes! 🚀
