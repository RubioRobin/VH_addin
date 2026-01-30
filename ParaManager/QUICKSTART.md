# ParaManager - Quick Start Guide

Due to API differences between Revit versions, the plugin requires some adjustments to build successfully. Here's what you need to know:

## Current Status

The plugin structure is complete with all major features implemented:
- ✅ General Parameters management
- ✅ Model Parameters transfer
- ✅ Family Parameters (nested instances)
- ✅ Shared Parameters Editor
- ✅ Excel import/export
- ⚠️ Build requires API compatibility fixes for Revit 2023

## Known Issues

Revit 2023 introduced ForgeTypeId to replace the legacy ParameterType enum. The current codebase needs updates in:
1. `GeneralParametersWindow.xaml.cs` - Parameter type handling
2. `ModelParametersWindow.xaml.cs` - Family parameter operations
3. `FamilyParametersWindow.xaml.cs` - Nested parameter display
4. `ExcelHelper.cs` - Parameter type serialization

## Recommended Next Steps

1. **For immediate testing**: Build for Revit 2024 or 2025 which have better API compatibility
   ```powershell
   dotnet build ParaManager.csproj -c Release2024
   ```

2. **For Revit 2023 support**: The following files need ForgeTypeId compatibility updates:
   - Replace `ParameterType` enum usage with `ForgeTypeId` and `SpecTypeId`
   - Update parameter creation to use `GetDataType()` instead of `.ParameterType`
   - Modify Excel export/import to handle ForgeTypeId serialization

## Alternative Approach

Consider using a helper class to abstract the API differences:
```csharp
public static class ParameterTypeHelper
{
    public static ForgeTypeId GetForgeTypeId(string typeName)
    {
        // Map string names to ForgeTypeId
    }
    
    public static string GetTypeName(ForgeTypeId typeId)
    {
        // Convert ForgeTypeId to display string
    }
}
```

## Contact

For assistance with Revit 2023 compatibility, please review the Revit API documentation for ForgeTypeId migration.
