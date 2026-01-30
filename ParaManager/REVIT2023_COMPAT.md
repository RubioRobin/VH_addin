# ParaManager - Revit 2023 API Compatibility Guide

## Overview

Revit 2023 introduced significant API changes, replacing the `ParameterType` enum with the `ForgeTypeId` system. This guide explains the changes needed to make ParaManager fully compatible with Revit 2023.

## Key API Changes in Revit 2023

### 1. ParameterType → ForgeTypeId

**Before (Revit 2022 and earlier):**
```csharp
ParameterType paramType = ParameterType.Text;
Definition def = ...;
ParameterType type = def.ParameterType;
```

**After (Revit 2023+):**
```csharp
ForgeTypeId paramType = SpecTypeId.String.Text;
Definition def = ...;
ForgeTypeId type = def.GetDataType();
```

### 2. Parameter Creation

**Before:**
```csharp
ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(name, ParameterType.Text);
```

**After:**
```csharp
ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(name, SpecTypeId.String.Text);
```

## Files Requiring Updates

### 1. GeneralParametersWindow.xaml.cs

**Lines to Fix:**
- Line 31-33: Parameter type ComboBox population
- Line 138: Parameter type selection
- Line 261: CreateProjectParameter call

**Solution:**
```csharp
// Initialize parameter types
private void InitializeControls()
{
    // Use ParameterTypeHelper
    foreach (string typeName in ParameterTypeHelper.GetAllTypeNames())
    {
        cmbParameterType.Items.Add(typeName);
    }
    if (cmbParameterType.Items.Count > 0)
        cmbParameterType.SelectedIndex = 0;
}

// Create parameter
string selectedType = cmbParameterType.SelectedItem.ToString();
ForgeTypeId paramTypeId = ParameterTypeHelper.GetForgeTypeId(selectedType);

bool success = ParameterHelper.CreateProjectParameter(
    _doc,
    paramName,
    paramTypeId,  // Use ForgeTypeId instead of ParameterType
    group,
    isInstance,
    selectedCategories);
```

### 2. ModelParametersWindow.xaml.cs

**Lines to Fix:**
- Line 71, 84: Reading parameter types from family parameters
- Line 221: Creating parameters in destination families

**Solution:**
```csharp
// Reading parameter type
FamilyParameterData paramData = new FamilyParameterData
{
    Name = param.Definition.Name,
    ParameterType = ParameterTypeHelper.GetParameterTypeString(param.Definition),
    IsInstance = param.IsInstance,
    Formula = param.Formula ?? "",
    Group = param.Definition.ParameterGroup
};

// Creating parameter
ForgeTypeId typeId = ParameterTypeHelper.GetForgeTypeId(param.ParameterType);
famMgr.AddParameter(param.Name, param.Group, typeId, param.IsInstance);
```

### 3. FamilyParametersWindow.xaml.cs

**Lines to Fix:**
- Line 73, 200: Displaying parameter types

**Solution:**
```csharp
FamilyParameterData paramData = new FamilyParameterData
{
    Name = param.Definition.Name,
    ParameterType = ParameterTypeHelper.GetParameterTypeString(param.Definition),
    IsInstance = param.IsInstance,
    Formula = param.Formula ?? "",
    Group = param.Definition.ParameterGroup
};
```

### 4. ExcelHelper.cs

**Lines to Fix:**
- Line 31: worksheet.Clear() - EPPlus 7.x doesn't have Clear method
- Line 122, 124: ParameterType parsing
- Line 168: worksheet.Clear()

**Solution:**
```csharp
// Instead of worksheet.Clear(), delete all worksheets individually
while (package.Workbook.Worksheets.Count > 0)
{
    package.Workbook.Worksheets.Delete(0);
}

// For parameter type handling
worksheet.Cells[row, 2].Value = param.ParameterType; // Already a string from ParameterData
```

### 5. ParameterData.cs

**Current Implementation:**
```csharp
public class ParameterData
{
    public string Name { get; set; }
#if REVIT2023
    public ForgeTypeId ParameterTypeId { get; set; }
    public string ParameterType => ParameterTypeId?.TypeId ?? "Unknown";
#else
    public ParameterType ParameterType { get; set; }
#endif
    // ... rest of properties
}
```

**Recommended Simplification:**
```csharp
public class ParameterData
{
    public string Name { get; set; }
    public string ParameterType { get; set; }  // Always use string for display
    public ForgeTypeId ParameterTypeId { get; set; }  // Store ForgeTypeId for API calls
    // ... rest of properties
}
```

## Implementation Strategy

### Phase 1: Update Helper Classes ✅
- [x] Create `ParameterTypeHelper` for unified API
- [x] Update `ParameterHelper` to use ForgeTypeId
- [x] Create `CsvHelper` for CSV operations
- [x] Create `SharedParameterHelper` for merge operations

### Phase 2: Update UI Windows
- [ ] Update `GeneralParametersWindow` to use ParameterTypeHelper
- [ ] Update `ModelParametersWindow` to use ParameterTypeHelper
- [ ] Update `FamilyParametersWindow` to use ParameterTypeHelper
- [ ] Add CSV event handlers to GeneralParametersWindow

### Phase 3: Fix Excel Helper
- [ ] Replace `worksheet.Clear()` with proper deletion
- [ ] Update parameter type serialization

### Phase 4: Testing
- [ ] Build for Revit 2023
- [ ] Build for Revit 2024
- [ ] Build for Revit 2025
- [ ] Test all features in each version

## Quick Fix Script

For rapid testing, you can build for Revit 2024/2025 which have better backward compatibility:

```powershell
cd c:\Users\Stage-VHEngineering\Downloads\ParaManager
dotnet build ParaManager.csproj -c Release2024
.\deploy.ps1
```

## Next Steps

1. Apply the fixes outlined above to each file
2. Test compilation for Revit 2023
3. Verify functionality in all three Revit versions
4. Update documentation with any additional findings

## Resources

- [Revit 2023 API Changes](https://help.autodesk.com/view/RVT/2023/ENU/?guid=Revit_API_Revit_API_Developers_Guide_Introduction_Changes_and_Additions_html)
- [ForgeTypeId Documentation](https://www.revitapidocs.com/2023/fb011c91-be7e-f737-28c7-3f1e1917a0e0.htm)
