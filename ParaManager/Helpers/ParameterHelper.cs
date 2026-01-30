using Autodesk.Revit.DB;
using ParaManager.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ParaManager.Helpers
{
    public static class ParameterHelper
    {
        /// <summary>
        /// Creates a project parameter and binds it to specified categories
        /// </summary>
        public static bool CreateProjectParameter(
            Document doc,
            string parameterName,
            ForgeTypeId parameterTypeId,
#if REVIT2025
            ForgeTypeId group,
#else
            BuiltInParameterGroup group,
#endif
            bool isInstance,
            IEnumerable<Category> categories)
        {
            try
            {
                using (Transaction trans = new Transaction(doc, "Create Project Parameter"))
                {
                    trans.Start();

                    // Get or create shared parameter
                    DefinitionFile defFile = doc.Application.OpenSharedParameterFile();
                    if (defFile == null)
                    {
                        // Create temporary shared parameter file
                        string tempPath = System.IO.Path.GetTempFileName();
                        doc.Application.SharedParametersFilename = tempPath;
                        defFile = doc.Application.OpenSharedParameterFile();
                    }

                    DefinitionGroup defGroup = defFile.Groups.get_Item("ParaManager") 
                        ?? defFile.Groups.Create("ParaManager");

                    Definition def = defGroup.Definitions.get_Item(parameterName);
                    if (def == null)
                    {
                        ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(parameterName, parameterTypeId);
                        def = defGroup.Definitions.Create(options);
                    }

                    // Create category set
                    CategorySet catSet = doc.Application.Create.NewCategorySet();
                    foreach (Category cat in categories)
                    {
                        catSet.Insert(cat);
                    }

                    // Bind parameter
                    Binding binding = isInstance 
                        ? doc.Application.Create.NewInstanceBinding(catSet) as Binding
                        : doc.Application.Create.NewTypeBinding(catSet) as Binding;

                    BindingMap bindingMap = doc.ParameterBindings;
                    bindingMap.Insert(def, binding, group);

                    trans.Commit();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Gets all project parameters in the document
        /// </summary>
        public static List<ParameterData> GetProjectParameters(Document doc)
        {
            List<ParameterData> parameters = new List<ParameterData>();
            BindingMap bindingMap = doc.ParameterBindings;

            DefinitionBindingMapIterator iterator = bindingMap.ForwardIterator();
            while (iterator.MoveNext())
            {
                Definition def = iterator.Key;
                Binding binding = iterator.Current as Binding;

                if (def != null)
                {
                    ForgeTypeId typeId = def.GetDataType();
                    
                    ParameterData paramData = new ParameterData
                    {
                        Name = def.Name,
                        ParameterTypeId = typeId,
                        ParameterType = ParameterTypeHelper.GetTypeName(typeId),
                        IsInstance = binding is InstanceBinding,
#if REVIT2025
                        Group = (def as InternalDefinition)?.GetGroupTypeId() ?? new ForgeTypeId()
#else
                        Group = (def as InternalDefinition)?.ParameterGroup ?? BuiltInParameterGroup.INVALID
#endif
                    };

                    // Get bound categories
                    CategorySet catSet = null;
                    if (binding is InstanceBinding instBinding)
                        catSet = instBinding.Categories;
                    else if (binding is TypeBinding typeBinding)
                        catSet = typeBinding.Categories;
                    
                    paramData.Categories = new List<string>();
                    if (catSet != null)
                    {
                        foreach (Category cat in catSet)
                        {
                            paramData.Categories.Add(cat.Name);
                        }
                    }

                    parameters.Add(paramData);
                }
            }

            return parameters;
        }

        /// <summary>
        /// Gets parameter value as string
        /// </summary>
        public static string GetParameterValueAsString(Parameter param)
        {
            if (param == null || !param.HasValue)
                return string.Empty;

            switch (param.StorageType)
            {
                case StorageType.String:
                    return param.AsString() ?? string.Empty;
                case StorageType.Integer:
                    return param.AsInteger().ToString();
                case StorageType.Double:
                    return param.AsDouble().ToString();
                case StorageType.ElementId:
#if REVIT2023
                    return param.AsElementId().IntegerValue.ToString();
#else
                    return param.AsElementId().Value.ToString();
#endif
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Sets parameter value from string
        /// </summary>
        public static bool SetParameterValue(Parameter param, string value)
        {
            if (param == null || param.IsReadOnly)
                return false;

            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        param.Set(value);
                        return true;
                    case StorageType.Integer:
                        if (int.TryParse(value, out int intValue))
                        {
                            param.Set(intValue);
                            return true;
                        }
                        break;
                    case StorageType.Double:
                        if (double.TryParse(value, out double doubleValue))
                        {
                            param.Set(doubleValue);
                            return true;
                        }
                        break;
                    case StorageType.ElementId:
                        if (int.TryParse(value, out int idValue))
                        {
#if REVIT2023
                            param.Set(new ElementId(idValue));
#else
                            param.Set(new ElementId((long)idValue));
#endif
                            return true;
                        }
                        break;
                }
            }
            catch { }

            return false;
        }
    }
}
