using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace ParaManager.Helpers
{
    /// <summary>
    /// Provides a unified interface for parameter type operations across Revit 2023, 2024, and 2025
    /// </summary>
    public static class ParameterTypeHelper
    {
        private static readonly Dictionary<string, ForgeTypeId> _typeMap = new Dictionary<string, ForgeTypeId>
        {
            { "Text", SpecTypeId.String.Text },
            { "Integer", SpecTypeId.Int.Integer },
            { "Number", SpecTypeId.Number },
            { "Length", SpecTypeId.Length },
            { "Area", SpecTypeId.Area },
            { "Volume", SpecTypeId.Volume },
            { "Angle", SpecTypeId.Angle },
            { "YesNo", SpecTypeId.Boolean.YesNo },
            { "URL", SpecTypeId.String.Url },
            { "Material", SpecTypeId.Reference.Material },
            { "MultilineText", SpecTypeId.String.MultilineText }
        };

        /// <summary>
        /// Gets ForgeTypeId from a string type name
        /// </summary>
        public static ForgeTypeId GetForgeTypeId(string typeName)
        {
            if (_typeMap.TryGetValue(typeName, out ForgeTypeId typeId))
                return typeId;

            // Default to Text if unknown
            return SpecTypeId.String.Text;
        }

        /// <summary>
        /// Gets a display-friendly type name from ForgeTypeId
        /// </summary>
        public static string GetTypeName(ForgeTypeId typeId)
        {
            if (typeId == null)
                return "Unknown";

            foreach (var kvp in _typeMap)
            {
                if (kvp.Value.Equals(typeId))
                    return kvp.Key;
            }

            // Return the TypeId as fallback
            return typeId.TypeId ?? "Unknown";
        }

        /// <summary>
        /// Gets all available parameter type names
        /// </summary>
        public static List<string> GetAllTypeNames()
        {
            return new List<string>(_typeMap.Keys);
        }

        /// <summary>
        /// Gets the data type (ForgeTypeId) from a Definition
        /// </summary>
        public static ForgeTypeId GetDataType(Definition def)
        {
            return def.GetDataType();
        }

        /// <summary>
        /// Gets a display string for a parameter's type
        /// </summary>
        public static string GetParameterTypeString(Definition def)
        {
            return GetTypeName(def.GetDataType());
        }
    }
}
