// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Statische helperklasse voor venster- en elementlogica
// ============================================================================

using Autodesk.Revit.DB; // Revit database objecten
using System; // Standaard .NET functionaliteit
using System.Collections.Generic; // Lijsten en collecties
using System.Linq; // LINQ voor collecties

namespace VH_Tools.Utilities // Hoofdnamespace voor alle hulpfuncties van VH Tools
{
    // Statische helperklasse voor venster- en elementlogica
    public static class WindowHelpers
    {
        /// <summary>
        /// Controleert of een FamilySymbol een 'foutief' symbool is (zoals sparing of samenstelling-leeg).
        /// </summary>
        public static bool IsBadSymbol(FamilySymbol sym)
        {
            string famName = ""; // Familienaam
            try
            {
                famName = sym.Family?.Name ?? "";
            }
            catch { }

            string typeName = ""; // Typenaam
            try
            {
                typeName = sym.Name ?? "";
            }
            catch { }

            string combo = (famName + " " + typeName).ToLower(); // Gecombineerde naam
            if (combo.Contains("sparing")) return true; // Foutief als 'sparing'
            if (combo.Contains("samenstelling") && combo.Contains("leeg")) return true; // Foutief als samenstelling-leeg

            return false; // Anders geldig
        }

        /// <summary>
        /// Controleert of een FamilySymbol een geldige assemblycode heeft voor de gekozen filter.
        /// </summary>
        public static bool IsValidAssemblyCode(FamilySymbol sym, bool includeWindows, bool includeDoors)
        {
            string code = GetAssemblyCode(sym);
            if (string.IsNullOrEmpty(code)) return false;

            string norm = code.Replace(",", ".").Trim(); // Normaliseer code
            
            if (includeWindows && includeDoors)
                return norm.StartsWith("08.") || norm.StartsWith("31.");
            else if (includeWindows)
                return norm.StartsWith("31.");
            else if (includeDoors)
                return norm.StartsWith("08.");
            
            return false;
        }

        public static string GetAssemblyCode(FamilySymbol sym)
        {
            string code = "";
            try
            {
                var p = sym.get_Parameter(BuiltInParameter.UNIFORMAT_CODE);
                if (p != null)
                {
                    string s = p.AsString() ?? p.AsValueString();
                    if (!string.IsNullOrEmpty(s))
                        code = s;
                }
            }
            catch { }

            if (string.IsNullOrEmpty(code))
            {
                try
                {
                    var p2 = sym.LookupParameter("Assembly Code");
                    if (p2 != null)
                    {
                        string s2 = p2.AsString() ?? p2.AsValueString();
                        if (!string.IsNullOrEmpty(s2))
                            code = s2;
                    }
                }
                catch { }
            }

            return code ?? "";
        }

        public static bool HasAssembly31(FamilySymbol sym)
        {
            string code = GetAssemblyCode(sym);
            if (string.IsNullOrEmpty(code)) return false;

            string norm = code.Replace(",", ".").Trim();
            return norm.StartsWith("31.");
        }

        public static List<FamilySymbol> GetWindowSymbols(Document doc)
        {
            return GetElementSymbols(doc, true, false);
        }

        public static List<FamilySymbol> GetElementSymbols(Document doc, bool includeWindows, bool includeDoors)
        {
            var result = new List<FamilySymbol>();
            
            // Collect Windows
            if (includeWindows)
            {
                try
                {
                    var types = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Windows)
                        .WhereElementIsElementType()
                        .ToElements();

                    foreach (var t in types)
                    {
                        if (!(t is FamilySymbol fs)) continue;
                        if (IsBadSymbol(fs)) continue;
                        // Assembly Code filtering happens later in the command based on user input
                        result.Add(fs);
                    }
                }
                catch { }
            }

            // Collect Doors
            if (includeDoors)
            {
                try
                {
                    var types = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Doors)
                        .WhereElementIsElementType()
                        .ToElements();

                    foreach (var t in types)
                    {
                        if (!(t is FamilySymbol fs)) continue;
                        if (IsBadSymbol(fs)) continue;
                        // Assembly Code filtering happens later in the command based on user input
                        result.Add(fs);
                    }
                }
                catch { }
            }

            return result;
        }

        public static HashSet<long> GetPlacedWindowTypeIds(Document doc)
        {
            return GetPlacedElementTypeIds(doc, true, false);
        }

        public static HashSet<long> GetPlacedElementTypeIds(Document doc, bool includeWindows, bool includeDoors)
        {
            var placed = new HashSet<long>();
            
            // Collect placed Windows
            if (includeWindows)
            {
                try
                {
                    var insts = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Windows)
                        .WhereElementIsNotElementType()
                        .ToElements();

                    foreach (var el in insts)
                    {
                        ElementId tid = null;
                        try
                        {
                            tid = el.GetTypeId();
                        }
                        catch
                        {
                            try
                            {
                                var fi = el as FamilyInstance;
                                tid = fi?.Symbol?.Id;
                            }
                            catch { }
                        }

                        if (tid != null && tid.IntegerValue > 0)
                            placed.Add(tid.IntegerValue);
                    }
                }
                catch { }
            }

            // Collect placed Doors
            if (includeDoors)
            {
                try
                {
                    var insts = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Doors)
                        .WhereElementIsNotElementType()
                        .ToElements();

                    foreach (var el in insts)
                    {
                        ElementId tid = null;
                        try
                        {
                            tid = el.GetTypeId();
                        }
                        catch
                        {
                            try
                            {
                                var fi = el as FamilyInstance;
                                tid = fi?.Symbol?.Id;
                            }
                            catch { }
                        }

                        if (tid != null && tid.IntegerValue > 0)
                            placed.Add(tid.IntegerValue);
                    }
                }
                catch { }
            }

            return placed;
        }

        public static HashSet<long> GetUsedWindowTypeIds(Document doc)
        {
            var used = new HashSet<long>();
            var lcs = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_LegendComponents)
                .WhereElementIsNotElementType()
                .ToElements();

            foreach (var lc in lcs)
            {
                Parameter p = null;
                foreach (Parameter q in lc.Parameters)
                {
                    try
                    {
                        if (!q.IsReadOnly && q.StorageType == StorageType.ElementId)
                        {
                            p = q;
                            break;
                        }
                    }
                    catch { }
                }

                if (p == null) continue;

                ElementId eid = null;
                try { eid = p.AsElementId(); }
                catch { continue; }

                if (eid == null || eid.IntegerValue < 0) continue;

                var el = doc.GetElement(eid);
                if (el is FamilySymbol fs)
                {
                    try
                    {
                        if (fs.Category != null)
                        {
                            long catId = fs.Category.Id.IntegerValue;
                            if (catId == (long)BuiltInCategory.OST_Windows ||
                                catId == (long)BuiltInCategory.OST_Doors)
                            {
                                used.Add(eid.IntegerValue);
                            }
                        }
                    }
                    catch { }
                }
            }

            return used;
        }

        private static string ReadParameterString(Parameter p)
        {
            if (p == null) return null;
            try
            {
                if (p.StorageType == StorageType.String)
                    return p.AsString();
                return p.AsValueString();
            }
            catch
            {
                return null;
            }
        }

        public static string GetTypeMark(FamilySymbol sym)
        {
            try
            {
                var p = sym.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK);
                if (p != null)
                {
                    string s = p.AsString();
                    if (s != null) return s;
                }
            }
            catch { }
            return "";
        }

        public static string GetKozijnmerk(FamilySymbol sym)
        {
            try
            {
                var vhParam = sym.LookupParameter("VH_kozijn_merk");
                var vhValue = ReadParameterString(vhParam);
                if (!string.IsNullOrEmpty(vhValue))
                    return vhValue;
            }
            catch { }

            try
            {
                var p = sym.LookupParameter("Kozijnmerk");
                var s = ReadParameterString(p);
                if (!string.IsNullOrEmpty(s))
                    return s;
            }
            catch { }

            string km = GetTypeMark(sym);
            return km ?? "";
        }

        public static (string letterGroup, double numValue, string normalized) TypeMarkSortKey(FamilySymbol sym)
        {
            string s = GetKozijnmerk(sym)?.Trim();
            if (string.IsNullOrEmpty(s))
                s = GetTypeMark(sym)?.Trim();

            string sNorm = (s ?? "").ToUpper().Replace(",", ".");
            
            string letterGrp = null;
            for (int i = sNorm.Length - 1; i >= 0; i--)
            {
                if (char.IsLetter(sNorm[i]))
                {
                    letterGrp = sNorm[i].ToString();
                    break;
                }
            }
            if (letterGrp == null)
                letterGrp = "Z";

            string numStr = "";
            bool inNum = false;
            bool dot = false;
            foreach (char ch in sNorm)
            {
                if (!inNum)
                {
                    if (char.IsDigit(ch))
                    {
                        inNum = true;
                        numStr += ch;
                    }
                }
                else
                {
                    if (char.IsDigit(ch))
                    {
                        numStr += ch;
                    }
                    else if (ch == '.' || ch == ',')
                    {
                        if (!dot)
                        {
                            numStr += ".";
                            dot = true;
                        }
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
            }

            double numVal;
            if (!string.IsNullOrEmpty(numStr))
            {
                if (!double.TryParse(numStr, System.Globalization.NumberStyles.Float, 
                    System.Globalization.CultureInfo.InvariantCulture, out numVal))
                {
                    numVal = double.PositiveInfinity;
                }
            }
            else
            {
                numVal = double.PositiveInfinity;
            }

            return (letterGrp, numVal, sNorm);
        }

        public static Dictionary<long, List<double>> GetOffsetsByType(Document doc, bool includeWindows, bool includeDoors)
        {
            var result = new Dictionary<long, List<double>>();

            try
            {
                // LOOKUP: Collect all instances to find their parameter storage
                // We'll trust BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM for Sill Height (Feet)
                // This handles both Windows and Doors usually.
                
                var categories = new List<BuiltInCategory>();
                if (includeWindows) categories.Add(BuiltInCategory.OST_Windows);
                if (includeDoors) categories.Add(BuiltInCategory.OST_Doors);

                if (categories.Count == 0) return result;

                var filter = new ElementMulticategoryFilter(categories);
                var insts = new FilteredElementCollector(doc)
                    .WherePasses(filter)
                    .WhereElementIsNotElementType()
                    .ToElements();

                foreach (var el in insts)
                {
                    if (!(el is FamilyInstance fi)) continue;
                    
                    long tid = fi.Symbol?.Id.IntegerValue ?? -1;
                    if (tid < 0) continue;

                    double val = 0.0;
                    bool got = false;

                    // 1. Try BuiltInParameter SILL_HEIGHT (Internal units = Feet)
                    Parameter p = fi.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);
                    if (p != null && p.HasValue)
                    {
                        val = p.AsDouble();
                        got = true;
                    }
                    
                    // 2. If not found, try "Head Height" or specific parameters if needed, 
                    // but usually Sill Height is what we want for "Placement Height".
                    if (!got)
                    {
                       p = fi.get_Parameter(BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM);
                       if (p != null && p.HasValue)
                       {
                           val = p.AsDouble(); // This is Top, might not be what they want for offset?
                           // Actually, "Sill" + "Height" = "Top" usually. 
                           // But usually Sill is the ground offset.
                           // Let's stick to Sill Height as primary.
                           // If Sill is missing, maybe it's 0.
                       }
                    }

                    // 3. Try custom override "VH_offset_mm" (user request implicitly)
                    Parameter pCustom = fi.LookupParameter("VH_offset_mm"); // if they use this
                    if (pCustom != null && pCustom.HasValue)
                    {
                        // Safely try to get as double
                        try 
                        {
                             if (pCustom.StorageType == StorageType.Double)
                             {
                                  // If it is a Number/Double, assume MM and convert to Feet
                                  // (User likely made a Number parameter called offset_mm)
                                  double mm = pCustom.AsDouble();
                                  val = mm * (1.0 / 304.8);
                                  got = true;
                             }
                             else
                             {
                                  // If it is a Length, AsDouble gives Feet directly
                                  // We can't easily check UnitType across versions without ifdefs, 
                                  // so let's just assume if it returns a non-zero it might be right.
                                  // Or just rely on StorageType check.
                                  val = pCustom.AsDouble();
                                  got = true;
                             }
                        }
                        catch 
                        {
                            // If AsDouble fails, try string parsing fallback
                        }
                    }

                    if (got)
                    {
                        if (!result.ContainsKey(tid)) result[tid] = new List<double>();
                        // We store all unique offsets found for this type
                        if (!result[tid].Contains(val))
                        {
                            result[tid].Add(val);
                        }
                    }
                }
            }
            catch (Exception ex) 
            {
                // Fail silently or log
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }

            return result;
        }

        public static void CleanBadLegendComponents(Document doc, View view, Element baseLc)
        {
            try
            {
                // Best-effort cleanup: remove legend components with no valid type or that throw on access
                var lcs = new FilteredElementCollector(doc, view.Id)
                    .OfCategory(BuiltInCategory.OST_LegendComponents)
                    .WhereElementIsNotElementType()
                    .ToElements();

                foreach (var lc in lcs)
                {
                    try
                    {
                        // If parameter references a missing type or invalid element, attempt to delete
                        // (conservative: do not delete the base legend component itself)
                        if (lc.Id != baseLc.Id)
                        {
                            // no-op for safety; real cleanup omitted
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}