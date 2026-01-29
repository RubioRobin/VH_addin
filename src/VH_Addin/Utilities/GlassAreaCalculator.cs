using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using static VH_Tools.Models.Constants;
using static VH_Tools.Utilities.CalculatorHelpers;
using VH_Tools.Models;

namespace VH_Tools.Utilities
{
    public class GlassAreaCalculator
    {
        private readonly Document _doc;
        private double _defaultSashWidthFt;

        public GlassAreaCalculator(Document doc, double defaultSashWidthMm)
        {
            _doc = doc;
            _defaultSashWidthFt = defaultSashWidthMm * MM_TO_FT;
        }

        public double Calculate(FamilyInstance window, BoundingBoxXYZ bbox, XYZ facing, double mmCut = 600.0)
        {
            // 1) Get base dimensions
            var (wName, hName) = AutodetectDims(window);
            if (wName == null || hName == null)
                return 0.0;

            double? wParam = GetDoubleParam(window, wName);
            double? hParam = GetDoubleParam(window, hName);
            
            if (!wParam.HasValue || !hParam.HasValue)
                return 0.0;

            // 2) Detect aluminum vs wood (Check both instance and type)
            bool hasAlu = GetDoubleParam(window, AL_SIDE).HasValue ||
                          GetDoubleParam(window, AL_TOP_BOT).HasValue ||
                          GetDoubleParam(window, AL_OFF_SIDE).HasValue ||
                          GetDoubleParam(window, AL_EXTRA_UNDER).HasValue;

            int nv, nh;
            double tSide, tTop, tBot, tv, th;

            double netW, netH;

            if (hasAlu)
            {
                // ALUMINUM logic
                double offSide = GetDoubleParam(window, AL_OFF_SIDE) ?? 0;
                double extraUnder = GetDoubleParam(window, AL_EXTRA_UNDER) ?? 0;
                double aSide = GetDoubleParam(window, AL_SIDE) ?? 0;
                double aTopBot = GetDoubleParam(window, AL_TOP_BOT) ?? 0;
                double aMullV = GetDoubleParam(window, AL_MULL_V) ?? aSide;
                double aMullH = GetDoubleParam(window, AL_MULL_H) ?? aTopBot;

                nv = Math.Max(0, GetIntParam(window, P_MULL_V_N));
                nh = Math.Max(0, GetIntParam(window, P_MULL_H_N));

                double BEff = wParam.Value;
                double HEff = hParam.Value;

                tSide = aSide;
                tTop = aTopBot;
                tBot = GetDoubleParam(window, AL_VIEW_SILL) ?? aTopBot;
                tv = aMullV;
                th = aMullH;

                netW = Math.Max(0, BEff - 2.0 * tSide - nv * tv);
                netH = Math.Max(0, HEff - 2.0 * tTop - nh * th - extraUnder);
            }
            else
            {
                // WOOD logic
                double? tTopVal = GetDoubleParam(window, P_TOP);
                double? tSideVal = GetDoubleParam(window, P_SIDE_G);
                double? tBotVal = GetDoubleParam(window, P_BOT_G);
                
                if (!tTopVal.HasValue || !tSideVal.HasValue || !tBotVal.HasValue)
                    return 0.0;

                tTop = tTopVal.Value;
                tSide = tSideVal.Value;
                tBot = tBotVal.Value;

                tv = GetDoubleParam(window, P_MULL_V_T) ?? 0;
                th = GetDoubleParam(window, P_MULL_H_T) ?? 0;
                nv = Math.Max(0, GetIntParam(window, P_MULL_V_N));
                nh = Math.Max(0, GetIntParam(window, P_MULL_H_N));

                double wEff = wParam.Value;
                double hEff = hParam.Value;

                double extraHFt = (17.0 * (1 + nh)) * MM_TO_FT;

                netW = Math.Max(0, wEff - 2.0 * tSide - nv * tv);
                netH = Math.Max(0, hEff - (tTop + tBot + nh * th) - extraHFt);
            }

            if (netW <= 0 || netH <= 0)
                return 0;

            // 4) Grid division
            int cols = Math.Max(1, nv + 1);
            int rows = Math.Max(1, nh + 1);

            var colSizes = GetAxisSizes(netW, cols, window, "cols");
            var rowSizes = GetAxisSizes(netH, rows, window, "rows");

            // 5) Cut at 600mm above level
            double mmCutFt = mmCut * MM_TO_FT;
            double sillFt = GetSillHeight(window);
            var rowCutUsed = new double[rowSizes.Count];

            double glassBottomFt;
            if (hasAlu)
            {
                double offSide = GetDoubleParam(window, AL_OFF_SIDE) ?? 0;
                double extraUnder = GetDoubleParam(window, AL_EXTRA_UNDER) ?? 0;
                double aTopBot = GetDoubleParam(window, AL_TOP_BOT) ?? 0;
                glassBottomFt = sillFt + Math.Max(0, aTopBot - offSide) + extraUnder;
            }
            else
            {
                glassBottomFt = sillFt + tBot + (17.0 * MM_TO_FT);
            }

            double cutFromBottomFt = Math.Max(0, mmCutFt - glassBottomFt);

            if (cutFromBottomFt > 0 && rowSizes.Sum() > 0)
            {
                double rem = Math.Min(cutFromBottomFt, rowSizes.Sum());
                int i = 0;
                while (rem > 0 && i < rowSizes.Count)
                {
                    double take = Math.Min(rowSizes[i], rem);
                    rowSizes[i] -= take;
                    rowCutUsed[i] += take;
                    rem -= take;
                    i++;
                }
                rowSizes = rowSizes.Select(x => Math.Max(0, x)).ToList();
            }

            // 6) Vertical Obstruction (Top Cut)
            double topCutFt = 0;
            try 
            {
                XYZ windowCenter = (bbox.Min + bbox.Max) * 0.5;
                double glassTopZ = bbox.Max.Z - tTop;
                
                View3D view3d = new FilteredElementCollector(_doc)
                    .OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate && (v.Name == "{3D}" || v.Name.Contains("3D")));
                
                if (view3d == null)
                {
                    view3d = _doc.ActiveView as View3D ?? new FilteredElementCollector(_doc)
                        .OfClass(typeof(View3D))
                        .Cast<View3D>()
                        .FirstOrDefault(v => !v.IsTemplate);
                }

                var categories = new List<BuiltInCategory> 
                { 
                    BuiltInCategory.OST_Walls, 
                    BuiltInCategory.OST_Floors, 
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_StructuralFraming,
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_Roofs,
                    BuiltInCategory.OST_Ceilings,
                    BuiltInCategory.OST_StructuralFoundation
                };
                ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(categories);
                var excludeIds = new List<ElementId> { window.Id };
                excludeIds.AddRange(window.GetSubComponentIds());
                if (window.Host != null) excludeIds.Add(window.Host.Id);
                
                double maxCutVal = 0;

                if (view3d != null) 
                {
                    ReferenceIntersector intersector = new ReferenceIntersector(catFilter, FindReferenceTarget.All, view3d);
                    intersector.FindReferencesInRevitLinks = true;

                    XYZ sideDir = facing.CrossProduct(XYZ.BasisZ).Normalize();
                    double winWidth = bbox.Max.DistanceTo(new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z)) * 0.8;
                    
                    double[] widthSteps = { -0.5, -0.25, 0.0, 0.25, 0.5 }; 
                    double[] thickOffsets = { -1.0, -0.2, 0.2, 1.0, 2.5 };

                    foreach (double wStep in widthSteps)
                    {
                        XYZ basePoint = windowCenter + sideDir * (wStep * winWidth);
                        foreach (double tOffset in thickOffsets)
                        {
                            XYZ rayOrigin = new XYZ(basePoint.X, basePoint.Y, bbox.Min.Z + 0.1) + facing * tOffset;
                            ReferenceWithContext refCtx = intersector.FindNearest(rayOrigin, XYZ.BasisZ);
                            
                            if (refCtx != null && refCtx.Proximity > 0)
                            {
                                ElementId refId = refCtx.GetReference().ElementId;
                                if (excludeIds.Contains(refId)) continue;

                                double obstructionZ = rayOrigin.Z + refCtx.Proximity;
                                if (obstructionZ < glassTopZ - 0.001)
                                {
                                    maxCutVal = Math.Max(maxCutVal, glassTopZ - obstructionZ);
                                }
                            }
                        }
                    }
                }

                if (maxCutVal <= 0)
                {
                    var potentials = new FilteredElementCollector(_doc)
                        .WherePasses(catFilter)
                        .WhereElementIsNotElementType()
                        .Excluding(excludeIds)
                        .ToElements();

                    foreach (var opt in potentials)
                    {
                        BoundingBoxXYZ obbox = opt.get_BoundingBox(null);
                        if (obbox == null) continue;

                        if (obbox.Min.X < bbox.Max.X && obbox.Max.X > bbox.Min.X &&
                            obbox.Min.Y < bbox.Max.Y && obbox.Max.Y > bbox.Min.Y)
                        {
                            if (obbox.Min.Z > bbox.Min.Z && obbox.Min.Z < glassTopZ - 0.001)
                            {
                                maxCutVal = Math.Max(maxCutVal, glassTopZ - obbox.Min.Z);
                            }
                        }
                    }
                }

                topCutFt = maxCutVal;
            }
            catch { }

            if (topCutFt > 0 && rowSizes.Sum() > 0)
            {
                double remTop = Math.Min(topCutFt, rowSizes.Sum());
                int i = rowSizes.Count - 1;
                while (remTop > 0 && i >= 0)
                {
                    double take = Math.Min(rowSizes[i], remTop);
                    rowSizes[i] -= take;
                    remTop -= take;
                    i--;
                }
                rowSizes = rowSizes.Select(x => Math.Max(0, x)).ToList();
            }

            // 7) Get fill types
            var fills = CollectVlakvullingen(window);

            // 8) Sash width
            double sash = GetSashWidthFromNested(window)
                       ?? GetDoubleParam(window, P_SIDE_G) 
                       ?? GetDoubleParam(window, AL_SIDE) 
                       ?? 0;
            sash = Math.Max(0, sash);

            double totalFt2 = 0;

            for (int rIdx = 0; rIdx < rowSizes.Count; rIdx++)
            {
                double ph = rowSizes[rIdx];
                for (int cIdx = 0; cIdx < colSizes.Count; cIdx++)
                {
                    double pw = colSizes[cIdx];
                    string label = fills.ContainsKey((cIdx + 1, rIdx + 1))
                        ? fills[(cIdx + 1, rIdx + 1)]
                        : "";
                    
                    if (IsPanel(label))
                        continue;

                    bool oper = IsOperabel(label);

                    double gw = Math.Max(0, pw - (oper ? 2.0 * sash : 0));
                    double gh;

                    if (oper)
                    {
                        double bottomSashRem = Math.Max(0, sash - rowCutUsed[rIdx]);
                        gh = Math.Max(0, ph - sash - bottomSashRem);
                    }
                    else
                    {
                        gh = ph;
                    }

                    totalFt2 += gw * gh;
                }
            }

            return totalFt2;
        }

        private (string wName, string hName) AutodetectDims(Element elem)
        {
            var names = new HashSet<string>();
            
            foreach (Parameter p in elem.Parameters)
            {
                if (p.Definition != null)
                    names.Add(p.Definition.Name);
            }

            ElementType type = _doc.GetElement(elem.GetTypeId()) as ElementType;
            if (type != null)
            {
                foreach (Parameter p in type.Parameters)
                {
                    if (p.Definition != null)
                        names.Add(p.Definition.Name);
                }
            }

            string w = W_NAMES.FirstOrDefault(n => names.Contains(n));
            string h = H_NAMES.FirstOrDefault(n => names.Contains(n));

            return (w, h);
        }

        private List<double> GetAxisSizes(double netLenFt, int count, Element elem, string axis)
        {
            if (count <= 0 || netLenFt <= 0)
                return Enumerable.Repeat(0.0, Math.Max(0, count)).ToList();

            string[] listParams = axis == "cols"
                ? new[] { "VH_verdeling_kolommen", "verdeling_kolommen", "kolom_breedtes", "kolombreedtes_mm" }
                : new[] { "VH_verdeling_rijen", "verdeling_rijen", "rij_hoogtes", "rijhoogtes_mm" };

            var lst = ReadDistList(elem, listParams);
            if (lst != null && lst.Count > 0)
            {
                double s = lst.Sum();
                if (s > 0)
                {
                    if (s > count * 1.5)
                    {
                        var partsFt = lst.Select(x => x * MM_TO_FT).ToList();
                        double d = Math.Max(1e-9, partsFt.Sum());
                        return Pad(partsFt.Select(x => x * netLenFt / d).ToList(), count);
                    }
                    else
                    {
                        double d = Math.Max(1e-9, s);
                        return Pad(lst.Select(x => x * netLenFt / d).ToList(), count);
                    }
                }
            }

            double part = netLenFt / Math.Max(1, count);
            return Enumerable.Repeat(part, count).ToList();
        }

        private List<double> ReadDistList(Element elem, string[] paramNames)
        {
            foreach (string name in paramNames)
            {
                Parameter p = elem.LookupParameter(name);
                if (p == null)
                {
                    ElementType type = _doc.GetElement(elem.GetTypeId()) as ElementType;
                    if (type != null)
                        p = type.LookupParameter(name);
                }

                if (p != null)
                {
                    string val = "";
                    try
                    {
                        val = p.AsString() ?? p.AsValueString() ?? "";
                    }
                    catch
                    {
                        try { val = p.AsString() ?? ""; } catch { }
                    }

                    var lst = TryParseNumberList(val);
                    if (lst != null && lst.Count > 0)
                        return lst;
                }
            }
            return null;
        }

        private List<double> TryParseNumberList(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return null;

            var raw = s.Replace('|', ';').Replace(',', ';').Split(';')
                .Select(t => t.Trim())
                .Where(t => !string.IsNullOrEmpty(t))
                .ToList();

            try
            {
                return raw.Select(x => double.Parse(x.Replace(',', '.'))).ToList();
            }
            catch
            {
                return null;
            }
        }

        private List<double> Pad(List<double> lst, int count)
        {
            if (lst.Count < count)
            {
                lst.AddRange(Enumerable.Repeat(0.0, count - lst.Count));
            }
            return lst.Take(count).ToList();
        }

        private Dictionary<(int col, int row), string> CollectVlakvullingen(Element elem)
        {
            var result = new Dictionary<(int, int), string>();

            void FromContainer(Element obj)
            {
                foreach (Parameter p in obj.Parameters)
                {
                    if (p.Definition == null)
                        continue;

                    string nm = p.Definition.Name;
                    if (!nm.ToLower().StartsWith("vh_vlakvulling_"))
                        continue;

                    string txt = "";
                    try
                    {
                        ElementId eid = p.AsElementId();
                        if (eid != null && eid.IntegerValue > 0)
                        {
                            Element tEl = _doc.GetElement(eid);
                            if (tEl != null)
                            {
                                string fam = "";
                                string typ = "";
                                if (tEl is FamilySymbol fs && fs.Family != null)
                                    fam = fs.Family.Name;
                                typ = tEl.Name ?? "";
                                txt = $"{fam} : {typ}".Trim(' ', ':');
                            }
                        }
                    }
                    catch { }

                    if (string.IsNullOrEmpty(txt))
                    {
                        try { txt = p.AsString() ?? ""; } catch { }
                    }

                    try
                    {
                        string suf = nm.Split(new[] { '_' }, 3).LastOrDefault()?.Trim() ?? "";
                        if (suf.Length >= 2)
                        {
                            char cChar = suf[0];
                            string rStr = new string(suf.Skip(1).TakeWhile(char.IsDigit).ToArray());
                            
                            int c = char.ToUpper(cChar) - 'A' + 1;
                            if (int.TryParse(rStr, out int r))
                            {
                                result[(c, r)] = txt.Trim().ToLower();
                            }
                        }
                    }
                    catch { }
                }
            }

            FromContainer(elem);
            ElementType type = _doc.GetElement(elem.GetTypeId()) as ElementType;
            if (type != null)
                FromContainer(type);

            return result;
        }

        private bool IsPanel(string textLower)
        {
            if (string.IsNullOrWhiteSpace(textLower)) return false;
            string t = textLower.ToLower();
            string[] tokens = { "paneel", "panel", "dicht", "bord", "plaat" };
            return tokens.Any(tok => t.Contains(tok));
        }

        private bool IsOperabel(string textLower)
        {
            if (string.IsNullOrWhiteSpace(textLower))
                return false;

            string t = textLower.ToLower();
            string[] tokens = { "raam", "draai", "kiep", "val", "opend", "openend", "open", "schuif", "dk", " op " };
            return tokens.Any(tok => t.Contains(tok));
        }

        private double? GetSashWidthFromNested(FamilyInstance window)
        {
            double maxSashFt = 0;

            var subIds = window.GetSubComponentIds();
            if (subIds != null && subIds.Count > 0)
            {
                foreach (ElementId id in subIds)
                {
                    Element sub = _doc.GetElement(id);
                    if (sub == null) continue;

                    double? val = GetDoubleParam(sub, "glas_offset") 
                                ?? GetDoubleParam(sub, "Glas_offset")
                                ?? GetDoubleParam(sub, "Glas_Offset");
                    
                    if (val.HasValue && val.Value > maxSashFt)
                    {
                        maxSashFt = val.Value;
                    }
                }
            }

            void CheckParameters(Element elem)
            {
                if (elem == null) return;
                
                IList<Parameter> paramsToScan;
                try { paramsToScan = elem.GetOrderedParameters(); }
                catch { paramsToScan = elem.Parameters.Cast<Parameter>().Cast<Parameter>().ToList(); }

                foreach (Parameter p in paramsToScan)
                {
                    if (p.Definition == null) continue;
                    if (p.StorageType != StorageType.ElementId) continue;

                    ElementId eid = p.AsElementId();
                    if (eid == null || eid == ElementId.InvalidElementId) continue;

                    Element sym = _doc.GetElement(eid);
                    if (sym == null) continue;

                    bool isLikelyFilling = p.Definition.Name.ToLower().Contains("vlakvulling") || 
                                           sym.Name.ToLower().Contains("vlakvulling");

                    if (isLikelyFilling)
                    {
                        IList<Parameter> symParams;
                        try { symParams = sym.GetOrderedParameters(); }
                        catch { symParams = sym.Parameters.Cast<Parameter>().Cast<Parameter>().ToList(); }

                        foreach (Parameter sp in symParams)
                        {
                            if (sp.Definition == null) continue;
                            string pName = sp.Definition.Name.ToLower();
                            
                            if (pName == "glas_offset" || pName.Contains("glas_offset"))
                            {
                                if (sp.StorageType == StorageType.Double)
                                {
                                    double val = sp.AsDouble(); 
                                    if (val > maxSashFt) maxSashFt = val;
                                }
                            }
                        }
                        
                        double? explicitVal = GetDoubleParam(sym, "glas_offset");
                        if (explicitVal.HasValue && explicitVal.Value > maxSashFt)
                            maxSashFt = explicitVal.Value;
                    }
                }
            }

            CheckParameters(window);
            ElementType type = _doc.GetElement(window.GetTypeId()) as ElementType;
            CheckParameters(type);

            if (maxSashFt <= 0.001)
            {
                void CheckNames(Element elem)
                {
                    if (elem == null) return;
                    IList<Parameter> paramsToScan;
                    try { paramsToScan = elem.GetOrderedParameters(); }
                    catch { paramsToScan = elem.Parameters.Cast<Parameter>().Cast<Parameter>().ToList(); }

                    foreach (Parameter p in paramsToScan)
                    {
                        if (p.StorageType != StorageType.ElementId) continue;
                        ElementId eid = p.AsElementId();
                        if (eid == null || eid == ElementId.InvalidElementId) continue;
                        Element sym = _doc.GetElement(eid);
                        if (sym == null) continue;

                        string sName = sym.Name.ToLower();
                        bool isOperable = sName.Contains("draaikiep") || 
                                          sName.Contains("valraam") || 
                                          sName.Contains("uitzet") ||
                                          sName.Contains("stolp") ||
                                          (sName.Contains("draai") && !sName.Contains("draaiend_deel"));

                        if (isOperable)
                        {
                            if (_defaultSashWidthFt > maxSashFt) maxSashFt = _defaultSashWidthFt;
                        }
                    }
                }
                
                CheckNames(window);
                CheckNames(type);
            }

            return maxSashFt > 0.001 ? maxSashFt : (double?)null;
        }
    }
}
