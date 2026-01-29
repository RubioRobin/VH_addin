// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Statische helperklasse met diverse Revit-hulpmethoden
// ============================================================================

using Autodesk.Revit.DB; // Revit database objecten
using System; // Standaard .NET functionaliteit
using System.Collections.Generic; // Lijsten en collecties
using System.Linq; // LINQ voor collecties

namespace VH_Tools.Utilities // Hoofdnamespace voor alle hulpfuncties van VH Tools
{
    // Statische helperklasse met diverse Revit-hulpmethoden
    public static class RevitHelpers
    {
        public const double FT_PER_MM = 1.0 / 304.8; // Omrekenfactor voet naar mm
        public const double MIN_SLOT_MM = 1400.0; // Minimale sleufmaat in mm
        public const double MIN_ROW_DY_MM = 1400.0; // Minimale rijhoogte in mm

        /// <summary>
        /// Geeft een veilige label-string terug voor een Revit-element (family/type).
        /// </summary>
        public static string SafeLabel(Element obj)
        {
            string famName = null; // Familienaam
            string typName = null; // Typenaam

            try
            {
                if (obj is FamilySymbol fs)
                {
                    famName = fs.Family?.Name;
                    typName = fs.Name;
                }
                else if (obj is ElementType et)
                {
                    typName = et.Name;
                }
            }
            catch { }

            if (!string.IsNullOrEmpty(famName) && !string.IsNullOrEmpty(typName))
                return $"{famName} : {typName}"; // Volledige label
            if (!string.IsNullOrEmpty(typName))
                return typName; // Alleen type

            try
            {
                var p = obj.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME);
                var s = p?.AsString();
                if (!string.IsNullOrEmpty(s))
                    return s;
            }
            catch { }

            try
            {
                var p = obj.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                var s = p?.AsString();
                if (!string.IsNullOrEmpty(s))
                    return s;
            }
            catch { }

#if REVIT2023
            return $"Type {obj.Id.IntegerValue}";
#else
            return $"Type {obj.Id.Value}";
#endif
        }

        public static double GetHeight(Element el, Document doc)
        {
            var bb = GetBoundingBox(el, doc);
            return bb != null ? bb.Max.Y - bb.Min.Y : 0.0;
        }

        public static bool IsLegendComponent(Element el)
        {
            try
            {
#if REVIT2023
                return el != null && el.Category != null &&
                       el.Category.Id.IntegerValue == (int)BuiltInCategory.OST_LegendComponents;
#else
                return el != null && el.Category != null &&
                       el.Category.Id.Value == (int)BuiltInCategory.OST_LegendComponents;
#endif
            }
            catch
            {
                return false;
            }
        }

        public static BoundingBoxXYZ GetBoundingBox(Element el, Document doc)
        {
            try
            {
                var view = doc.GetElement(el.OwnerViewId) as View;
                return el.get_BoundingBox(view);
            }
            catch
            {
                return null;
            }
        }

        public static double GetCenterX(Element el, Document doc)
        {
            var bb = GetBoundingBox(el, doc);
            return bb != null ? (bb.Min.X + bb.Max.X) / 2.0 : 0.0;
        }

        public static double GetBottomY(Element el, Document doc)
        {
            var bb = GetBoundingBox(el, doc);
            return bb?.Min.Y ?? 0.0;
        }

        public static double GetWidth(Element el, Document doc)
        {
            var bb = GetBoundingBox(el, doc);
            return bb != null ? bb.Max.X - bb.Min.X : 0.0;
        }

        public static (double minX, double minY, double maxX, double maxY) GetMinMax(Element el, Document doc)
        {
            var bb = GetBoundingBox(el, doc);
            return bb != null ? (bb.Min.X, bb.Min.Y, bb.Max.X, bb.Max.Y) : (0, 0, 0, 0);
        }

        public static void MoveVertical(Document doc, Element el, double newBottomY)
        {
            var bb = GetBoundingBox(el, doc);
            if (bb == null) return;

            double dy = newBottomY - bb.Min.Y;
            if (Math.Abs(dy) > 1e-6)
            {
                ElementTransformUtils.MoveElement(doc, el.Id, new XYZ(0.0, dy, 0.0));
                doc.Regenerate();
            }
        }

        public static (double axisX, double error) AlignPairXOnly(Document doc, Element topEl, Element botEl, 
            double axisXHint, int maxIter = 6, double tolMM = 0.05)
        {
            double tol = tolMM * FT_PER_MM;
            double axisX = axisXHint;

            for (int i = 0; i < maxIter; i++)
            {
                double ct = GetCenterX(topEl, doc);
                double cb = GetCenterX(botEl, doc);
                double dxT = axisX - ct;
                double dxB = axisX - cb;

                if (Math.Abs(ct - cb) <= tol && Math.Abs(dxT) < 1e-6 && Math.Abs(dxB) < 1e-6)
                    return (axisX, Math.Abs(ct - cb));

                if (Math.Abs(dxT) > 1e-6)
                    ElementTransformUtils.MoveElement(doc, topEl.Id, new XYZ(dxT, 0.0, 0.0));
                if (Math.Abs(dxB) > 1e-6)
                    ElementTransformUtils.MoveElement(doc, botEl.Id, new XYZ(dxB, 0.0, 0.0));

                doc.Regenerate();

                ct = GetCenterX(topEl, doc);
                cb = GetCenterX(botEl, doc);
                if (Math.Abs(ct - cb) <= tol)
                    return (axisX, Math.Abs(ct - cb));

                axisX = (ct + cb) / 2.0;
            }

            return (axisX, Math.Abs(GetCenterX(topEl, doc) - GetCenterX(botEl, doc)));
        }

        public static Element CopyAndMove(Document doc, ElementId elId, XYZ delta)
        {
            try
            {
                var newIds = ElementTransformUtils.CopyElement(doc, elId, delta);
                if (newIds != null && newIds.Count == 1)
                    return doc.GetElement(newIds.First());
            }
            catch { }
            return null;
        }

        public static Parameter FindViewParameter(Element el)
        {
            foreach (Parameter p in el.Parameters)
            {
                try
                {
                    if (p.IsReadOnly) continue;
                    string n = p.Definition?.Name?.Trim().ToLower() ?? "";
                    if (n.Contains("view direction") || n.Contains("weergaverichting"))
                        return p;
                }
                catch { }
            }
            return null;
        }

        public static int? DiscoverCodeForLabel(Document doc, Element el, List<string> labels, 
            int rangeStart = -20, int rangeEnd = 1)
        {
            var p = FindViewParameter(el);
            if (p == null) return null;

            int? old = null;
            try { old = p.AsInteger(); } catch { }

            bool Matches()
            {
                try
                {
                    string cur = p.AsValueString()?.Trim() ?? "";
                    string vn = new string(cur.ToLower().Where(char.IsLetterOrDigit).ToArray());
                    foreach (var lab in labels)
                    {
                        string ln = new string((lab ?? "").ToLower().Where(char.IsLetterOrDigit).ToArray());
                        if (vn == ln) return true;
                    }
                }
                catch { }
                return false;
            }

            for (int v = rangeStart; v < rangeEnd; v++)
            {
                try
                {
                    if (p.Set(v))
                    {
                        doc.Regenerate();
                        if (Matches())
                        {
                            if (old.HasValue)
                            {
                                try { p.Set(old.Value); doc.Regenerate(); } catch { }
                            }
                            return v;
                        }
                    }
                }
                catch { }
            }

            if (old.HasValue)
            {
                try { p.Set(old.Value); doc.Regenerate(); } catch { }
            }
            return null;
        }

        public static (bool success, string value) SetViewByCode(Document doc, Element el, int code, string labelFallback = null)
        {
            var p = FindViewParameter(el);
            if (p == null) return (false, null);

            try
            {
                if (p.StorageType == StorageType.Integer)
                {
                    p.Set(code);
                    doc.Regenerate();
                    return (true, p.AsValueString());
                }
            }
            catch { }

            if (!string.IsNullOrEmpty(labelFallback))
            {
                try
                {
                    p.SetValueString(labelFallback);
                    doc.Regenerate();
                    return (true, p.AsValueString());
                }
                catch { }
            }

            return (false, null);
        }

        public static ReferencePlane CreateReferencePlane(Document doc, View view, XYZ p0, XYZ p1, string name = null)
        {
            ReferencePlane rp = null;
            try
            {
                rp = doc.Create.NewReferencePlane(p0, p1, new XYZ(0, 0, 1), view);
            }
            catch
            {
                try
                {
                    rp = doc.Create.NewReferencePlane2(p0, p1, p0 + new XYZ(0, 0, 1), view);
                }
                catch { }
            }

            if (rp != null && !string.IsNullOrEmpty(name))
            {
                try { rp.Name = name; } catch { }
            }

            // Note: SetDatumExtentTypeInView is not available in all Revit API versions

            return rp;
        }

        public static DetailCurve CreateDetailLine(Document doc, View view, XYZ p0, XYZ p1)
        {
            try
            {
                var ln = Line.CreateBound(p0, p1);
                return doc.Create.NewDetailCurve(view, ln);
            }
            catch
            {
                return null;
            }
        }
    }
}