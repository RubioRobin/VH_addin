using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using static VH_DaglichtPlugin.Constants;
using static VH_DaglichtPlugin.CalculatorHelpers;

namespace VH_DaglichtPlugin
{
    public class NENCalculator
    {
        private readonly Document _doc;
        private const double REQUIRED_RATIO = 0.55;

        public NENCalculator(Document doc)
        {
            _doc = doc;
        }

        public void Calculate(List<FamilyInstance> windows, bool doExport)
        {
            // Collect window data
            var windowInfos = new List<WindowInfo>();

            foreach (var window in windows)
            {
                double? alphaDeg = GetAngleDeg(window, PARAM_ALPHA);
                double? betaDeg = GetAngleDeg(window, PARAM_BETA_EPS) 
                                  ?? GetAngleDeg(window, "VH_kozijn_β") 
                                  ?? GetAngleDeg(window, "VH_kozijn_ε");
                double? AdM2 = GetAreaM2(window, "VH_kozijn_Ad");

                if (!alphaDeg.HasValue || !betaDeg.HasValue || !AdM2.HasValue)
                    continue;

                double? Cb = CbTable.GetCb(alphaDeg.Value, betaDeg.Value);
                if (!Cb.HasValue)
                    continue;

                double AeM2 = AdM2.Value * Cb.Value;

                // Set parameters
                SetDouble(window, "VH_kozijn_Cb", Cb.Value);
                SetDouble(window, "VH_kozijn_Ae", AeM2, asAreaM2: true);

                windowInfos.Add(new WindowInfo
                {
                    Element = window,
                    Id = window.Id.IntegerValue,
                    AlphaDeg = alphaDeg.Value,
                    BetaDeg = betaDeg.Value,
                    AdM2 = AdM2.Value,
                    Cb = Cb.Value,
                    AeM2 = AeM2,
                    VGId = null
                });
            }

            // Collect VG areas
            var areaData = new List<AreaData>();
            var areasRaw = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Areas)
                .WhereElementIsNotElementType()
                .Cast<Area>()
                .ToList();

            foreach (var area in areasRaw)
            {
                double? areaM2 = GetVGAreaM2(area);
                if (!areaM2.HasValue || areaM2.Value <= 0)
                    continue;

                var loops = GetAreaLoops(area);
                if (loops.Count == 0)
                    continue;

                string name = area.Name ?? $"VG_{area.Id.IntegerValue}";
                string group = GetParamString(area, new[] { "VG_groep", "Groep", "Group", "Area Group" });
                string levelName = null;
                try
                {
                    Level lvl = _doc.GetElement(area.LevelId) as Level;
                    levelName = lvl?.Name;
                }
                catch { }

                areaData.Add(new AreaData
                {
                    Id = area.Id.IntegerValue,
                    Name = name,
                    Group = group,
                    Level = levelName,
                    AreaM2 = areaM2.Value,
                    Loops = loops,
                    AeSum = 0,
                    Kozijnen = new List<int>(),
                    Cbi = null
                });
            }

            // Link windows to VG areas
            foreach (var info in windowInfos)
            {
                XYZ pt = GetBBoxCenter(info.Element);
                if (pt == null)
                    continue;

                foreach (var ad in areaData)
                {
                    bool found = false;
                    foreach (var loop in ad.Loops)
                    {
                        if (PointInPolygon(pt, loop))
                        {
                            ad.AeSum += info.AeM2;
                            ad.Kozijnen.Add(info.Id);
                            info.VGId = ad.Id;
                            found = true;
                            break;
                        }
                    }
                    if (found) break;
                }
            }

            // Calculate Cbi per VG
            var vgResults = new List<VGResult>();

            foreach (var ad in areaData)
            {
                double AVG = ad.AreaM2;
                double AeVG = ad.AeSum;
                double vereistAe = REQUIRED_RATIO * AVG;
                double verhouding = AVG > 0 ? AeVG / AVG : 0;
                double tekortAe = Math.Max(0, vereistAe - AeVG);
                double maxAVG = AeVG > 0 ? AeVG / REQUIRED_RATIO : 0;
                double vgReductie = Math.Max(0, AVG - maxAVG);
                double CbiVG = AVG > 0 ? AeVG / (REQUIRED_RATIO * AVG) : 0;
                ad.Cbi = CbiVG;

                string statusVG = CbiVG >= 1.0 ? "OK" : "NIET_OK";

                vgResults.Add(new VGResult
                {
                    VGId = ad.Id,
                    VGNaam = ad.Name,
                    VGGroup = ad.Group,
                    VGLevel = ad.Level,
                    AVGm2 = AVG,
                    AeVGm2 = AeVG,
                    AeDivAVG = verhouding,
                    CbiVG = CbiVG,
                    EisRatio = REQUIRED_RATIO,
                    TekortAeM2 = tekortAe,
                    VGReductieM2 = vgReductie,
                    AantalKozijnen = ad.Kozijnen.Count,
                    StatusVG = statusVG
                });
            }

            // Set Cbi on windows
            foreach (var info in windowInfos)
            {
                if (info.VGId.HasValue)
                {
                    var ad = areaData.FirstOrDefault(a => a.Id == info.VGId.Value);
                    if (ad != null && ad.Cbi.HasValue)
                    {
                        SetDouble(info.Element, "VH_kozijn_Cbi", ad.Cbi.Value);
                    }
                }
            }

            // Export CSV if requested
            if (doExport)
            {
                ExportCSV(windowInfos, vgResults);
            }

            // Show summary
            string summary = $"NEN 55% Toets voltooid\n\n";
            summary += $"Kozijnen verwerkt: {windowInfos.Count}\n";
            summary += $"VG-gebieden: {vgResults.Count}\n";
            summary += $"Voldoet: {vgResults.Count(v => v.StatusVG == "OK")}\n";
            summary += $"Voldoet niet: {vgResults.Count(v => v.StatusVG == "NIET_OK")}\n\n";
            summary += "CSV export aangemaakt op bureaublad.";

            // TaskDialog.Show("NEN Daglichttoets", summary);
        }

        private void ExportCSV(List<WindowInfo> windowInfos, List<VGResult> vgResults)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string csvPath = Path.Combine(desktop, $"NEN_daglicht_{timestamp}_detail.csv");

            // Use Encoding.Default (ANSI) instead of UTF8 to avoid BOM issues in some Excel versions
            using (var writer = new StreamWriter(csvPath, false, Encoding.Default))
            {
                foreach (var vg in vgResults.Where(v => v.AantalKozijnen > 0))
                {
                    double AVG = vg.AVGm2;
                    double AeVG = vg.AeVGm2;
                    double eisAe = REQUIRED_RATIO * AVG;
                    double cbi = vg.CbiVG;
                    double reductieVG = vg.VGReductieM2;

                    writer.WriteLine($"VERBLIJFSGEBIED;;{vg.VGNaam};{Fmt2(AVG)};;;;;;;;");

                    if (vg.StatusVG == "OK")
                    {
                        writer.WriteLine($"geen aanpassingen benodigd;;;{Fmt2(0)};;;;;;;;");
                    }
                    else
                    {
                        writer.WriteLine($"benodigde reductie VG;;;{Fmt2(reductieVG)};;;;;;;;");
                    }

                    writer.WriteLine($"verblijsgebied totaal;;;{Fmt2(AVG)};;;;;;;;");
                    writer.WriteLine(";;;;;;;;;;;;");

                    writer.WriteLine("Koz;Ruimte;Ad,i;α;β / ε;Cb,i;Cu,i;CLTA;Aantal;Ae,i;;");

                    foreach (var w in windowInfos.Where(wi => wi.VGId == vg.VGId))
                    {
                        string kozCode = GetWindowCode(w.Element);
                        string ruimte = GetWindowRoom(w.Element) ?? "";
                        writer.WriteLine($"{kozCode};{ruimte};{Fmt2(w.AdM2)};{FmtDeg(w.AlphaDeg)};{FmtDeg(w.BetaDeg)};{Fmt2(w.Cb)};{Fmt2(1.0)};{Fmt2(1.0)};1;{Fmt2(w.AeM2)};;");
                    }

                    writer.WriteLine(";;;;;;;;;;;;");

                    string voldoet = vg.StatusVG == "OK" ? "VOLDOET" : "VOLDOET NIET";
                    writer.WriteLine($"{voldoet};;;;;;Totaal Ae,i;aanwezig;;{Fmt2(AeVG)};;");
                    writer.WriteLine($";;;;;;Totaal Ae,i;EIS - VG (0,55·A_VG);;{Fmt2(eisAe)};;");
                    writer.WriteLine(";;;;;;;;;;;;");
                    writer.WriteLine(";;;;;;;;;;;;");
                }
            }
        }

        // Helper methods
        private double? GetAngleDeg(Element elem, string paramName)
        {
            Parameter p = elem.LookupParameter(paramName);
            if (p != null && p.StorageType == StorageType.Double)
            {
                return p.AsDouble() * 180.0 / Math.PI;
            }
            return null;
        }

        private double? GetAreaM2(Element elem, string paramName)
        {
            Parameter p = elem.LookupParameter(paramName);
            if (p != null && p.StorageType == StorageType.Double)
            {
                return p.AsDouble() * SQFT_TO_SQM;
            }
            return null;
        }

        private void SetDouble(Element elem, string paramName, double value, bool asAreaM2 = false)
        {
            Parameter p = elem.LookupParameter(paramName);
            if (p != null && !p.IsReadOnly)
            {
                double v = asAreaM2 ? value / SQFT_TO_SQM : value;
                try { p.Set(v); } catch { }
            }
        }

        private double? GetVGAreaM2(Area area)
        {
            try
            {
                Parameter p = area.get_Parameter(BuiltInParameter.ROOM_AREA);
                if (p != null)
                    return p.AsDouble() * SQFT_TO_SQM;
            }
            catch { }
            return null;
        }

        private List<List<XYZ>> GetAreaLoops(Area area)
        {
            var loops = new List<List<XYZ>>();
            try
            {
                var opt = new SpatialElementBoundaryOptions();
                var segLists = area.GetBoundarySegments(opt);
                if (segLists == null) return loops;

                foreach (var segList in segLists)
                {
                    var pts = new List<XYZ>();
                    foreach (BoundarySegment seg in segList)
                    {
                        Curve c = seg.GetCurve();
                        pts.Add(c.GetEndPoint(0));
                    }
                    if (pts.Count >= 3 && pts[0].DistanceTo(pts[pts.Count - 1]) < 1e-6)
                    {
                        pts.RemoveAt(pts.Count - 1);
                    }
                    if (pts.Count >= 3)
                    {
                        loops.Add(pts);
                    }
                }
            }
            catch { }
            return loops;
        }

        private bool PointInPolygon(XYZ pt, List<XYZ> poly)
        {
            if (poly == null || poly.Count < 3)
                return false;

            double x = pt.X, y = pt.Y;
            bool inside = false;
            int n = poly.Count;

            for (int i = 0; i < n; i++)
            {
                XYZ p1 = poly[i];
                XYZ p2 = poly[(i + 1) % n];
                double x1 = p1.X, y1 = p1.Y;
                double x2 = p2.X, y2 = p2.Y;

                if (Math.Abs(y2 - y1) < 1e-9)
                    continue;

                if ((y1 > y) != (y2 > y))
                {
                    double xinters = (y - y1) * (x2 - x1) / (y2 - y1) + x1;
                    if (x < xinters)
                        inside = !inside;
                }
            }
            return inside;
        }

        private XYZ GetBBoxCenter(Element elem)
        {
            try
            {
                BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                if (bb != null)
                {
                    return new XYZ(
                        (bb.Min.X + bb.Max.X) * 0.5,
                        (bb.Min.Y + bb.Max.Y) * 0.5,
                        (bb.Min.Z + bb.Max.Z) * 0.5
                    );
                }
            }
            catch { }
            return null;
        }

        private string GetParamString(Element elem, string[] paramNames)
        {
            foreach (string name in paramNames)
            {
                try
                {
                    Parameter p = elem.LookupParameter(name);
                    if (p != null && p.StorageType == StorageType.String)
                    {
                        string s = p.AsString();
                        if (!string.IsNullOrEmpty(s))
                            return s;
                    }
                }
                catch { }
            }
            return null;
        }

        private string GetWindowCode(Element elem)
        {
            string code = GetParamString(elem, new[] { "VH_kozijn_code", "Kozijnnummer", "Kozijn", "Mark", "Type Mark" });
            return code ?? $"Koz_{elem.Id.IntegerValue}";
        }

        private string GetWindowRoom(Element elem)
        {
            try
            {
                if (elem is FamilyInstance fi)
                {
                    var phases = _doc.Phases.Cast<Phase>().ToList();
                    if (phases.Any())
                    {
                        Phase phase = phases.Last();
                        Room room = fi.get_FromRoom(phase) ?? fi.get_ToRoom(phase);
                        return room?.Name;
                    }
                }
            }
            catch { }
            return null;
        }

        private string Fmt2(double? x)
        {
            if (!x.HasValue) return "";
            return x.Value.ToString("F2").Replace('.', ',');
        }

        private string FmtDeg(double? x)
        {
            if (!x.HasValue) return "";
            return $"{(int)Math.Round(x.Value)}°";
        }

        // Data classes
        private class WindowInfo
        {
            public FamilyInstance Element { get; set; }
            public int Id { get; set; }
            public double AlphaDeg { get; set; }
            public double BetaDeg { get; set; }
            public double AdM2 { get; set; }
            public double Cb { get; set; }
            public double AeM2 { get; set; }
            public int? VGId { get; set; }
        }

        private class AreaData
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string Group { get; set; }
            public string Level { get; set; }
            public double AreaM2 { get; set; }
            public List<List<XYZ>> Loops { get; set; }
            public double AeSum { get; set; }
            public List<int> Kozijnen { get; set; }
            public double? Cbi { get; set; }
        }

        private class VGResult
        {
            public int VGId { get; set; }
            public string VGNaam { get; set; }
            public string VGGroup { get; set; }
            public string VGLevel { get; set; }
            public double AVGm2 { get; set; }
            public double AeVGm2 { get; set; }
            public double AeDivAVG { get; set; }
            public double CbiVG { get; set; }
            public double EisRatio { get; set; }
            public double TekortAeM2 { get; set; }
            public double VGReductieM2 { get; set; }
            public int AantalKozijnen { get; set; }
            public string StatusVG { get; set; }
        }
    }
}
