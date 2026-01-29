using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Globalization;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using static VH_Tools.Models.Constants;
using static VH_Tools.Utilities.CalculatorHelpers;
using VH_Tools.Models;

namespace VH_Tools.Utilities
{
    public class NENCalculator
    {
        private readonly Document _doc;
        private const double REQUIRED_RATIO = 0.55;

        public NENCalculator(Document doc)
        {
            _doc = doc;
        }

        /// <summary>
        /// Main entry point for NEN 2057 calculation. 
        /// Now accepts WindowResults directly to use in-memory values for CSV export.
        /// </summary>
        public void Calculate(List<FamilyInstance> windows, List<WindowResult> results, bool doExport)
        {
            // Collect window data
            var windowInfos = new List<WindowInfo>();

            foreach (var window in windows)
            {
                // Find matching result from memory
                var res = results.FirstOrDefault(r => r.ElementId == window.Id.IntegerValue);
                
                // Get values from memory FIRST, then fallback to parameters
                double? alphaDeg = res?.AlphaAvgDeg ?? GetAngleDeg(window, PARAM_ALPHA);
                double? betaDeg = res?.BetaDeg 
                                  ?? GetAngleDeg(window, PARAM_BETA_EPS) 
                                  ?? GetAngleDeg(window, "VH_kozijn_β") 
                                  ?? GetAngleDeg(window, "VH_kozijn_ε");
                double? AdM2 = res?.GlassM2.HasValue == true ? res.GlassM2.Value : GetAreaM2(window, "VH_kozijn_Ad");

                double? Cb = null;
                double? AeM2 = null;

                if (alphaDeg.HasValue && betaDeg.HasValue)
                {
                    Cb = CbTable.GetCb(alphaDeg.Value, betaDeg.Value);
                    if (Cb.HasValue && AdM2.HasValue)
                    {
                        AeM2 = AdM2.Value * Cb.Value;

                        // Try to set parameters if they exist, but don't fail if they don't
                        SetDouble(window, "VH_kozijn_Cb", Cb.Value);
                        SetDouble(window, "VH_kozijn_Ae", AeM2.Value, asAreaM2: true);
                    }
                }

                // Add to list even if incomplete (for "Kozijnen" tab)
                windowInfos.Add(new WindowInfo
                {
                    Element = window,
                    Id = window.Id.IntegerValue,
                    AlphaDeg = alphaDeg,
                    BetaDeg = betaDeg,
                    AdM2 = AdM2,
                    Cb = Cb,
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
                            if (info.AeM2.HasValue)
                            {
                                ad.AeSum += info.AeM2.Value;
                            }
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
                ExportExcel(windowInfos, vgResults);
            }
        }

        private void ExportExcel(List<WindowInfo> windowInfos, List<VGResult> vgResults)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            // Use .xls for XML Spreadsheet 2003 (warning is unavoidable but harmless)
            string excelPath = Path.Combine(desktop, $"NEN_daglicht_{timestamp}_detail.xls");

            using (var writer = new StreamWriter(excelPath, false, Encoding.UTF8))
            {
                // XML Header for SpreadsheetML
                writer.WriteLine("<?xml version=\"1.0\"?>");
                writer.WriteLine("<?mso-application progid=\"Excel.Sheet\"?>");
                writer.WriteLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
                writer.WriteLine(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
                writer.WriteLine(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"");
                writer.WriteLine(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"");
                writer.WriteLine(" xmlns:html=\"http://www.w3.org/TR/REC-html40\">");
                
                writer.WriteLine(" <Styles>");
                writer.WriteLine("  <Style ss:ID=\"Default\" ss:Name=\"Normal\">");
                writer.WriteLine("   <Alignment ss:Vertical=\"Bottom\"/>");
                writer.WriteLine("   <Borders/>");
                writer.WriteLine("   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/>");
                writer.WriteLine("   <Interior/>");
                writer.WriteLine("   <NumberFormat/>");
                writer.WriteLine("   <Protection/>");
                writer.WriteLine("  </Style>");
                writer.WriteLine("  <Style ss:ID=\"Header\">");
                writer.WriteLine("   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Bold=\"1\"/>");
                writer.WriteLine("  </Style>");
                writer.WriteLine("  <Style ss:ID=\"Title\">");
                writer.WriteLine("   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"14\" ss:Color=\"#000000\" ss:Bold=\"1\"/>");
                writer.WriteLine("  </Style>");
                writer.WriteLine(" </Styles>");

                // --- TAB 1: Gebieden ---
                writer.WriteLine(" <Worksheet ss:Name=\"Gebieden\">");
                writer.WriteLine("  <Table>");
                
                var validVgs = vgResults.Where(v => v.AantalKozijnen > 0).ToList();
                if (!validVgs.Any())
                {
                    writer.WriteLine("   <Row><Cell><Data ss:Type=\"String\">Geen verblijfsgebieden met gekoppelde kozijnen gevonden in de selectie.</Data></Cell></Row>");
                }
                else
                {
                    foreach (var vg in validVgs)
                    {
                        double AVG = vg.AVGm2;
                        double AeVG = vg.AeVGm2;
                        double eisAe = REQUIRED_RATIO * AVG;
                        double reductieVG = vg.VGReductieM2;

                        // Title row
                        writer.WriteLine("   <Row>");
                        writer.WriteLine($"    <Cell ss:StyleID=\"Title\"><Data ss:Type=\"String\">VERBLIJFSGEBIED</Data></Cell>");
                        writer.WriteLine("    <Cell/><Cell ss:StyleID=\"Header\"><Data ss:Type=\"String\">" + Esc(vg.VGNaam) + "</Data></Cell>");
                        writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + AVG.ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                        writer.WriteLine("   </Row>");

                        if (vg.StatusVG == "OK")
                        {
                            writer.WriteLine("   <Row><Cell><Data ss:Type=\"String\">geen aanpassingen benodigd</Data></Cell><Cell/><Cell/><Cell><Data ss:Type=\"Number\">0</Data></Cell></Row>");
                        }
                        else
                        {
                            writer.WriteLine("   <Row><Cell><Data ss:Type=\"String\">benodigde reductie VG</Data></Cell><Cell/><Cell/><Cell><Data ss:Type=\"Number\">" + reductieVG.ToString(CultureInfo.InvariantCulture) + "</Data></Cell></Row>");
                        }

                        writer.WriteLine("   <Row><Cell><Data ss:Type=\"String\">verblijsgebied totaal</Data></Cell><Cell/><Cell/><Cell><Data ss:Type=\"Number\">" + AVG.ToString(CultureInfo.InvariantCulture) + "</Data></Cell></Row>");
                        writer.WriteLine("   <Row/>"); // Empty row

                        // Header windows
                        writer.WriteLine("   <Row ss:StyleID=\"Header\">");
                        writer.WriteLine("    <Cell><Data ss:Type=\"String\">Koz</Data></Cell>");
                        writer.WriteLine("    <Cell><Data ss:Type=\"String\">Ruimte</Data></Cell>");
                        writer.WriteLine("    <Cell><Data ss:Type=\"String\">Ad,i</Data></Cell>");
                        writer.WriteLine("    <Cell><Data ss:Type=\"String\">α</Data></Cell>");
                        writer.WriteLine("    <Cell><Data ss:Type=\"String\">β / ε</Data></Cell>");
                        writer.WriteLine("    <Cell><Data ss:Type=\"String\">Cb,i</Data></Cell>");
                        writer.WriteLine("    <Cell><Data ss:Type=\"String\">Cu,i</Data></Cell>");
                        writer.WriteLine("    <Cell><Data ss:Type=\"String\">CLTA</Data></Cell>");
                        writer.WriteLine("    <Cell><Data ss:Type=\"String\">Aantal</Data></Cell>");
                        writer.WriteLine("    <Cell><Data ss:Type=\"String\">Ae,i</Data></Cell>");
                        writer.WriteLine("   </Row>");

                        foreach (var w in windowInfos.Where(wi => wi.VGId == vg.VGId))
                        {
                            string kozCode = GetWindowCode(w.Element);
                            string ruimte = GetWindowRoom(w.Element) ?? "";
                            writer.WriteLine("   <Row>");
                            writer.WriteLine("    <Cell><Data ss:Type=\"String\">" + Esc(kozCode) + "</Data></Cell>");
                            writer.WriteLine("    <Cell><Data ss:Type=\"String\">" + Esc(ruimte) + "</Data></Cell>");
                            writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + (w.AdM2 ?? 0).ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                            writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + (w.AlphaDeg ?? 0).ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                            writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + (w.BetaDeg ?? 0).ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                            writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + (w.Cb ?? 0).ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                            writer.WriteLine("    <Cell><Data ss:Type=\"Number\">1</Data></Cell>");
                            writer.WriteLine("    <Cell><Data ss:Type=\"Number\">1</Data></Cell>");
                            writer.WriteLine("    <Cell><Data ss:Type=\"Number\">1</Data></Cell>");
                            writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + (w.AeM2 ?? 0).ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                            writer.WriteLine("   </Row>");
                        }

                        writer.WriteLine("   <Row/>");

                        string voldoet = vg.StatusVG == "OK" ? "VOLDOET" : "VOLDOET NIET";
                        writer.WriteLine("   <Row>");
                        writer.WriteLine($"    <Cell ss:StyleID=\"Header\"><Data ss:Type=\"String\">{voldoet}</Data></Cell>");
                        writer.WriteLine("    <Cell/><Cell/><Cell/><Cell/><Cell/><Cell ss:StyleID=\"Header\"><Data ss:Type=\"String\">Totaal Ae,i aanwezig</Data></Cell>");
                        writer.WriteLine("    <Cell/><Cell/><Cell ss:StyleID=\"Header\"><Data ss:Type=\"Number\">" + AeVG.ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                        writer.WriteLine("   </Row>");
                        writer.WriteLine("   <Row>");
                        writer.WriteLine("    <Cell/><Cell/><Cell/><Cell/><Cell/><Cell/><Cell ss:StyleID=\"Header\"><Data ss:Type=\"String\">Totaal Ae,i EIS - VG (0,55·A_VG)</Data></Cell>");
                        writer.WriteLine("    <Cell/><Cell/><Cell ss:StyleID=\"Header\"><Data ss:Type=\"Number\">" + eisAe.ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                        writer.WriteLine("   </Row>");
                        writer.WriteLine("   <Row/><Row/>");
                    }
                }
                writer.WriteLine("  </Table>");
                writer.WriteLine(" </Worksheet>");

                // --- TAB 2: Kozijnlijst ---
                writer.WriteLine(" <Worksheet ss:Name=\"Kozijnlijst\">");
                writer.WriteLine("  <Table>");
                writer.WriteLine("   <Row ss:StyleID=\"Header\">");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">Koz</Data></Cell>");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">Ruimte</Data></Cell>");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">VG</Data></Cell>");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">Ad,i</Data></Cell>");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">α</Data></Cell>");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">β / ε</Data></Cell>");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">Cb,i</Data></Cell>");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">Cu,i</Data></Cell>");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">CLTA</Data></Cell>");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">Aantal</Data></Cell>");
                writer.WriteLine("    <Cell><Data ss:Type=\"String\">Ae,i</Data></Cell>");
                writer.WriteLine("   </Row>");

                foreach (var w in windowInfos.OrderBy(wi => GetWindowCode(wi.Element)))
                {
                    string kozCode = GetWindowCode(w.Element);
                    string ruimte = GetWindowRoom(w.Element) ?? "";
                    string vgName = "";
                    if (w.VGId.HasValue)
                    {
                        var vg = vgResults.FirstOrDefault(v => v.VGId == w.VGId.Value);
                        vgName = vg?.VGNaam ?? "";
                    }

                    writer.WriteLine("   <Row>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"String\">" + Esc(kozCode) + "</Data></Cell>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"String\">" + Esc(ruimte) + "</Data></Cell>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"String\">" + Esc(vgName) + "</Data></Cell>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + (w.AdM2 ?? 0).ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + (w.AlphaDeg ?? 0).ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + (w.BetaDeg ?? 0).ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + (w.Cb ?? 0).ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"Number\">1</Data></Cell>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"Number\">1</Data></Cell>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"Number\">1</Data></Cell>");
                    writer.WriteLine("    <Cell><Data ss:Type=\"Number\">" + (w.AeM2 ?? 0).ToString(CultureInfo.InvariantCulture) + "</Data></Cell>");
                    writer.WriteLine("   </Row>");
                }
                writer.WriteLine("  </Table>");
                writer.WriteLine(" </Worksheet>");

                writer.WriteLine("</Workbook>");
            }
        }

        private string Esc(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
        }

        private double? GetAngleDeg(Element elem, string paramName)
        {
            Parameter p = elem.LookupParameter(paramName);
            if (p != null && (p.StorageType == StorageType.Double || p.StorageType == StorageType.Integer))
            {
                if (p.StorageType == StorageType.Double)
                    return p.AsDouble() * 180.0 / Math.PI;
                else
                    return (double)p.AsInteger();
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
                    if (p != null)
                    {
                        string s = p.AsString();
                        if (!string.IsNullOrEmpty(s))
                            return s;
                        
                        s = p.AsValueString();
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
                    var phases = fi.Document.Phases.Cast<Phase>().ToList();
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

        private class WindowInfo
        {
            public FamilyInstance Element { get; set; }
            public int Id { get; set; }
            public double? AlphaDeg { get; set; }
            public double? BetaDeg { get; set; }
            public double? AdM2 { get; set; }
            public double? Cb { get; set; }
            public double? AeM2 { get; set; }
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
