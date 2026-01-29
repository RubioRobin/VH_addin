// ============================================================================
// KozijnstaatCommand.cs
//
// Doel: Start het Kozijnstaat venster voor het genereren van kozijnstaten en maatvoering.
// Gebruik: Wordt aangeroepen als Revit External Command. Opent een WPF-venster.
// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Deze command opent een UI waarmee de gebruiker automatisch kozijnstaten,
// maatvoering en optioneel offsetlijnen kan genereren. Zie Views/KozijnstaatSettingsWindow.xaml.
// ============================================================================

using Autodesk.Revit.Attributes; // Revit attribuut voor transactiemodus
using Autodesk.Revit.DB; // Revit database objecten
using Autodesk.Revit.UI; // Revit UI API
using Autodesk.Revit.UI.Selection; // Revit selectie API
using System; // Standaard .NET functionaliteit
using System.Collections.Generic; // Lijsten en collecties
using System.Linq; // LINQ voor collecties
using VH_Tools.Models; // Eigen model voor instellingen
using VH_Tools.Utilities; // Eigen hulpfuncties
using VH_Addin.Views; // Eigen vensters

namespace VH_Tools.Commands // Hoofdnamespace voor alle commands van VH Tools
{
    [Transaction(TransactionMode.Manual)] // Revit: deze command start een handmatige transactie
    [Regeneration(RegenerationOption.Manual)] // Revit: handmatige regeneratie
    public class KozijnstaatCommand : IExternalCommand // Implementatie van een externe Revit-command
    {
        private Document _doc; // Huidig Revit-document
        private UIDocument _uidoc; // Huidig UI-document
        private KozijnstaatConfig _config; // Instellingen voor kozijnstaat
        private int _refPlaneCounter = 1; // Teller voor referentievlakken

        // Mogelijke labels voor achteraanzichten
        private static readonly List<string> BackLabels = new List<string> 
        { 
            "Elevation : Back", "Elevation: Back", "Back", "Achteraanzicht" 
        };

        // Mogelijke labels voor vooraanzichten
        private static readonly List<string> FrontLabels = new List<string> 
        { 
            "Elevation : Front", "Elevation: Front", "Front", "Voorkant", "Voorzijde" 
        };
        
        // Mogelijke labels voor plattegronden
        private static readonly List<string> PlanLabels = new List<string> 
        { 
            "Floor Plan", "Plattegrond" 
        };

        /// <summary>
        /// Hoofdmethode die wordt aangeroepen door Revit wanneer de command wordt gestart.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // TaskDialog.Show("TEST", "Nieuwe versie van KozijnstaatCommand is geladen! (" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ")");
            try
            {
                _uidoc = commandData.Application.ActiveUIDocument;
                _doc = _uidoc.Document;

                if (_doc == null)
                {
                    TaskDialog.Show("Fout", "Open een Revit-project.");
                    return Result.Failed;
                }

                // Show settings dialog
                // Load saved settings or defaults
                _config = SettingsManager.LoadSettings();
                
                var settingsWindow = new KozijnstaatSettingsWindow(_config, _doc);
                
                if (settingsWindow.ShowDialog() != true || !settingsWindow.DialogAccepted)
                {
                    SettingsManager.SaveSettings(_config);
                    return Result.Cancelled;
                }

                // Save settings immediately after valid dialog confirmation
                SettingsManager.SaveSettings(_config);

                // Pick legend component
                Element baseLc = PickOneLegendComponent();
                if (baseLc == null)
                {
                    TaskDialog.Show("Geannuleerd", "Geen Legend Component geselecteerd.");
                    return Result.Cancelled;
                }

                View view = _doc.GetElement(baseLc.OwnerViewId) as View;
                if (view == null || view.ViewType != ViewType.Legend)
                {
                    TaskDialog.Show("Fout", "Kies een Legend Component op een Legend View.");
                    return Result.Failed;
                }

                // Determine which elements to include based on filter
                bool includeWindows = _config.ElementFilter == ElementTypeFilter.Kozijnen || 
                                     _config.ElementFilter == ElementTypeFilter.Beide;
                bool includeDoors = _config.ElementFilter == ElementTypeFilter.Deuren || 
                                   _config.ElementFilter == ElementTypeFilter.Beide;

                // Get element symbols based on filter
                var allSyms = WindowHelpers.GetElementSymbols(_doc, includeWindows, includeDoors);
                if (allSyms.Count == 0)
                {
                    string elementType = _config.ElementFilter == ElementTypeFilter.Deuren ? "Deuren" :
                                        _config.ElementFilter == ElementTypeFilter.Beide ? "Kozijnen/Deuren" : "Kozijnen";
                    string assemblyCode = _config.ElementFilter == ElementTypeFilter.Deuren ? "08.*" :
                                         _config.ElementFilter == ElementTypeFilter.Beide ? "08.*/31.*" : "31.*";
                    TaskDialog.Show("Fout", $"Geen {elementType} types met Assembly Code {assemblyCode} gevonden.");
                    return Result.Failed;
                }

                // Filter by placed instances
                var placedIds = WindowHelpers.GetPlacedElementTypeIds(_doc, includeWindows, includeDoors);
                if (placedIds.Count == 0)
                {
                    string elementType = _config.ElementFilter == ElementTypeFilter.Deuren ? "deuren" :
                                        _config.ElementFilter == ElementTypeFilter.Beide ? "kozijnen/deuren" : "kozijnen";
                    TaskDialog.Show("Fout", $"Geen geplaatste {elementType} in dit project.");
                    return Result.Failed;
                }

                allSyms = allSyms.Where(s => placedIds.Contains(s.Id.IntegerValue)).ToList();
                if (allSyms.Count == 0)
                {
                    string elementType = _config.ElementFilter == ElementTypeFilter.Deuren ? "deur" :
                                        _config.ElementFilter == ElementTypeFilter.Beide ? "kozijn/deur" : "kozijn";
                    TaskDialog.Show("Fout", $"Geen geschikte {elementType}-types (geen sparing/samenstelling-leeg, én geplaatst).");
                    return Result.Failed;
                }

                // Filter by Assembly Code if specified (empty = no filter)
                if (!string.IsNullOrWhiteSpace(_config.AssemblyCodeFilter))
                {
                    string filterPattern = _config.AssemblyCodeFilter.Trim();
                    allSyms = allSyms.Where(s => 
                    {
                        var asmCodeParam = s.LookupParameter("Assembly Code");
                        if (asmCodeParam == null) return false;
                        var asmCode = asmCodeParam.AsString() ?? "";
                        return asmCode.StartsWith(filterPattern);
                    }).ToList();

                    if (allSyms.Count == 0)
                    {
                        TaskDialog.Show("Fout", $"Geen types met Assembly Code '{_config.AssemblyCodeFilter}' gevonden.");
                        return Result.Failed;
                    }
                }

                // Get offsets if heights mode is enabled
                var offsetsByType = _config.UseHeights ? WindowHelpers.GetOffsetsByType(_doc, includeWindows, includeDoors) : new Dictionary<long, List<double>>();

                // Filter out already used types
                var usedIds = WindowHelpers.GetUsedWindowTypeIds(_doc);
                var availableSyms = allSyms.Where(s => !usedIds.Contains(s.Id.IntegerValue)).ToList();
                
                // Sort by type mark
                availableSyms = availableSyms.OrderBy(s => WindowHelpers.TypeMarkSortKey(s)).ToList();

                if (availableSyms.Count == 0)
                {
                    TaskDialog.Show("Info", "Geen beschikbare types (alles staat al op de Legend).");
                    return Result.Succeeded;
                }

                // Limit types if needed
                var remaining = _config.UseAllTypes 
                    ? new List<FamilySymbol>(availableSyms) 
                    : availableSyms.Take(Math.Min(_config.MaxTypes, availableSyms.Count)).ToList();

                if (remaining.Count == 0)
                {
                    TaskDialog.Show("Info", "Geen types om te plaatsen.");
                    return Result.Succeeded;
                }

                // Execute placement
                using (Transaction trans = new Transaction(_doc, "Kozijnstaat Genereren"))
                {
                    trans.Start();

                    var result = PlaceKozijnstaat(view, baseLc, remaining, offsetsByType);

                    trans.Commit();

                    if (result.success)
                    {
                        TaskDialog.Show("Succes", 
                            $"Kozijnstaat gegenereerd!\n\n" +
                            $"Geplaatst: {result.placedCount} types\n" +
                            $"Format: {_config.Format}\n" +
                            $"Schaal: {view.Scale}\n" +
                            $"Hoogtes-modus: {(_config.UseHeights ? "AAN" : "UIT")}");
                        return Result.Succeeded;
                    }
                    else
                    {
                        TaskDialog.Show("Fout", result.message ?? "Onbekende fout tijdens plaatsing.");
                        return Result.Failed;
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Fout", $"Er is een fout opgetreden:\n{ex.Message}\n\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private Element PickOneLegendComponent()
        {
            var sel = _uidoc.Selection.GetElementIds();
            if (sel != null && sel.Count == 1)
            {
                var el = _doc.GetElement(sel.First());
                if (RevitHelpers.IsLegendComponent(el))
                    return el;
            }

            while (true)
            {
                try
                {
                    var reference = _uidoc.Selection.PickObject(ObjectType.Element, "Kies 1 Legend Component op een Legend View");
                    var el = _doc.GetElement(reference.ElementId);
                    if (RevitHelpers.IsLegendComponent(el))
                        return el;
                    
                    TaskDialog.Show("Niet geldig", "Dit is geen Legend Component.");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return null;
                }
            }
        }

        private (bool success, int placedCount, string message) PlaceKozijnstaat(
            View view, 
            Element baseLc, 
            List<FamilySymbol> remaining, 
            Dictionary<long, List<double>> offsetsByType)
        {
            try
            {
                // Calculate paper dimensions
                int paperW_mm, paperH_mm;

                if (_config.Format == "Custom")
                {
                    paperW_mm = (int)_config.CustomWidth_MM;
                    paperH_mm = (int)_config.CustomHeight_MM;
                }
                else
                {
                    if (!KozijnstaatConfig.PaperSizes.TryGetValue(_config.Format, out var paperSize))
                        paperSize = KozijnstaatConfig.PaperSizes["A1"];

                    if (_config.Orientation.ToLower().StartsWith("lig"))
                    {
                        paperW_mm = Math.Max(paperSize.Item1, paperSize.Item2);
                        paperH_mm = Math.Min(paperSize.Item1, paperSize.Item2);
                    }
                    else
                    {
                        paperW_mm = Math.Min(paperSize.Item1, paperSize.Item2);
                        paperH_mm = Math.Max(paperSize.Item1, paperSize.Item2);
                    }
                }

                double marginBLR_mm = Math.Max(0, _config.MarginBLR_MM);
                double marginBottom_mm = Math.Max(0, _config.MarginBottom_MM);

                double usablePaperW_mm = Math.Max(0, paperW_mm - 2.0 * marginBLR_mm);
                double usablePaperH_mm = Math.Max(0, paperH_mm - (marginBLR_mm + marginBottom_mm));

                int scale = Math.Max(1, view.Scale);
                double targetModelW_mm = usablePaperW_mm * scale;
                double targetModelH_mm = usablePaperH_mm * scale;

                double targetW_ft = targetModelW_mm * RevitHelpers.FT_PER_MM;
                double targetH_ft = targetModelH_mm * RevitHelpers.FT_PER_MM;

                double dy_mm = Math.Max(RevitHelpers.MIN_ROW_DY_MM, _config.RowPitch_MM);
                double dy_ft = dy_mm * RevitHelpers.FT_PER_MM;

                double gap_ft = Math.Max(0, _config.Gap_MM) * RevitHelpers.FT_PER_MM;
                double minW_ft = RevitHelpers.MIN_SLOT_MM * RevitHelpers.FT_PER_MM;
                double lineOff_ft = Math.Max(0, _config.LineOffset_MM) * RevitHelpers.FT_PER_MM;
                double minBackGap_ft = Math.Max(0, _config.MinBackGap_MM) * RevitHelpers.FT_PER_MM;

                double textDy_ft = 1200.0 * RevitHelpers.FT_PER_MM;
                double textDx_ft = 0.0 * RevitHelpers.FT_PER_MM;
                double refLeft_ft = _config.RefPlaneLeft_MM * RevitHelpers.FT_PER_MM;
                double refTop_ft = _config.RefPlaneTop_MM * RevitHelpers.FT_PER_MM;

                // Calculate global max offset
                double maxOffGlobal = 0.0;
                if (_config.UseHeights)
                {
                    foreach (var arr in offsetsByType.Values)
                    {
                        if (arr != null && arr.Count > 0)
                        {
                            double m = arr.Max();
                            if (m > maxOffGlobal)
                                maxOffGlobal = m;
                        }
                    }
                }

                double rowPitchFtConst = Math.Max(
                    _config.RowPitch_MM * RevitHelpers.FT_PER_MM,
                    Math.Max(maxOffGlobal, minBackGap_ft) + dy_ft
                );

                // Clean bad legend components
                WindowHelpers.CleanBadLegendComponents(_doc, view, baseLc);

                var placed = new List<PlacedItem>();
                double totalUsedH = 0.0;
                double? leftAnchorX = null;

                double topY = RevitHelpers.GetBottomY(baseLc, _doc);
                double botY = topY - dy_ft;

                // Fill rows
                while (remaining.Count > 0)
                {
                    var (rowInfo, usedW_ft, rowPitch_ft, newLeftAnchor) = FillOneRow(
                        botY, remaining, leftAnchorX, baseLc, placed.Count > 0, 
                        targetW_ft, gap_ft, minW_ft, offsetsByType);

                    if (rowInfo.Count == 0)
                        break;

                    leftAnchorX = newLeftAnchor;

                    if (totalUsedH + rowPitch_ft > targetH_ft && placed.Count > 0)
                    {
                        // Delete overflow items
                        foreach (var ri in rowInfo)
                        {
                            try { _doc.Delete(ri.TopElement.Id); } catch { }
                            try { _doc.Delete(ri.BotElement.Id); } catch { }
                        }
                        break;
                    }

                    placed.AddRange(rowInfo);
                    totalUsedH += rowPitch_ft;

                    remaining = remaining.Skip(rowInfo.Count).ToList();
                    if (remaining.Count > 0)
                    {
                        topY = topY - rowPitch_ft;
                        botY = topY - dy_ft;
                    }
                }

                // Place GA tags
                var gaSym = AnnotationHelpers.PickDefaultGASymbol(_doc);
                if (placed.Count > 0 && gaSym != null)
                {
                    foreach (var info in placed)
                    {
                        var v = _doc.GetElement(info.BotElement.OwnerViewId) as View;
                        var (xmin, ymin, xmax, ymax) = RevitHelpers.GetMinMax(info.BotElement, _doc);
                        var pt = new XYZ(xmin + textDx_ft, ymin - textDy_ft, 0.0);
                        var ga = AnnotationHelpers.PlaceGATag(_doc, gaSym, v, pt);
                        
                        if (ga != null)
                        {
                            try
                            {
                                string km = WindowHelpers.GetKozijnmerk(info.Symbol) ?? "";
                                var p = ga.LookupParameter("tekst_kozijnmerk");
                                if (p != null && !p.IsReadOnly)
                                    p.Set(km);
                            }
                            catch { }
                        }
                    }
                }

                // Draw lines
                if (placed.Count > 0)
                {
                    DrawLines(placed, view, lineOff_ft);
                }

                // Create reference planes
                if (placed.Count > 0 && _config.CreateReferencePlanes)
                {
                    CreateReferenceCross(placed, view, refLeft_ft, refTop_ft, targetW_ft, targetH_ft);
                }

                 // ...existing code...

                return (true, placed.Count, null);
            }
            catch (Exception ex)
            {
                return (false, 0, ex.Message);
            }
        }

        private (List<PlacedItem> rowInfo, double usedW_ft, double rowPitch_ft, double? leftAnchorX) FillOneRow(
            double rowBotY, 
            List<FamilySymbol> typesForRow,
            double? leftAnchorX,
            Element baseLc,
            bool hasPlaced,
            double targetW_ft,
            double gap_ft,
            double minW_ft,
            Dictionary<long, List<double>> offsetsByType)
        {
            if (typesForRow.Count == 0)
                return (new List<PlacedItem>(), 0, 0, leftAnchorX);

            Element t0 = hasPlaced 
                ? RevitHelpers.CopyAndMove(_doc, baseLc.Id, new XYZ(0.01, 0, 0))
                : baseLc;

            if (t0 == null)
                return (new List<PlacedItem>(), 0, 0, leftAnchorX);

            double dy0 = rowBotY - RevitHelpers.GetBottomY(baseLc, _doc);
            Element b0 = RevitHelpers.CopyAndMove(_doc, baseLc.Id, new XYZ(0, dy0, 0));
            if (b0 == null)
                return (new List<PlacedItem>(), 0, 0, leftAnchorX);

            var (offs0, effTop0, trueMax0) = SetupPair(t0, b0, typesForRow[0], rowBotY, offsetsByType);

            var (xmin_t, _, _, _) = RevitHelpers.GetMinMax(t0, _doc);
            var (xmin_b, _, _, _) = RevitHelpers.GetMinMax(b0, _doc);
            double curLeft = Math.Min(xmin_t, xmin_b);

            if (!leftAnchorX.HasValue)
            {
                leftAnchorX = curLeft;
            }
            else
            {
                double dx = leftAnchorX.Value - curLeft;
                if (Math.Abs(dx) > 1e-6)
                {
                    ElementTransformUtils.MoveElement(_doc, t0.Id, new XYZ(dx, 0, 0));
                    ElementTransformUtils.MoveElement(_doc, b0.Id, new XYZ(dx, 0, 0));
                    _doc.Regenerate();
                }
            }

            var rowItems = new List<PlacedItem>
            {
                new PlacedItem
                {
                    TopElement = t0,
                    BotElement = b0,
                    Symbol = typesForRow[0],
                    Offsets = offs0,
                    EffTopOffset = effTop0,
                    TrueMaxOffset = trueMax0
                }
            };

            double usedW = EffWidth(t0, b0, minW_ft);
            double wPrev = usedW;
            double axisX = RevitHelpers.GetCenterX(t0, _doc);

            double kozijnHeight0 = RevitHelpers.GetHeight(t0, _doc);
            double maxPointAboveZero0 = Math.Max(kozijnHeight0, trueMax0);
            
            double rowPitchFt = Math.Max(
                _config.RowPitch_MM * RevitHelpers.FT_PER_MM,
                effTop0 + maxPointAboveZero0 + (1000.0 * RevitHelpers.FT_PER_MM)
            );

            for (int i = 1; i < typesForRow.Count; i++)
            {
                var sym = typesForRow[i];
                var ti = RevitHelpers.CopyAndMove(_doc, t0.Id, new XYZ(0.01 * (i + 1), 0, 0));
                var bi = RevitHelpers.CopyAndMove(_doc, t0.Id, new XYZ(0.01 * (i + 1), 0, 0));
                
                if (ti == null || bi == null)
                    break;

                var (offsI, effTopI, trueMaxI) = SetupPair(ti, bi, sym, rowBotY, offsetsByType);
                double wCur = EffWidth(ti, bi, minW_ft);

                if (usedW + gap_ft + wCur > targetW_ft)
                {
                    try { _doc.Delete(ti.Id); } catch { }
                    try { _doc.Delete(bi.Id); } catch { }
                    break;
                }

                axisX = axisX + (wPrev / 2.0) + gap_ft + (wCur / 2.0);
                RevitHelpers.AlignPairXOnly(_doc, ti, bi, axisX);

                rowItems.Add(new PlacedItem
                {
                    TopElement = ti,
                    BotElement = bi,
                    Symbol = sym,
                    Offsets = offsI,
                    EffTopOffset = effTopI,
                    TrueMaxOffset = trueMaxI
                });

                // Row pitch must account for:
                // 1. Basic row pitch from config
                // 2. The gap + Kozijn height (min ~7.2m) + any offsets above the Kozijn
                double kozijnHeight = RevitHelpers.GetHeight(ti, _doc);
                double maxPointAboveZero = Math.Max(kozijnHeight, trueMaxI);
                
                double currentRowPitch = Math.Max(
                    _config.RowPitch_MM * RevitHelpers.FT_PER_MM,
                    effTopI + maxPointAboveZero + (1000.0 * RevitHelpers.FT_PER_MM) // Add 1m margin for tags
                );
                
                if (currentRowPitch > rowPitchFt)
                    rowPitchFt = currentRowPitch;

                usedW += gap_ft + wCur;
                wPrev = wCur;
            }

            return (rowItems, usedW, rowPitchFt, leftAnchorX);
        }

        private (List<double> offsets, double effTopOffset, double trueMaxOffset) SetupPair(
            Element tEl, 
            Element bEl, 
            FamilySymbol sym,
            double rowBotY,
            Dictionary<long, List<double>> offsetsByType)
        {
            SetTypeOnLC(tEl, sym.Id);
            SetTypeOnLC(bEl, sym.Id);
            _doc.Regenerate();

            // Determine which elevation view to use based on config
            List<string> elevationLabels = _config.ViewType == "Front" ? FrontLabels : BackLabels;
            string elevationViewName = _config.ViewType == "Front" ? "Elevation : Front" : "Elevation : Back";

            int? cElevation = RevitHelpers.DiscoverCodeForLabel(_doc, tEl, elevationLabels, -20, 1) ?? -9;
            int? cPlan = RevitHelpers.DiscoverCodeForLabel(_doc, tEl, PlanLabels, -20, 1) ?? -8;
            
            RevitHelpers.SetViewByCode(_doc, tEl, cElevation.Value, elevationViewName);
            RevitHelpers.SetViewByCode(_doc, bEl, cPlan.Value, "Floor Plan");

            List<double> otherOffs;
            double trueMaxOff;
            double effTopOff;

            if (_config.UseHeights)
            {
                int tidInt = sym.Id.IntegerValue;
                var offs = offsetsByType.ContainsKey(tidInt) ? offsetsByType[tidInt] : new List<double>();
                var uniqueOffs = new List<double>();
                
                if (offs.Count > 0)
                {
                    uniqueOffs = offs.Distinct().OrderBy(o => o).ToList();
                }

                if (uniqueOffs.Count > 0)
                {
                    trueMaxOff = uniqueOffs.Max();
                    otherOffs = uniqueOffs.Where(o => o < trueMaxOff).ToList();
                }
                else
                {
                    trueMaxOff = 0.0;
                    otherOffs = new List<double>();
                }
            }
            else
            {
                trueMaxOff = 0.0;
                otherOffs = new List<double>();
            }

            // The gap is now the absolute 0 for the elevations (Baseline). 
            // We push the elevation up by the maximum offset from this baseline.
            effTopOff = (_config.MinBackGap_MM * RevitHelpers.FT_PER_MM) + trueMaxOff;

            // Position elements
            double axisX = RevitHelpers.GetCenterX(bEl, _doc);
            RevitHelpers.MoveVertical(_doc, bEl, rowBotY);
            RevitHelpers.MoveVertical(_doc, tEl, rowBotY + effTopOff);
            RevitHelpers.AlignPairXOnly(_doc, tEl, bEl, axisX);

            return (otherOffs, effTopOff, trueMaxOff);
        }

        private void SetTypeOnLC(Element lc, ElementId symId)
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

            if (p != null)
            {
                try { p.Set(symId); }
                catch { }
            }
        }

        private double EffWidth(Element elTop, Element elBot, double minWFt)
        {
            return Math.Max(
                Math.Max(RevitHelpers.GetWidth(elTop, _doc), RevitHelpers.GetWidth(elBot, _doc)),
                minWFt);
        }

        private void DrawLines(List<PlacedItem> placed, View view, double lineOffFt)
        {
            var rows = new Dictionary<double, (double y, double minx, double maxx)>();

            foreach (var info in placed)
            {
                var topEl = info.TopElement;
                var botEl = info.BotElement;
                var v = _doc.GetElement(topEl.OwnerViewId) as View;

                var (xminT, yminT, xmaxT, ymaxT) = RevitHelpers.GetMinMax(topEl, _doc);
                var (xminB, yminB, xmaxB, ymaxB) = RevitHelpers.GetMinMax(botEl, _doc);

                // Vertical line next to BACK
                double xBack = xmaxT + lineOffFt;
                RevitHelpers.CreateDetailLine(_doc, v, new XYZ(xBack, yminT, 0.0), new XYZ(xBack, ymaxT, 0.0));

                // Vertical line next to FP
                double xPlan = xmaxB + lineOffFt;
                RevitHelpers.CreateDetailLine(_doc, v, new XYZ(xPlan, yminB, 0.0), new XYZ(xPlan, ymaxB, 0.0));

                // Bottom line FP
                double yBottom = yminB - lineOffFt;
                RevitHelpers.CreateDetailLine(_doc, v, new XYZ(xminB, yBottom, 0.0), new XYZ(xmaxB, yBottom, 0.0));

                // Aggregate baseline per row
                double baselineY = yminT - info.TrueMaxOffset;
                double key = Math.Round(baselineY, 6);
                if (!rows.ContainsKey(key))
                {
                    rows[key] = (baselineY, xminT, xmaxT);
                }
                else
                {
                    var existing = rows[key];
                    rows[key] = (existing.y, Math.Min(existing.minx, xminT), Math.Max(existing.maxx, xmaxT));
                }

                // Offset lines (only if heights mode is ON)
                if (_config.UseHeights)
                {
                    var offsetsFt = info.Offsets;
                    double trueMax = info.TrueMaxOffset;
                    
                    // Intermediate offsets only (skip 0.0 and max)
                    var intermediateOffs = offsetsFt.Where(o => o > 1e-4 && Math.Abs(o - trueMax) > 1e-4).ToList();

                    if (intermediateOffs.Count > 0)
                    {
                        // Find a blue line style if possible
                        Element blueStyle = null;
                        try
                        {
                            var categories = _doc.Settings.Categories;
                            var lineCat = categories.get_Item(BuiltInCategory.OST_Lines);
                            if (lineCat != null)
                            {
                                foreach (Category sub in lineCat.SubCategories)
                                {
                                    if (sub.Name.Contains("VH_Blauw") || sub.Name.Equals("Blue", StringComparison.OrdinalIgnoreCase))
                                    {
                                        blueStyle = sub.GetGraphicsStyle(GraphicsStyleType.Projection);
                                        break;
                                    }
                                }
                            }
                        }
                        catch { }

                        foreach (double off in intermediateOffs.OrderBy(o => o))
                        {
                            try
                            {
                                // The tick is placed relative to the baseline
                                double yTick = baselineY + off;
                                
                                // Draw horizontal line across the elevation area
                                var line = RevitHelpers.CreateDetailLine(_doc, v, 
                                    new XYZ(xminT, yTick, 0.0), 
                                    new XYZ(xmaxT, yTick, 0.0));

                                if (line != null && blueStyle != null)
                                {
                                    line.LineStyle = blueStyle;
                                }
                            }
                            catch { }
                        }
                    }
                }
            }

            // Draw zero-line per row
            foreach (var r in rows.Values)
            {
                double y0 = r.y;
                double minx = r.minx - lineOffFt;
                double maxx = r.maxx + lineOffFt;
                RevitHelpers.CreateDetailLine(_doc, view, new XYZ(minx, y0, 0.0), new XYZ(maxx, y0, 0.0));
            }
        }

        private void CreateReferenceCross(List<PlacedItem> placed, View view, 
            double refLeftFt, double refTopFt, double targetWFt, double targetHFt)
        {
            var first = placed[0].TopElement;
            var (xminT, yminT, _, _) = RevitHelpers.GetMinMax(first, _doc);
            double originX = xminT;
            double originY = yminT;
            double crossX = originX + refLeftFt;
            double crossY = originY + refTopFt;
            double halfW = Math.Max(targetWFt / 2.0, 1.0);
            double halfH = Math.Max(targetHFt / 2.0, 1.0);
            RevitHelpers.CreateReferencePlane(_doc, view,
                new XYZ(crossX - halfW, crossY, 0.0),
                new XYZ(crossX + halfW, crossY, 0.0),
                null);
            RevitHelpers.CreateReferencePlane(_doc, view,
                new XYZ(crossX, crossY - halfH, 0.0),
                new XYZ(crossX, crossY + halfH, 0.0),
                null);
        }

        private (List<Line> vLines, List<Line> hLines, BoundingBoxXYZ bbox) GetGeometryInfo(Element el, View view)
        {
            Options opts = new Options
            {
                ComputeReferences = true,
                View = view,
                IncludeNonVisibleObjects = true
            };

            List<Line> vLines = new List<Line>();
            List<Line> hLines = new List<Line>();
            
            GeometryElement geom = el.get_Geometry(opts);
            if (geom != null) ProcessGeometry(geom, vLines, hLines);

            // If it's a Legend Component, the geometry returned might be incomplete.
            // But usually el.get_Geometry(opts) is enough.

            var bbox = el.get_BoundingBox(view);
            if (bbox == null) bbox = RevitHelpers.GetBoundingBox(el, _doc);

            // Deduplicate by position to clean up duplicates from dual geometry processing
            vLines = FilterByPosition(vLines, true);
            hLines = FilterByPosition(hLines, false);

            return (vLines, hLines, bbox);
        }

        // Helper: filter lines by bounding box
        private List<Line> FilterByBoundingBox(List<Line> lines, BoundingBoxXYZ bbox)
        {
            if (lines == null) return new List<Line>();
            if (bbox == null) return lines;
            double tol = 0.05; // ~15mm
            return lines.Where(l =>
                l.Origin.X >= bbox.Min.X - tol && l.Origin.X <= bbox.Max.X + tol &&
                l.Origin.Y >= bbox.Min.Y - tol && l.Origin.Y <= bbox.Max.Y + tol
            ).ToList();
        }

        // Helper: deduplicate lines by position (X for vertical lines, Y for horizontal)
        private List<Line> FilterByPosition(List<Line> lines, bool isVertical)
        {
            if (lines == null) return new List<Line>();
            return lines
                // Prefer lines with references during grouping
                .OrderByDescending(l => l.Reference != null ? 1 : 0)
                .GroupBy(l => Math.Round(isVertical ? l.Origin.X : l.Origin.Y, 4))
                .Select(g => g.First())
                .OrderBy(l => isVertical ? l.Origin.X : l.Origin.Y)
                .ToList();
        }



        private void ProcessGeometry(GeometryElement geomElem, List<Line> vLines, List<Line> hLines)
        {
            foreach (GeometryObject obj in geomElem)
            {
                if (obj is Line line)
                {
                    // Sanity check: ignore lines that are absurdly long (junk geometry)
                    // A window component should not have 1.5km long reference lines.
                    if (line.Length > 50.0) continue; 

                    if (IsInvisibleLine(line))
                    {
                        if (IsVertical(line))
                            vLines.Add(line);
                        else if (IsHorizontal(line))
                            hLines.Add(line);
                    }
                }
                else if (obj is GeometryInstance inst)
                {
                    // Recursively process BOTH instance and symbol geometry. 
                    // Some Legend Components only provide references in one or the other.
                    var instGeom = inst.GetInstanceGeometry();
                    if (instGeom != null) ProcessGeometry(instGeom, vLines, hLines);
                    
                    var symGeom = inst.GetSymbolGeometry();
                    if (symGeom != null) ProcessGeometry(symGeom, vLines, hLines);
                }
            }
        }

        private bool IsInvisibleLine(Line line)
        {
            try
            {
                ElementId gsId = line.GraphicsStyleId;
                if (gsId == ElementId.InvalidElementId) return false;

                if (_doc.GetElement(gsId) is GraphicsStyle gs)
                {
                    string name = gs.Name.ToLower();
                    // Broader match for invisible lines
                    if (name.Contains("inv") || name.Contains("onz") || 
                        gs.GraphicsStyleCategory.Id.IntegerValue == (long)BuiltInCategory.OST_InvisibleLines)
                    {
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool IsVertical(Line line)
        {
            XYZ dir = line.Direction;
            return Math.Abs(dir.X) < 1e-3 && Math.Abs(dir.Y) > 1e-3;
        }

        private bool IsHorizontal(Line line)
        {
            XYZ dir = line.Direction;
            return Math.Abs(dir.Y) < 1e-3 && Math.Abs(dir.X) > 1e-3;
        }
    }
}
