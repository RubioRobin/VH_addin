using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace VH_DaglichtPlugin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DaglichtCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (doc == null)
            {
                TaskDialog.Show("Error", "Geen actief document gevonden.");
                return Result.Failed;
            }

            try
            {
                // Show UI and get user options
                var dlg = new DaglichtWindow();
                bool? result = dlg.ShowDialog();

                if (result != true)
                {
                    return Result.Cancelled;
                }

                var options = dlg.GetOptions();

                // Collect windows based on selection mode
                List<FamilyInstance> windows = CollectWindows(doc, uidoc, options);

                if (windows == null || windows.Count == 0)
                {
                    TaskDialog.Show("Info", "Geen kozijnen gevonden voor de gekozen selectie.");
                    return Result.Cancelled;
                }

                // Filter by VG areas if requested
                if (options.OnlyVG)
                {
                    var vgAreas = CollectVGAreas(doc);
                    if (vgAreas.Any())
                    {
                        windows = windows.Where(w => WindowTouchesVGArea(w, vgAreas, doc.ActiveView)).ToList();
                    }
                }

                if (windows.Count == 0)
                {
                    TaskDialog.Show("Info", "Geen kozijnen gevonden na VG-filtering.");
                    return Result.Cancelled;
                }

                // Start transaction
                using (Transaction trans = new Transaction(doc, "VH Daglichttool"))
                {
                    var failOptions = trans.GetFailureHandlingOptions();
                    failOptions.SetFailuresPreprocessor(new WarningSwallower());
                    trans.SetFailureHandlingOptions(failOptions);

                    trans.Start();

                    try
                    {
                        // Cleanup existing lines if drawing new ones
                        if (options.DrawLines)
                        {
                            CleanupVHLines(doc);
                        }

                        // Create line styles
                        GraphicsStyle alphaStyle = null;
                        GraphicsStyle betaStyle = null;
                        GraphicsStyle alpha3DStyle = null;

                        if (options.DrawLines)
                        {
                            alphaStyle = GetOrCreateLineStyle(doc, "VH_Alpha", new Color(255, 0, 0));
                            betaStyle = GetOrCreateLineStyle(doc, "VH_Beta", new Color(0, 0, 255));
                            alpha3DStyle = alphaStyle; // Same style for 3D lines
                        }

                        if (windows == null || windows.Count == 0)
                        {
                            TaskDialog.Show("Info", "Geen kozijnen gevonden voor de gekozen selectie.");
                            return Result.Cancelled;
                        }

                        // Process each window
                        var calculator = new DaglichtCalculator(doc, options);
                        var results = new List<WindowResult>();

                        foreach (var window in windows)
                        {
                            var windowResult = calculator.ProcessWindow(
                                window, 
                                doc.ActiveView, 
                                alphaStyle, 
                                betaStyle, 
                                alpha3DStyle
                            );
                            
                            if (windowResult != null)
                            {
                                results.Add(windowResult);
                            }
                        }

                        // If NEN calculation was requested, run the second phase
                        if (options.DoGlass)
                        {
                            doc.Regenerate(); // Ensure values set in previous step are available
                            RunNENCalculation(doc, windows, options.DoExport);
                        }

                        trans.Commit();

                        // Show results summary
                        string summary = $"Verwerkt: {results.Count} kozijnen\n\n";
                        
                        if (options.DoAlpha)
                        {
                            int alphaCount = results.Count(r => r.AlphaAvgDeg.HasValue);
                            summary += $"α berekend: {alphaCount}\n";
                        }
                        
                        if (options.DoBeta)
                        {
                            int betaCount = results.Count(r => r.BetaDeg.HasValue);
                            summary += $"β berekend: {betaCount}\n";
                        }
                        
                        if (options.DoGlass)
                        {
                            int glassCount = results.Count(r => r.GlassM2.HasValue);
                            summary += $"Glasoppervlakte berekend: {glassCount}\n";
                        }

                        // TaskDialog.Show("VH Daglichttool", summary);

                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        TaskDialog.Show("Error", $"Fout tijdens berekening:\n{ex.Message}\n\n{ex.StackTrace}");
                        return Result.Failed;
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Fout:\n{ex.Message}");
                return Result.Failed;
            }
        }

        private bool IsWindow(FamilyInstance fi)
        {
            return fi.Category != null && fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows;
        }

        private List<FamilyInstance> CollectWindows(Document doc, UIDocument uidoc, DaglichtOptions options)
        {
            var windows = new List<FamilyInstance>();

            if (options.SelectionOnly)
            {
                var selIds = uidoc.Selection.GetElementIds();
                foreach (ElementId id in selIds)
                {
                    var elem = doc.GetElement(id);
                    if (elem is FamilyInstance fi && IsWindow(fi))
                    {
                        windows.Add(fi);
                    }
                }
            }
            else if (options.CurrentLevelOnly)
            {
                Level level = doc.ActiveView.GenLevel;
                if (level != null)
                {
                    var collector = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .OfCategory(BuiltInCategory.OST_Windows)
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>();

                    windows = collector.Where(w => IsWindow(w) && w.LevelId == level.Id).ToList();
                }
            }
            else if (options.ActiveViewOnly)
            {
                windows = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_Windows)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(w => IsWindow(w))
                    .ToList();
            }
            else if (options.AllModelOnly)
            {
                windows = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Windows)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(w => IsWindow(w))
                    .ToList();
            }

            return windows;
        }

        private List<SpatialElement> CollectVGAreas(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Areas)
                .WhereElementIsNotElementType()
                .Cast<SpatialElement>()
                .ToList();
        }

        private bool WindowTouchesVGArea(FamilyInstance window, List<SpatialElement> areas, View view)
        {
            BoundingBoxXYZ bboxW = window.get_BoundingBox(view) ?? window.get_BoundingBox(null);
            if (bboxW == null) return false;

            foreach (var area in areas)
            {
                BoundingBoxXYZ bboxA = area.get_BoundingBox(view) ?? area.get_BoundingBox(null);
                if (bboxA != null && BBoxesIntersect(bboxW, bboxA))
                {
                    return true;
                }
            }
            return false;
        }

        private bool BBoxesIntersect(BoundingBoxXYZ b1, BoundingBoxXYZ b2)
        {
            double xOverlap = Math.Min(b1.Max.X, b2.Max.X) - Math.Max(b1.Min.X, b2.Min.X);
            double yOverlap = Math.Min(b1.Max.Y, b2.Max.Y) - Math.Max(b1.Min.Y, b2.Min.Y);
            return xOverlap > 0 && yOverlap > 0;
        }

        private void CleanupVHLines(Document doc)
        {
            var categories = doc.Settings.Categories;
            var linesCat = categories.get_Item(BuiltInCategory.OST_Lines);
            var subcats = linesCat.SubCategories;

            var styleIds = new HashSet<ElementId>();
            string[] styleNames = { "VH_Alpha", "VH_Beta", "VH_Alpha_X" };

            foreach (string name in styleNames)
            {
                try
                {
                    var subcat = subcats.get_Item(name);
                    if (subcat != null)
                    {
                        var gs = subcat.GetGraphicsStyle(GraphicsStyleType.Projection);
                        if (gs != null)
                        {
                            styleIds.Add(gs.Id);
                        }
                    }
                }
                catch { }
            }

            if (styleIds.Count > 0)
            {
                var collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(CurveElement));

                var toDelete = new List<ElementId>();
                foreach (CurveElement ce in collector)
                {
                    try
                    {
                        var ls = ce.LineStyle;
                        if (ls is GraphicsStyle gs && styleIds.Contains(gs.Id))
                        {
                            toDelete.Add(ce.Id);
                        }
                    }
                    catch { }
                }

                if (toDelete.Count > 0)
                {
                    doc.Delete(toDelete);
                }
            }
        }

        private GraphicsStyle GetOrCreateLineStyle(Document doc, string name, Color color)
        {
            var categories = doc.Settings.Categories;
            var linesCat = categories.get_Item(BuiltInCategory.OST_Lines);
            var subcats = linesCat.SubCategories;

            Category subcat = null;
            try
            {
                subcat = subcats.get_Item(name);
            }
            catch
            {
                subcat = categories.NewSubcategory(linesCat, name);
            }

            if (subcat != null)
            {
                subcat.LineColor = color;
                return subcat.GetGraphicsStyle(GraphicsStyleType.Projection);
            }

            return null;
        }

        private void RunNENCalculation(Document doc, List<FamilyInstance> windows, bool doExport)
        {
            // This will be implemented in the NEN calculator class
            var nenCalc = new NENCalculator(doc);
            nenCalc.Calculate(windows, doExport);
        }
    }
}
