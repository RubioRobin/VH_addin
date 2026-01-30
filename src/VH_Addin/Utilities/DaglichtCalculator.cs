using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using static VH_Tools.Models.Constants;
using static VH_Tools.Utilities.CalculatorHelpers;
using VH_Tools.Models;

namespace VH_Tools.Utilities
{
    public partial class DaglichtCalculator
    {
        private readonly Document _doc;
        private readonly DaglichtOptions _options;
        private readonly List<(Element element, Transform transform)> _allObstructions;

        public DaglichtCalculator(Document doc, DaglichtOptions options)
        {
            _doc = doc;
            _options = options;
            
            // Pre-collect all relevant categories once to avoid repeated collectors
            var builtInCats = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Roofs,
                BuiltInCategory.OST_Ceilings,
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralColumns
            };
            
            var catFilter = new ElementMulticategoryFilter(builtInCats);
            _allObstructions = new List<(Element, Transform)>();

            // Local
            var locals = new FilteredElementCollector(doc)
                .WherePasses(catFilter)
                .WhereElementIsNotElementType()
                .ToElements();
            foreach (var e in locals) _allObstructions.Add((e, Transform.Identity));

            // Links
            var links = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>();
            foreach (var link in links)
            {
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc != null)
                {
                    Transform tr = link.GetTotalTransform();
                    var linkElements = new FilteredElementCollector(linkDoc)
                        .WherePasses(catFilter)
                        .WhereElementIsNotElementType()
                        .ToElements();
                    foreach (var e in linkElements) _allObstructions.Add((e, tr));
                }
            }
        }

        public WindowResult ProcessWindow(
            FamilyInstance window,
            View view,
            GraphicsStyle alphaStyle,
            GraphicsStyle betaStyle,
            GraphicsStyle alpha3DStyle)
        {
            var result = new WindowResult
            {
                ElementId = window.Id.IntegerValue
            };

            try
            {
                // Check if sparingmaker
                if (IsSparingmaker(window))
                {
                    result.Message = "Element is sparingmaker; α/β/glas niet berekend.";
                    result.BetaReason = "Sparingmaker";
                    return result;
                }

                // Get bounding box
                BoundingBoxXYZ bbox = window.get_BoundingBox(null);
                if (bbox == null)
                {
                    result.Message = "Geen bounding box gevonden voor kozijn.";
                    return result;
                }

                // Get window orientation
                if (window.FacingOrientation == null)
                {
                    result.Message = "Element is geen kozijn (geen FacingOrientation).";
                    return result;
                }

                XYZ facing = window.FacingOrientation;
                XYZ horiz = new XYZ(facing.X, facing.Y, 0);
                double L = Math.Sqrt(horiz.X * horiz.X + horiz.Y * horiz.Y);
                
                if (L < 1e-9)
                {
                    result.Message = "FacingOrientation is nulvector.";
                    return result;
                }

                // Standard behavior: Fan points INWARD (Opposite to Facing)
                // User requested "Rotated 180" as standard.
                XYZ baseDir = new XYZ(-horiz.X / L, -horiz.Y / L, 0); 
                XYZ nDir = new XYZ(horiz.X / L, horiz.Y / L, 0);

                // Option: Flip back to OUTWARD (Original)
                if (_options.FlipAlphaFan)
                {
                    baseDir = baseDir.Negate();
                    nDir = nDir.Negate();
                }

                // Calculate glass area if requested
                if (_options.DoGlass)
                {
                    // Pass default sash width from options
                    var gCalc = new GlassAreaCalculator(_doc, _options.DefaultSashWidthMm);
                    double gAreaSqFt = gCalc.Calculate(window, bbox, facing); 
                    
                    result.GlassM2 = gAreaSqFt * SQFT_TO_SQM;
                    
                    // Set parameter in internal units (SqFt)
                    CalculatorHelpers.SetDoubleParam(window, "VH_kozijn_Ad", gAreaSqFt);
                }

                // Get reference height for calculations
                Level level = _doc.GetElement(window.LevelId) as Level;
                
                // Calculate Sash Offset (Operable Windows)
                // USER REQUEST: Always use the UI value for sash width, ignoring parameters
                double sashOffsetFt = GetSashOffset(window);
                
                // Pass sash offset and UI defaults to Helper to ensure they're applied BEFORE Max(600) check
                double defFrameFt = _options.DefaultFrameHeightMm * MM_TO_FT;
                double defGlassOffFt = _options.DefaultGlassOffsetMm * MM_TO_FT;
                
                var (zRef, zRefMm) = GetGlazingRefHeight(window, bbox, level, sashOffsetFt, defFrameFt, defGlassOffFt);

                double centerX = 0.5 * (bbox.Min.X + bbox.Max.X);
                double centerY = 0.5 * (bbox.Min.Y + bbox.Max.Y);
                XYZ centerXY = new XYZ(centerX, centerY, zRef);

                // 1. Find walls that could provide the interior face for the window starting point.
                // We keep this broad (500mm) to catch insulation/inner leaves of multi-layer walls.
                double faceSearchTolerance = 500.0 * MM_TO_FT; 
                var faceCandidateWalls = _allObstructions
                    .Where(o => o.element is Wall)
                    .Select(o => (Wall)o.element)
                    .Where(w => BBoxesIntersectWithTolerance(bbox, w.get_BoundingBox(null), faceSearchTolerance))
                    .ToList();

                XYZ start = centerXY;
                if (faceCandidateWalls.Any())
                {
                    start = GetStartOnInsideFace(faceCandidateWalls, centerXY, zRef, nDir);
                }

                // 2. Determine which walls should be EXCLUDED from being obstructions (self-shadowing).
                // Only exclude walls that are truly part of the host assembly (parallel AND intersecting tightly).
                var assemblyWallIds = new HashSet<long>();
                if (window.Host != null) assemblyWallIds.Add(window.Host.Id.IntegerValue);

                foreach (var w in faceCandidateWalls)
                {
                    // Use a very tight tolerance (10mm) for "part of same assembly"
                    if (BBoxesIntersectWithTolerance(bbox, w.get_BoundingBox(null), 10.0 * MM_TO_FT))
                    {
                        // Check if wall is roughly parallel to window facing
                        if (w.Location is LocationCurve lc && lc.Curve is Line line)
                        {
                            XYZ wallDir = line.Direction.Normalize();
                            // Parallel if wall direction is perpendicular to window facing
                            if (Math.Abs(wallDir.DotProduct(facing)) < 0.1)
                            {
                                assemblyWallIds.Add(w.Id.IntegerValue);
                            }
                        }
                        else
                        {
                            // Fallback for curved walls
                            assemblyWallIds.Add(w.Id.IntegerValue);
                        }
                    }
                }

                Wall hostWall = window.Host as Wall;

                // Calculate alpha if requested
                if (_options.DoAlpha)
                {
                    CalculateAlpha(window, view, bbox, baseDir, nDir, start, zRef, zRefMm, 
                                 hostWall, assemblyWallIds, alphaStyle, alpha3DStyle, result);
                }

                // Calculate beta if requested
                if (_options.DoBeta)
                {
                    CalculateBeta(window, bbox, baseDir, hostWall, betaStyle, result, start, assemblyWallIds);
                }

                // NEN 2057 Check: Alpha and Beta overlap
                // If the Alpha angle (to external obstruction) is effectively blocked by the overhang (Beta obstruction),
                // then Alpha is determined by the underside of the overhang.
                if (_options.DoAlpha && _options.DoBeta && result.BetaObstructPoint != null && result.AlphaAvgDeg.HasValue)
                {
                    // Calculate angle from P (start) to the overhang tip
                    XYZ p = result.BetaObstructPoint;
                    double dx = p.X - start.X;
                    double dy = p.Y - start.Y;
                    double distHor = Math.Sqrt(dx * dx + dy * dy);
                    double dz = p.Z - start.Z;

                    if (distHor > 1e-9) // Determine angle from horizon at P
                    {
                        double angleToOverhang = Math.Atan2(dz, distHor) * 180.0 / Math.PI;
                        
                        // If current Alpha is "higher" (steeper) than the angle to the overhang,
                        // it means the overhang is blocking the view to the external obstruction.
                        // So we cap Alpha to the overhang angle.
                        // NEN 2057: Min alpha is 20 degrees.
                        
                        if (result.AlphaAvgDeg.Value > angleToOverhang)
                        {
                            double newAlpha = Math.Max(20.0, angleToOverhang);
                            result.AlphaAvgDeg = newAlpha;
                            // Update parameter
                            CalculatorHelpers.SetDoubleParam(window, PARAM_ALPHA, newAlpha);
                        }
                    }
                }

                // Add diagnostic info to message
                string debugInfo = $"P={zRefMm:0}mm (sash={sashOffsetFt * FOOT_TO_MM:0})";
                result.Message = $"{debugInfo} | α {(_options.DoAlpha ? "aan" : "uit")} / β {(_options.DoBeta ? "aan" : "uit")}";
                
                return result;
            }
            catch (Exception ex)
            {
                result.Message = $"Fout in compute_fan_and_beta: {ex.Message}";
                return result;
            }
        }

        private void CalculateAlpha(
            FamilyInstance window,
            View view,
            BoundingBoxXYZ bbox,
            XYZ baseDir,
            XYZ nDir,
            XYZ start,
            double zRef,
            double zRefMm,
            Wall hostWall,
            HashSet<long> assemblyWallIds,
            GraphicsStyle alphaStyle,
            GraphicsStyle alpha3DStyle,
            WindowResult result)
        {
            var alphaVals = new List<double>();

            // Narrow search to 5m radius (LINE_LENGTH_FEET) using spatial filter
            var candidates = _allObstructions
                .Where(o => o.element is Wall && !assemblyWallIds.Contains(o.element.Id.IntegerValue))
                .Select(o => {
                    BoundingBoxXYZ wb = o.element.get_BoundingBox(null);
                    if (wb == null) return (element: (Wall)null, bbox: (BoundingBoxXYZ)null);
                    
                    // Transform BBox
                    XYZ min = o.transform.OfPoint(wb.Min);
                    XYZ max = o.transform.OfPoint(wb.Max);
                    BoundingBoxXYZ trBbox = new BoundingBoxXYZ { Min = min, Max = max };
                    
                    return (element: (Wall)o.element, bbox: trBbox);
                })
                .Where(c => c.element != null &&
                           c.bbox.Min.X < start.X + LINE_LENGTH_FEET && c.bbox.Max.X > start.X - LINE_LENGTH_FEET &&
                           c.bbox.Min.Y < start.Y + LINE_LENGTH_FEET && c.bbox.Max.Y > start.Y - LINE_LENGTH_FEET)
                .ToList();

            SketchPlane sketchPlaneH = null;
            if (_options.DrawLines)
            {
                Plane planeH = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, start);
                sketchPlaneH = SketchPlane.Create(_doc, planeH);
            }

            for (int i = 0; i < RAY_COUNT; i++)
            {
                double angleDeg = ANGLE_START + i * ANGLE_STEP;
                XYZ dirI = RotateXY(baseDir, angleDeg * Math.PI / 180.0);

                double? bestT = null;
                Wall bestWallHit = null;

                foreach (var cand in candidates)
                {
                    double? t = RayBox2DDistance(start, dirI, cand.bbox);
                    if (t.HasValue && t.Value > 0 && t.Value <= LINE_LENGTH_FEET)
                    {
                        if (!bestT.HasValue || t.Value < bestT.Value)
                        {
                            bestT = t;
                            bestWallHit = cand.element;
                        }
                    }
                }

                double alpha = 20.0;
                double lineLen = LINE_LENGTH_FEET;
                int? obstacleId = null;
                XYZ hitPt = null;
                XYZ footPt = null;
                double? xDist = null;
                int? alpha3DId = null;
                int? x3DId = null;

                if (bestWallHit != null && bestT.HasValue)
                {
                    double dHoriz = bestT.Value;
                    BoundingBoxXYZ wbHit = bestWallHit.get_BoundingBox(null);
                    double zObst = wbHit != null ? wbHit.Max.Z : zRef;
                    lineLen = Math.Min(LINE_LENGTH_FEET, dHoriz);
                    obstacleId = bestWallHit.Id.IntegerValue;

                    hitPt = new XYZ(
                        start.X + dirI.X * dHoriz,
                        start.Y + dirI.Y * dHoriz,
                        zRef
                    );

                    XYZ vHit = hitPt - start;
                    double distN = vHit.X * nDir.X + vHit.Y * nDir.Y;
                    
                    if (Math.Abs(distN) > 1e-6)
                    {
                        xDist = Math.Abs(distN);
                        footPt = hitPt - nDir * distN;
                        alpha = CalculateAlphaFromX(zRef, zObst, xDist.Value);
                    }

                    // Draw 3D X and A lines
                    if (_options.DrawLines && footPt != null)
                    {
                        try
                        {
                            XYZ base3D = new XYZ(footPt.X, footPt.Y, zRef);
                            XYZ outer3D = new XYZ(hitPt.X, hitPt.Y, zRef);
                            XYZ top3D = new XYZ(hitPt.X, hitPt.Y, zObst);

                            XYZ normal = nDir.CrossProduct(XYZ.BasisZ);
                            if (normal.GetLength() < 1e-6)
                                normal = XYZ.BasisX;

                            Plane planeA = Plane.CreateByNormalAndOrigin(normal, base3D);
                            SketchPlane spA = SketchPlane.Create(_doc, planeA);

                            // X line (horizontal)
                            Line lineX3D = Line.CreateBound(base3D, outer3D);
                            ModelCurve mcX3D = _doc.Create.NewModelCurve(lineX3D, spA);
                            x3DId = mcX3D.Id.IntegerValue;
                            if (alpha3DStyle != null)
                                mcX3D.LineStyle = alpha3DStyle;

                            // A line (vertical)
                            Line lineA3D = Line.CreateBound(base3D, top3D);
                            ModelCurve mcA3D = _doc.Create.NewModelCurve(lineA3D, spA);
                            alpha3DId = mcA3D.Id.IntegerValue;
                            if (alpha3DStyle != null)
                                mcA3D.LineStyle = alpha3DStyle;
                        }
                        catch { }
                    }
                }

                int? mcId = null;
                if (_options.DrawLines && sketchPlaneH != null)
                {
                    try
                    {
                        XYZ p0 = start;
                        XYZ p1 = new XYZ(
                            start.X + dirI.X * lineLen,
                            start.Y + dirI.Y * lineLen,
                            zRef
                        );
                        Line mcLine = Line.CreateBound(p0, p1);
                        ModelCurve mc = _doc.Create.NewModelCurve(mcLine, sketchPlaneH);
                        mcId = mc.Id.IntegerValue;
                        if (alphaStyle != null)
                            mc.LineStyle = alphaStyle;
                    }
                    catch { }
                }

                result.Rays.Add(new RayResult
                {
                    Index = i,
                    AngleOffsetDeg = angleDeg,
                    AlphaDeg = alpha,
                    XDistMm = xDist.HasValue ? xDist.Value * FOOT_TO_MM : (double?)null,
                    DHorizMm = bestT.HasValue ? bestT.Value * FOOT_TO_MM : (double?)null,
                    ZRefMm = zRefMm,
                    ZObstMm = bestWallHit != null ? (bestWallHit.get_BoundingBox(null)?.Max.Z ?? zRef) * FOOT_TO_MM : (double?)null,
                    LineLengthMm = lineLen * FOOT_TO_MM,
                    ObstacleId = obstacleId,
                    ModelCurveId = mcId,
                    Alpha3DModelCurveId = alpha3DId,
                    X3DModelCurveId = x3DId,
                    SketchPlaneHId = sketchPlaneH?.Id.IntegerValue
                });

                alphaVals.Add(alpha);
            }

            if (alphaVals.Any())
            {
                double alphaAvg = alphaVals.Average();
                result.AlphaAvgDeg = alphaAvg;
                SetAlphaParameter(window, alphaAvg);
            }
        }

        private double GetSashOffset(FamilyInstance window)
        {
            if (!IsOperableWindow(window))
                return 0;

            // USER REQUEST: Ignore breedte_raamhout parameter and always use UI value
            return _options.DefaultSashWidthMm * MM_TO_FT;
        }

        private bool IsOperableWindow(FamilyInstance window)
        {
            var names = new List<string> { window.Name ?? "" };
            if (window.Symbol != null) names.Add(window.Symbol.Name ?? "");
            if (window.Symbol?.Family != null) names.Add(window.Symbol.Family.Name ?? "");

            // Check specific VH parameters for nested components (Rows A, B, C; columns 1-4)
            string[] rows = { "A", "B", "C" };
            foreach (var r in rows)
            {
                for (int i = 1; i <= 4; i++)
                {
                    string pName = $"VH_vlakvulling_{r}{i}";
                    string val = CalculatorHelpers.GetStringParam(window, pName);
                    if (!string.IsNullOrEmpty(val))
                    {
                        names.Add(val);
                    }
                }
            }

            // Also check plain "vlakvulling"
            string v = CalculatorHelpers.GetStringParam(window, "vlakvulling");
            if (!string.IsNullOrEmpty(v)) names.Add(v);

            string combined = string.Join(" ", names).ToLower();

            // Broadened keywords implying operability (Dutch + English)
            // REMOVED generic terms like "raam", "kozijn", "window" which caused false positives on fixed windows.
            string[] opTokens = { 
                "draai", "kiep", "val", "schuif", "stolp", "dk", 
                "openend", "vleugel", // "opening" removed as it can match "wandopening"
                "sash", "operable", "vent" 
            };
            bool isOp = opTokens.Any(t => combined.Contains(t));
            
            return isOp;
        }
    }
}
