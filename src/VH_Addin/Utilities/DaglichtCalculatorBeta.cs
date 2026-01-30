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
        private void CalculateBeta(
            FamilyInstance window,
            BoundingBoxXYZ bbox,
            XYZ nDir,
            Wall hostWall,
            GraphicsStyle betaStyle,
            WindowResult result,
            XYZ startPtAlpha,
            HashSet<long> assemblyWallIds)
        {
            var (bEff, glassBottomAbs, glassTopAbs) = GetGlassBounds(window, bbox);
            double zMidEff = 0.5 * (bEff + glassTopAbs);

            double centerX = 0.5 * (bbox.Min.X + bbox.Max.X);
            double centerY = 0.5 * (bbox.Min.Y + bbox.Max.Y);
            XYZ centerXY = new XYZ(centerX, centerY, zMidEff);

            XYZ startPt = new XYZ(startPtAlpha.X, startPtAlpha.Y, zMidEff);

            // Get t-vector (perpendicular to n in XY plane)
            XYZ tVec = XYZ.BasisZ.CrossProduct(nDir);
            double Lt = Math.Sqrt(tVec.X * tVec.X + tVec.Y * tVec.Y + tVec.Z * tVec.Z);
            if (Lt < 1e-9)
            {
                tVec = XYZ.BasisX;
                Lt = 1;
            }
            XYZ t = new XYZ(tVec.X / Lt, tVec.Y / Lt, tVec.Z / Lt);

            // Select candidates within 5m (MAX_B_DISTANCE) using rough spatial check
            var candidatesByDist = _allObstructions
                .Where(o => o.element.Id.IntegerValue != window.Id.IntegerValue)
                .Select(o => {
                    BoundingBoxXYZ ob = o.element.get_BoundingBox(null);
                    if (ob == null) return (element: (Element)null, transform: (Transform)null, bbox: (BoundingBoxXYZ)null);
                    
                    // Transform BBox for spatial check
                    XYZ min = o.transform.OfPoint(ob.Min);
                    XYZ max = o.transform.OfPoint(ob.Max);
                    BoundingBoxXYZ trBbox = new BoundingBoxXYZ { Min = min, Max = max };
                    
                    return (element: o.element, transform: o.transform, bbox: trBbox);
                })
                .Where(c => c.element != null &&
                            c.bbox.Min.X < startPt.X + MAX_B_DISTANCE && c.bbox.Max.X > startPt.X - MAX_B_DISTANCE &&
                            c.bbox.Min.Y < startPt.Y + MAX_B_DISTANCE && c.bbox.Max.Y > startPt.Y - MAX_B_DISTANCE &&
                            c.bbox.Max.Z > zMidEff)
                .ToList();

            var candidates = new List<(XYZ point, Element elem)>();

            foreach (var cand in candidatesByDist)
            {
                var (corners, edges) = GetEdgePointsForElement(cand.element);
                if (corners == null || edges == null || corners.Count == 0)
                    continue;

                // Transform corners to host project coordinates
                var trCorners = corners.Select(p => cand.transform.OfPoint(p)).ToList();

                // Calculate t-values for all corners
                var tVals = new List<double>();
                foreach (var corner in trCorners)
                {
                    XYZ v = corner - startPt;
                    double tVal = v.X * t.X + v.Y * t.Y + v.Z * t.Z;
                    tVals.Add(tVal);
                }

                // Check if all corners are on one side
                bool allPos = tVals.All(tv => tv > 1e-9);
                bool allNeg = tVals.All(tv => tv < -1e-9);
                if (allPos || allNeg)
                    continue;

                // Find intersection points
                for (int edgeIdx = 0; edgeIdx < edges.Count; edgeIdx++)
                {
                    var (i, j) = edges[edgeIdx];
                    double ti = tVals[i];
                    double tj = tVals[j];
                    XYZ pi = trCorners[i];
                    XYZ pj = trCorners[j];

                    var rawPoints = new List<XYZ>();

                    if (Math.Abs(ti) <= 1e-9 && Math.Abs(tj) <= 1e-9)
                    {
                        rawPoints.Add(pi);
                        rawPoints.Add(pj);
                    }
                    else if (Math.Abs(ti) <= 1e-9 && Math.Abs(tj) > 1e-9)
                    {
                        rawPoints.Add(pi);
                    }
                    else if (Math.Abs(tj) <= 1e-9 && Math.Abs(ti) > 1e-9)
                    {
                        rawPoints.Add(pj);
                    }
                    if (ti * tj <= 0)
                    {
                        double denom = tj - ti;
                        if (Math.Abs(denom) > 1e-9)
                        {
                            double s = -ti / denom;
                            if (s >= -0.01 && s <= 1.01)
                            {
                                XYZ intersect = new XYZ(
                                    pi.X + (pj.X - pi.X) * s,
                                    pi.Y + (pj.Y - pi.Y) * s,
                                    pi.Z + (pj.Z - pi.Z) * s
                                );
                                rawPoints.Add(intersect);
                            }
                        }
                    }

                    foreach (var p in rawPoints)
                    {
                        XYZ v = p - startPt;
                        double d = v.X * nDir.X + v.Y * nDir.Y + v.Z * nDir.Z;
                        
                        if (d > 0 && d <= MAX_B_DISTANCE)
                        {
                            candidates.Add((p, cand.element));
                            
                            // Update effective top if needed
                            if (p.Z > bEff + 1e-9 && p.Z < glassTopAbs - 1e-9)
                            {
                                glassTopAbs = p.Z;
                            }
                        }
                    }
                }
            }

            if (candidates.Count == 0)
            {
                result.BetaDeg = null;
                result.BetaRad = null;
                result.BetaReason = "Geen belemmeringen in doorsnedevlak (≤5 m) gevonden.";
                return;
            }

            // Recalculate with effective top
            zMidEff = 0.5 * (bEff + glassTopAbs);
            startPt = new XYZ(startPtAlpha.X, startPtAlpha.Y, zMidEff);

            double? bestBetaRad = null;
            double? bestD = null;
            double? bestH = null;
            XYZ bestP = null;
            Element bestElem = null;
            string bestReason = null;

            foreach (var (p, elem) in candidates)
            {
                double d = (p.X - startPt.X) * nDir.X + (p.Y - startPt.Y) * nDir.Y;
                if (d <= 0 || d > MAX_B_DISTANCE)
                    continue;

                double h = p.Z - zMidEff;
                if (h <= 0)
                    continue;

                double betaRad = Math.Atan2(d, h);
                if (!bestBetaRad.HasValue || betaRad > bestBetaRad.Value)
                {
                    bestBetaRad = betaRad;
                    bestD = d;
                    bestH = h;
                    bestP = p;
                    bestElem = elem;

                    if (elem is Wall)
                        bestReason = "Wand / sparing (geometrie)";
                    else if (elem.Category != null)
                        bestReason = elem.Category.Name;
                    else
                        bestReason = "Belemmering";
                }
            }

            if (bestBetaRad.HasValue)
            {
                double betaDeg = bestBetaRad.Value * 180.0 / Math.PI;
                result.BetaDeg = betaDeg;
                result.BetaRad = bestBetaRad.Value;
                result.BetaReason = bestReason;
                result.BetaObstructPoint = bestP;

                // Set parameter
                try
                {
                    Parameter pBeta = window.LookupParameter(PARAM_BETA_EPS)
                                      ?? window.LookupParameter("VH_kozijn_β")
                                      ?? window.LookupParameter("VH_kozijn_ε");

                    if (pBeta != null && !pBeta.IsReadOnly && pBeta.StorageType == StorageType.Double)
                    {
                        pBeta.Set(bestBetaRad.Value);
                    }
                }
                catch { }

                // Draw beta line
                if (_options.DrawLines && bestD.HasValue && bestH.HasValue)
                {
                    result.BetaModelCurveId = DrawBetaLine(startPt, nDir, bestD.Value, bestH.Value, betaStyle);
                }
            }
            else
            {
                result.BetaReason = "Geen belemmeringen boven midden effectieve glas gevonden.";
            }
        }

        private (double bEff, double glassBottomAbs, double glassTopAbs) GetGlassBounds(
            FamilyInstance window,
            BoundingBoxXYZ bbox)
        {
            Level level = _doc.GetElement(window.LevelId) as Level;
            double zLevelFt = level?.Elevation ?? 0;

            double sillFt = GetSillHeight(window);

            bool hasAlu = GetDoubleParam(window, AL_SIDE).HasValue ||
                          GetDoubleParam(window, AL_TOP_BOT).HasValue ||
                          GetDoubleParam(window, AL_OFF_SIDE).HasValue;

            double glassBottomAbsRaw, glassTopAbsRaw;

            if (hasAlu)
            {
                double offStijlFt = GetDoubleParam(window, AL_OFF_SIDE) ?? 0;
                double extraUnderFt = GetDoubleParam(window, AL_EXTRA_UNDER) ?? 0;
                double viewSillFt = GetDoubleParam(window, AL_VIEW_SILL) ?? GetDoubleParam(window, AL_TOP_BOT) ?? 0;

                double glassBottomRelRaw = sillFt - offStijlFt + extraUnderFt + viewSillFt;
                glassBottomAbsRaw = zLevelFt + glassBottomRelRaw;

                double topProfileFt = viewSillFt > 0 ? viewSillFt : offStijlFt;
                glassTopAbsRaw = bbox.Max.Z - topProfileFt;
            }
            else
            {
                double tBotFt = GetDoubleParam(window, P_BOT) ?? 0;
                double glassBottomRelRaw = sillFt + tBotFt + (17.0 * MM_TO_FT);
                glassBottomAbsRaw = zLevelFt + glassBottomRelRaw;
                glassTopAbsRaw = bbox.Max.Z - tBotFt;
            }

            double minRefAbs = zLevelFt + FT_600;
            double bEff = Math.Max(glassBottomAbsRaw, minRefAbs);

            if (glassTopAbsRaw <= bEff)
            {
                glassBottomAbsRaw = bbox.Min.Z;
                glassTopAbsRaw = bbox.Max.Z;
                bEff = bbox.Min.Z;
            }

            return (bEff, glassBottomAbsRaw, glassTopAbsRaw);
        }

        private (List<XYZ> corners, List<(int, int)> edges) GetEdgePointsForElement(Element elem)
        {
            var corners = new List<XYZ>();
            var edges = new List<(int, int)>();
            var pointMap = new Dictionary<string, int>();

            Options opt = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine
            };

            int GetPointIndex(XYZ p)
            {
                string key = $"{Math.Round(p.X, 5)}_{Math.Round(p.Y, 5)}_{Math.Round(p.Z, 5)}";
                if (pointMap.TryGetValue(key, out int idx)) return idx;
                int newIdx = corners.Count;
                corners.Add(p);
                pointMap[key] = newIdx;
                return newIdx;
            }

            void AddGeometry(GeometryElement geom)
            {
                if (geom == null) return;
                foreach (GeometryObject gObj in geom)
                {
                    if (gObj is Solid solid && solid.Volume > 1e-6)
                    {
                        foreach (Edge edge in solid.Edges)
                        {
                            Curve c = edge.AsCurve();
                            int i = GetPointIndex(c.GetEndPoint(0));
                            int j = GetPointIndex(c.GetEndPoint(1));
                            edges.Add((i, j));
                        }
                        foreach (Face face in solid.Faces)
                        {
                            Mesh mesh = face.Triangulate();
                            if (mesh == null) continue;
                            for (int k = 0; k < mesh.NumTriangles; k++)
                            {
                                MeshTriangle tri = mesh.get_Triangle(k);
                                int i = GetPointIndex(tri.get_Vertex(0));
                                int j = GetPointIndex(tri.get_Vertex(1));
                                int l = GetPointIndex(tri.get_Vertex(2));
                                edges.Add((i, j));
                                edges.Add((j, l));
                                edges.Add((l, i));
                            }
                        }
                    }
                    else if (gObj is Mesh mesh)
                    {
                        for (int k = 0; k < mesh.NumTriangles; k++)
                        {
                            MeshTriangle tri = mesh.get_Triangle(k);
                            int i = GetPointIndex(tri.get_Vertex(0));
                            int j = GetPointIndex(tri.get_Vertex(1));
                            int l = GetPointIndex(tri.get_Vertex(2));
                            edges.Add((i, j));
                            edges.Add((j, l));
                            edges.Add((l, i));
                        }
                    }
                    else if (gObj is GeometryInstance gi)
                    {
                        AddGeometry(gi.GetInstanceGeometry());
                    }
                }
            }

            AddGeometry(elem.get_Geometry(opt));

            if (corners.Any())
            {
                // Deduplicate edges to keep it lean
                var distinctEdges = edges
                    .Select(e => e.Item1 < e.Item2 ? (e.Item1, e.Item2) : (e.Item2, e.Item1))
                    .Distinct()
                    .ToList();
                return (corners, distinctEdges);
            }

            // Fallback to bounding box
            BoundingBoxXYZ bb = elem.get_BoundingBox(null);
            if (bb == null) return (null, null);

            corners = new List<XYZ>
            {
                new XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z), new XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
                new XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z), new XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
                new XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z), new XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
                new XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z), new XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z)
            };
            edges = new List<(int, int)> { (0,1),(0,2),(1,3),(2,3), (4,5),(4,6),(5,7),(6,7), (0,4),(1,5),(2,6),(3,7) };
            return (corners, edges);
        }

        private int? DrawBetaLine(XYZ startPt, XYZ n, double d, double h, GraphicsStyle betaStyle)
        {
            if (!_options.DrawLines || startPt == null || n == null || d <= 0 || h <= 0)
                return null;

            try
            {
                XYZ vHor = n * d;
                XYZ vVer = new XYZ(0, 0, h);
                XYZ end = startPt + vHor + vVer;

                XYZ normal = n.CrossProduct(XYZ.BasisZ);
                if (normal.GetLength() < 1e-6)
                    normal = XYZ.BasisX;

                Plane plane = Plane.CreateByNormalAndOrigin(normal, startPt);
                SketchPlane sp = SketchPlane.Create(_doc, plane);
                Line line = Line.CreateBound(startPt, end);
                ModelCurve mc = _doc.Create.NewModelCurve(line, sp);
                
                if (betaStyle != null)
                    mc.LineStyle = betaStyle;

                return mc.Id.IntegerValue;
            }
            catch
            {
                return null;
            }
        }
    }
}
