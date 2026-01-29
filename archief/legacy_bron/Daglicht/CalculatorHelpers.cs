using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using static VH_DaglichtPlugin.Constants;

namespace VH_DaglichtPlugin
{
    public static class CalculatorHelpers
    {
        public static bool IsSparingmaker(FamilyInstance window)
        {
            // Check VH_categorie parameter
            Parameter catParam = window.LookupParameter(PARAM_CAT);
            if (catParam != null && catParam.StorageType == StorageType.String)
            {
                string val = catParam.AsString()?.ToLower() ?? "";
                if (val.Contains("sparing"))
                    return true;
            }

            // Check names
            string[] names = new string[3];
            names[0] = window.Name ?? "";
            
            ElementType type = window.Document.GetElement(window.GetTypeId()) as ElementType;
            if (type != null)
            {
                names[1] = type.Name ?? "";
                if (type is FamilySymbol fs && fs.Family != null)
                {
                    names[2] = fs.Family.Name ?? "";
                }
            }

            string combined = string.Join(" ", names).ToLower();
            return combined.Contains("sparing");
        }

        public static (double zRefFt, double zRefMm) GetGlazingRefHeight(
            FamilyInstance window,
            BoundingBoxXYZ bbox,
            Level level)
        {
            double zLevelFt = level?.Elevation ?? (bbox?.Min.Z ?? 0);

            bool hasAlu = window.LookupParameter(AL_SIDE) != null ||
                         window.LookupParameter(AL_TOP_BOT) != null ||
                         window.LookupParameter(AL_OFF_SIDE) != null;

            double sillFt = GetSillHeight(window);

            double glassBottomRelFt;

            if (hasAlu)
            {
                double offStijlFt = GetDoubleParam(window, AL_OFF_SIDE) ?? 0;
                double extraUnderFt = GetDoubleParam(window, AL_EXTRA_UNDER) ?? 0;
                double viewSillFt = GetDoubleParam(window, AL_VIEW_SILL) ?? GetDoubleParam(window, AL_TOP_BOT) ?? 0;

                glassBottomRelFt = sillFt - offStijlFt + extraUnderFt + viewSillFt;
            }
            else
            {
                double tBotFt = GetDoubleParam(window, P_BOT) ?? 0;
                glassBottomRelFt = sillFt + tBotFt + (17.0 * MM_TO_FT);
            }

            double glassBottomAbsFt = zLevelFt + glassBottomRelFt;
            double minRefAbsFt = zLevelFt + FT_600;

            double zRefFt = Math.Max(minRefAbsFt, glassBottomAbsFt);
            double zRefMm = zRefFt * FOOT_TO_MM;

            return (zRefFt, zRefMm);
        }

        public static double GetSillHeight(FamilyInstance window)
        {
            var bipNames = new[]
            {
                BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM,
                BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
                BuiltInParameter.INSTANCE_ELEVATION_PARAM
            };

            foreach (var bip in bipNames)
            {
                try
                {
                    Parameter p = window.get_Parameter(bip);
                    if (p != null && p.StorageType == StorageType.Double)
                    {
                        return p.AsDouble();
                    }
                }
                catch { }
            }

            return 0;
        }
        
        public static void SetDoubleParam(Element elem, string paramName, double value)
        {
            if (elem == null) return;
            Parameter p = elem.LookupParameter(paramName);
            if (p != null && !p.IsReadOnly)
            {
                p.Set(value);
            }
        }

        public static double? GetDoubleParam(Element elem, string paramName)
        {
            if (string.IsNullOrEmpty(paramName) || elem == null)
                return null;

            Parameter p = elem.LookupParameter(paramName);
            if (p == null)
            {
                ElementType type = elem.Document.GetElement(elem.GetTypeId()) as ElementType;
                if (type != null)
                {
                    p = type.LookupParameter(paramName);
                }
            }

            if (p != null)
            {
                try
                {
                    if (p.StorageType == StorageType.Double)
                        return p.AsDouble();
                    if (p.StorageType == StorageType.Integer)
                        return (double)p.AsInteger();
                    if (p.StorageType == StorageType.String)
                    {
                        string str = p.AsString()?.Replace(",", ".");
                        if (double.TryParse(str, out double val))
                            return val * MM_TO_FT;
                    }
                }
                catch { }
            }

            return null;
        }

        public static int GetIntParam(Element elem, string paramName, int defaultValue = 0)
        {
            Parameter p = elem.LookupParameter(paramName);
            if (p == null)
            {
                ElementType type = elem.Document.GetElement(elem.GetTypeId()) as ElementType;
                if (type != null)
                {
                    p = type.LookupParameter(paramName);
                }
            }

            if (p != null)
            {
                try
                {
                    if (p.StorageType == StorageType.Integer)
                        return p.AsInteger();
                    if (p.StorageType == StorageType.Double)
                        return (int)Math.Round(p.AsDouble());
                }
                catch { }
            }

            return defaultValue;
        }

        public static XYZ RotateXY(XYZ direction, double angleRad)
        {
            double c = Math.Cos(angleRad);
            double s = Math.Sin(angleRad);
            return new XYZ(
                direction.X * c - direction.Y * s,
                direction.X * s + direction.Y * c,
                0
            );
        }

        public static double? RayBox2DDistance(XYZ origin, XYZ direction, BoundingBoxXYZ bbox, double eps = 1e-9)
        {
            double ox = origin.X, oy = origin.Y;
            double dx = direction.X, dy = direction.Y;

            double xmin = bbox.Min.X, ymin = bbox.Min.Y;
            double xmax = bbox.Max.X, ymax = bbox.Max.Y;

            double tmin = -1e30;
            double tmax = 1e30;

            if (Math.Abs(dx) < eps)
            {
                if (ox < xmin || ox > xmax)
                    return null;
            }
            else
            {
                double tx1 = (xmin - ox) / dx;
                double tx2 = (xmax - ox) / dx;
                tmin = Math.Max(tmin, Math.Min(tx1, tx2));
                tmax = Math.Min(tmax, Math.Max(tx1, tx2));
            }

            if (Math.Abs(dy) < eps)
            {
                if (oy < ymin || oy > ymax)
                    return null;
            }
            else
            {
                double ty1 = (ymin - oy) / dy;
                double ty2 = (ymax - oy) / dy;
                tmin = Math.Max(tmin, Math.Min(ty1, ty2));
                tmax = Math.Min(tmax, Math.Max(ty1, ty2));
            }

            if (tmax < Math.Max(tmin, 0))
                return null;

            double tHit = tmin >= 0 ? tmin : tmax;
            if (tHit < 0)
                return null;

            return tHit;
        }

        public static double CalculateAlphaFromX(double zRef, double zObst, double xDist)
        {
            if (xDist <= 0)
                return 20.0;

            double dz = zObst - zRef;
            if (dz <= 0)
                return 20.0;

            double alpha = Math.Atan(dz / xDist) * 180.0 / Math.PI;
            return Math.Max(20.0, alpha);
        }

        public static void SetAlphaParameter(FamilyInstance window, double alphaDeg)
        {
            try
            {
                Parameter p = window.LookupParameter(PARAM_ALPHA);
                if (p != null && !p.IsReadOnly)
                {
                    p.Set(alphaDeg * Math.PI / 180.0);
                }
            }
            catch { }
        }

        public static Wall FindHostWallForWindow(BoundingBoxXYZ winBbox, List<Wall> walls)
        {
            Wall bestWall = null;
            double bestOverlap = 0;

            foreach (var wall in walls)
            {
                BoundingBoxXYZ wb = wall.get_BoundingBox(null);
                if (wb == null) continue;

                if (!BBoxesIntersect(winBbox, wb))
                    continue;

                double xOverlap = Math.Max(0, Math.Min(winBbox.Max.X, wb.Max.X) - Math.Max(winBbox.Min.X, wb.Min.X));
                double yOverlap = Math.Max(0, Math.Min(winBbox.Max.Y, wb.Max.Y) - Math.Max(winBbox.Min.Y, wb.Min.Y));
                double area = xOverlap * yOverlap;

                if (area > bestOverlap)
                {
                    bestOverlap = area;
                    bestWall = wall;
                }
            }

            return bestWall;
        }

        public static bool BBoxesIntersect(BoundingBoxXYZ b1, BoundingBoxXYZ b2)
        {
            if (b1 == null || b2 == null)
                return false;

            double xOverlap = Math.Min(b1.Max.X, b2.Max.X) - Math.Max(b1.Min.X, b2.Min.X);
            double yOverlap = Math.Min(b1.Max.Y, b2.Max.Y) - Math.Max(b1.Min.Y, b2.Min.Y);

            return xOverlap > -0.01 && yOverlap > -0.01;
        }

        public static bool BBoxesIntersectWithTolerance(BoundingBoxXYZ b1, BoundingBoxXYZ b2, double toleranceFt)
        {
            if (b1 == null || b2 == null)
                return false;

            double xOverlap = Math.Min(b1.Max.X + toleranceFt, b2.Max.X + toleranceFt) - Math.Max(b1.Min.X - toleranceFt, b2.Min.X - toleranceFt);
            double yOverlap = Math.Min(b1.Max.Y + toleranceFt, b2.Max.Y + toleranceFt) - Math.Max(b1.Min.Y - toleranceFt, b2.Min.Y - toleranceFt);

            return xOverlap > 0 && yOverlap > 0;
        }

        public static XYZ GetStartOnInsideFace(List<Wall> walls, XYZ p0, double zMidGlass, XYZ nDir)
        {
            try
            {
                XYZ bestPt = p0;
                double maxDist = -double.MaxValue;
                bool foundAny = false;

                foreach (Wall wall in walls)
                {
                    var sideFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior);
                    foreach (Reference r in sideFaces)
                    {
                        Face f = wall.GetGeometryObjectFromReference(r) as Face;
                        if (f is PlanarFace pf)
                        {
                            // Use the plane of the face to project, handling openings/voids
                            XYZ normal = pf.FaceNormal;
                            XYZ origin = pf.Origin;

                            // Calculate distance from p0 to the plane along nDir
                            // Plane equation: n * (P - origin) = 0
                            // Ray: P = p0 + d * nDir
                            // n * (p0 + d * nDir - origin) = 0
                            // d * (n * nDir) = n * (origin - p0)
                            // d = (n * (origin - p0)) / (n * nDir)

                            double denom = normal.DotProduct(nDir);
                            if (Math.Abs(denom) > 1e-6)
                            {
                                double d = normal.DotProduct(origin - p0) / denom;
                                XYZ ptOnPlane = p0 + d * nDir;
                                
                                // We want the face that is furthest in the nDir direction (inwards)
                                if (d > maxDist)
                                {
                                    maxDist = d;
                                    bestPt = ptOnPlane;
                                    foundAny = true;
                                }
                            }
                        }
                        else if (f != null)
                        {
                            // Fallback for non-planar faces
                            IntersectionResult res = f.Project(p0);
                            if (res != null)
                            {
                                XYZ ptOnFace = res.XYZPoint;
                                XYZ vec = ptOnFace - p0;
                                double dist = vec.DotProduct(nDir);
                                if (dist > maxDist)
                                {
                                    maxDist = dist;
                                    bestPt = ptOnFace;
                                    foundAny = true;
                                }
                            }
                        }
                    }
                }
                return foundAny ? bestPt : p0;
            }
            catch { }
            return p0;
        }

        public static HashSet<int> GetHostLikeWallIds(BoundingBoxXYZ bbox, List<Wall> allWalls)
        {
            var ids = new HashSet<int>();
            foreach (var wall in allWalls)
            {
                BoundingBoxXYZ wb = wall.get_BoundingBox(null);
                if (wb != null && BBoxesIntersect(bbox, wb))
                {
                    ids.Add(wall.Id.IntegerValue);
                }
            }
            return ids;
        }
    }
}
