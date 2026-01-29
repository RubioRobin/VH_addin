using System;
using System.Collections.Generic;

namespace VH_DaglichtPlugin
{
    public class WindowResult
    {
        public int ElementId { get; set; }
        public double? AlphaAvgDeg { get; set; }
        public double? BetaDeg { get; set; }
        public double? BetaRad { get; set; }
        public string BetaReason { get; set; }
        public int? BetaModelCurveId { get; set; }
        public double? GlassM2 { get; set; }
        public List<RayResult> Rays { get; set; }
        public string Message { get; set; }

        public WindowResult()
        {
            Rays = new List<RayResult>();
        }
    }

    public class RayResult
    {
        public int Index { get; set; }
        public double AngleOffsetDeg { get; set; }
        public double AlphaDeg { get; set; }
        public double? XDistMm { get; set; }
        public double? DHorizMm { get; set; }
        public double ZRefMm { get; set; }
        public double? ZObstMm { get; set; }
        public double LineLengthMm { get; set; }
        public int? ObstacleId { get; set; }
        public int? ModelCurveId { get; set; }
        public int? XLineModelCurveId { get; set; }
        public int? Alpha3DModelCurveId { get; set; }
        public int? X3DModelCurveId { get; set; }
        public int? SketchPlaneHId { get; set; }
    }

    public static class Constants
    {
        public const double FOOT_TO_MM = 304.8;
        public const double MM_TO_FT = 1.0 / 304.8;
        public const double SQFT_TO_SQM = 0.09290304;
        
        public const double LINE_LENGTH_FEET = 5000.0 * MM_TO_FT;  // 5m for α rays
        public const int RAY_COUNT = 11;
        public const double ANGLE_START = -50.0;
        public const double ANGLE_STEP = 10.0;
        
        public const double MAX_B_DISTANCE = 5000.0 * MM_TO_FT;  // 5m max for β
        public const double FT_600 = 600.0 * MM_TO_FT;
        
        // Parameter names - Daylight tool
        public const string P_BOT = "dikte_onderdorpel";
        public const string AL_TOP_BOT = "aanzicht_raamprofiel";
        public const string AL_SIDE = "aanzicht_stijl";
        public const string AL_OFF_SIDE = "offset_stijl";
        public const string AL_EXTRA_UNDER = "extra_stelruimte_onder";
        public const string AL_VIEW_SILL = "aanzicht_onderdorpel";
        
        public const string PARAM_ALPHA = "VH_kozijn_α";
        public const string PARAM_BETA_EPS = "VH_kozijn_β/ε";
        public const string PARAM_CAT = "VH_categorie";
        
        // Glass area parameters
        public static readonly string[] TARGET_PARAMS = { "VH_kozijn_Ad", "berekende_Ad" };
        public const string P_TYPENUMMER = "VH_typenummer";
        
        // Wood glass parameters
        public const string P_TOP = "dikte_bovendorpel";
        public const string P_SIDE_G = "dikte_eindstijlen";
        public const string P_BOT_G = "dikte_onderdorpel";
        public const string P_MULL_V_T = "dikte_tussenstijl";
        public const string P_MULL_H_T = "dikte_tussendorpel";
        public const string P_MULL_V_N = "aantal_tussenstijlen";
        public const string P_MULL_H_N = "aantal_tussendorpels";
        public const string P_SASH_WALL = "breedte_raamhout";
        
        // Aluminum glass parameters
        public const string AL_MULL_V = "aanzicht_tussenstijl";
        public const string AL_MULL_H = "aanzicht_tussendorpel";
        
        // Width/height parameter names
        public static readonly string[] W_NAMES = { "VH_kozijn_breedte", "Width", "Breedte", "Rough Width", "Nominal Width" };
        public static readonly string[] H_NAMES = { "VH_kozijn_hoogte", "Height", "Hoogte", "Rough Height", "Nominal Height" };
    }
}
