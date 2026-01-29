// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Configuratieklasse voor de instellingen van de kozijnstaat-functionaliteit
// ============================================================================

using System;
using System.Collections.Generic;

namespace VH_Tools.Models
{
    public enum ElementTypeFilter
    {
        Kozijnen,
        Deuren,
        Beide
    }

    public class KozijnstaatConfig
    {
        public string Format { get; set; } = "A1";
        public double CustomWidth_MM { get; set; } = 0;
        public double CustomHeight_MM { get; set; } = 0;
        public string Orientation { get; set; } = "Liggend";
        public string ViewType { get; set; } = "Back";
        public string AssemblyCodeFilter { get; set; } = "31.";
        public double MarginBLR_MM { get; set; } = 75.0;
        public double MarginBottom_MM { get; set; } = 75.0;
        public double RowPitch_MM { get; set; } = 7200.0;
        public double Gap_MM { get; set; } = 2250.0;
        public double LineOffset_MM { get; set; } = 250.0;
        public double OffsetValue_MM { get; set; } = 0.0;
        public double MinBackGap_MM { get; set; } = 500.0;
        public bool UseHeights { get; set; } = true;
        public bool UseAllTypes { get; set; } = true;
        public int MaxTypes { get; set; } = 999;
        public bool OffsetDims { get; set; } = true;
        public bool CreateReferencePlanes { get; set; } = true;
        public double RefPlaneLeft_MM { get; set; } = -1000.0;
        public double RefPlaneTop_MM { get; set; } = 4000.0;
        public ElementTypeFilter ElementFilter { get; set; } = ElementTypeFilter.Kozijnen;

        public static Dictionary<string, (int, int)> PaperSizes = new Dictionary<string, (int, int)>
        {
            { "A0", (841, 1189) },
            { "A1", (594, 841) },
            { "A2", (420, 594) },
            { "A3", (297, 420) },
            { "A4", (210, 297) }
        };
    }

    public class PlacedItem
    {
        public Autodesk.Revit.DB.Element TopElement { get; set; }
        public Autodesk.Revit.DB.Element BotElement { get; set; }
        public Autodesk.Revit.DB.FamilySymbol Symbol { get; set; }
        public List<double> Offsets { get; set; } = new List<double>();
        public double EffTopOffset { get; set; }
        public double TrueMaxOffset { get; set; }
    }
}