using System;
using System.Collections.Generic;

namespace VH_Addin.Models
{
    public class ExportSettings
    {
        public string LastExportFolder { get; set; }
        public bool OpenFolderAfterExport { get; set; }
        public bool AutoSeparator { get; set; }
        public string Separator { get; set; } = "_";
        public List<string> NamingRule { get; set; } = new List<string>(); // Stored as string representations, e.g. "[Sheet Number]", "_"
        public bool ExportPdf { get; set; } = true;
        public bool ExportDwg { get; set; } = false;
        public bool CombinePdf { get; set; } = false;
        
        // Appearance
        public int ColorMode { get; set; } = 0; // 0=Color, 1=BlackLine, 2=Grayscale
        public int RasterQuality { get; set; } = 3; // 3=Presentation
        
        // DWG
        public string LastDwgSetupName { get; set; }

        public string ExtraParamName { get; set; }
    }
}
