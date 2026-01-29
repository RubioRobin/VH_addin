// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Statische helperklasse voor generieke annotaties in Revit
// ============================================================================

using Autodesk.Revit.DB; // Revit database objecten
using System; // Standaard .NET functionaliteit
using System.Collections.Generic; // Lijsten en collecties
using System.Linq; // LINQ voor collecties

namespace VH_Tools.Utilities // Hoofdnamespace voor alle hulpfuncties van VH Tools
{
    // Statische helperklasse voor generieke annotaties in Revit
    public static class AnnotationHelpers
    {
        /// <summary>
        /// Haalt alle Generic Annotation-symbolen op uit het document.
        /// </summary>
        public static List<FamilySymbol> GetGASymbols(Document doc)
        {
            var result = new List<FamilySymbol>(); // Resultaatlijst
            
            try
            {
                // Verzamel alle FamilySymbols van het type Generic Annotation
                var famTypes = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericAnnotation)
                    .WhereElementIsElementType()
                    .ToElements();

                foreach (var t in famTypes)
                {
                    if (t is FamilySymbol fs)
                        result.Add(fs); // Voeg toe aan resultaat
                }
            }
            catch { }

            try
            {
                // Verzamel alle AnnotationSymbolTypes (alternatieve route)
                var annTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(AnnotationSymbolType))
                    .ToElements();

                foreach (var t in annTypes)
                {
                    if (t is FamilySymbol fs && !result.Contains(fs))
                        result.Add(fs); // Voeg toe als nog niet aanwezig
                }
            }
            catch { }

            return result;
        }

        public static FamilySymbol PickDefaultGASymbol(Document doc)
        {
            FamilySymbol best = null;
            int scoreBest = -1;

            foreach (var s in GetGASymbols(doc))
            {
                string name = RevitHelpers.SafeLabel(s)?.ToLower() ?? "";
                int score = 0;
                if (name.Contains("kozijnstaat")) score += 10;
                if (name.Contains("tekst")) score += 5;
                if (name.Contains("ga")) score += 1;

                if (score > scoreBest)
                {
                    scoreBest = score;
                    best = s;
                }
            }

            return best;
        }

        public static FamilyInstance PlaceGATag(Document doc, FamilySymbol symType, View viewOwner, XYZ point)
        {
            if (symType == null) return null;

            if (symType is FamilySymbol fs)
            {
                if (!fs.IsActive)
                {
                    try
                    {
                        fs.Activate();
                        doc.Regenerate();
                    }
                    catch { }
                }

                try
                {
                    // Use NewFamilyInstance which is generally available; keep try/catch to handle API differences
                    return doc.Create.NewFamilyInstance(point, symType, viewOwner);
                }
                catch
                {
                    try
                    {
                        return doc.Create.NewFamilyInstance(point, symType, viewOwner);
                    }
                    catch { }
                }
            }

            return null;
        }
    }
}