// ============================================================================
// SpotElevationsCommand.cs
//
// Doel: Start het Spot Elevations venster voor het aanpassen van spot elevation weergave.
// Gebruik: Wordt aangeroepen als Revit External Command. Opent een WPF-venster.
// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Deze command opent een UI waarmee de gebruiker spot elevations in de actieve view
// kan aanpassen, bijvoorbeeld het toepassen van een lijnstijl.
// De UI is te vinden in Views/SpotElevationsWindow.xaml.
// ============================================================================

using Autodesk.Revit.Attributes; // Revit attribuut voor transactiemodus
using Autodesk.Revit.DB; // Revit database objecten
using Autodesk.Revit.UI; // Revit UI API
using System; // Standaard .NET functionaliteit
using VH_Addin.Views; // Eigen vensters

namespace VH_Tools.Commands // Hoofdnamespace voor alle commands van VH Tools
{
    [Transaction(TransactionMode.Manual)] // Revit: deze command start een handmatige transactie
    public class SpotElevationsCommand : IExternalCommand // Implementatie van een externe Revit-command
    {
        /// <summary>
        /// Start het Spot Elevations venster en past instellingen toe op de actieve view.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument; // Huidig UI-document
                Document doc = uidoc.Document; // Huidig Revit-document
                View view = doc.ActiveView; // Actieve view

                SpotElevationsWindow window = new SpotElevationsWindow(doc); // Open venster
                if (window.ShowDialog() != true)
                    return Result.Cancelled; // Annuleer als gebruiker annuleert

                var linePatternId = window.SelectedLinePatternId; // Gekozen lijnstijl
                var selectedColor = window.SelectedColor; // Gekozen kleur
                if (linePatternId == null || selectedColor == null)
                    return Result.Cancelled; // Annuleer als geen lijnstijl of kleur gekozen

                using (Transaction trans = new Transaction(doc, "Spot Elevations Override"))
                {
                    trans.Start();

                    // Verzamel alle spot elevations in de actieve view
                    var spotElevations = new FilteredElementCollector(doc, view.Id)
                        .OfCategory(BuiltInCategory.OST_SpotElevations)
                        .WhereElementIsNotElementType()
                        .ToElements();

                    // Maak grafische instellingen aan voor de gekozen lijnstijl en kleur
                    OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                    if (linePatternId != null) ogs.SetProjectionLinePatternId(linePatternId);
                    if (selectedColor != null) ogs.SetProjectionLineColor(selectedColor);

                    // Pas toe op elk element in de lijst
                    foreach (var spot in spotElevations)
                    {
                        view.SetElementOverrides(spot.Id, ogs);
                    }

                    trans.Commit();
                }

                return Result.Succeeded; // Command geslaagd
            }
            catch (Exception ex)
            {
                message = ex.Message; // Toon foutmelding in Revit
                return Result.Failed;
            }
        }
    }
}
