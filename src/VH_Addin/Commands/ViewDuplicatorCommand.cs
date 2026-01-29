// ============================================================================
// ViewDuplicatorCommand.cs
//
// Doel: Start het View Duplicator venster voor het dupliceren van views.
// Gebruik: Wordt aangeroepen als Revit External Command. Opent een WPF-venster.
// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Deze command opent een UI waarmee de gebruiker views kan dupliceren met
// verschillende opties en instellingen. De UI is te vinden in Views/ViewDuplicatorWindow.xaml.
// ============================================================================

using Autodesk.Revit.Attributes; // Revit attribuut voor transactiemodus
using Autodesk.Revit.DB; // Revit database objecten
using Autodesk.Revit.UI; // Revit UI API
using VH_Addin.Views; // Eigen vensters

namespace VH_Tools // Hoofdnamespace voor alle functionaliteit van VH Tools
{
    [Transaction(TransactionMode.Manual)] // Revit: deze command start een handmatige transactie
    [Regeneration(RegenerationOption.Manual)] // Revit: handmatige regeneratie
    public class ViewDuplicatorCommand : IExternalCommand // Implementatie van een externe Revit-command
    {
        /// <summary>
        /// Start het View Duplicator venster. Opent een moderne UI voor view-duplicatie.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument; // Huidig UI-document
            Document doc = uidoc.Document; // Huidig Revit-document

            // Open het WPF-venster voor view duplicatie
            ViewDuplicatorWindow window = new ViewDuplicatorWindow(doc);
            window.ShowDialog();

            return Result.Succeeded; // Command geslaagd
        }
    }
}
