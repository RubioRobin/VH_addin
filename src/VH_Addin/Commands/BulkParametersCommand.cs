// ============================================================================
// BulkParametersCommand.cs
//
// Doel: Start het Bulk Parameters venster voor het bewerken van parameters in bulk.
// Gebruik: Wordt aangeroepen als Revit External Command. Opent een WPF-venster.
// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Deze command is bedoeld voor het snel aanpassen van parameters van meerdere elementen
// in één keer. De UI is te vinden in Views/BulkParametersWindow.xaml.
// ============================================================================

using Autodesk.Revit.Attributes; // Revit attribuut voor transactiemodus
using Autodesk.Revit.DB; // Revit database objecten
using Autodesk.Revit.UI; // Revit UI API
using System; // Standaard .NET functionaliteit
using VH_Tools.Views; // Eigen vensters

namespace VH_Tools.Commands // Namespace matching project convention
{
    [Transaction(TransactionMode.Manual)] // Revit: deze command start een handmatige transactie
    [Regeneration(RegenerationOption.Manual)] // Revit: handmatige regeneratie
    public class BulkParametersCommand : IExternalCommand // Implementatie van een externe Revit-command
    {
        /// <summary>
        /// Start het Bulk Parameters venster. Foutafhandeling zorgt dat Revit niet crasht bij exceptions.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument; // Huidig UI-document
                Document doc = uidoc.Document; // Huidig Revit-document
                
                // Open het WPF-venster voor bulk parameterbewerking
                BulkParametersWindow window = new BulkParametersWindow(doc);
                window.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message; // Toon foutmelding in Revit
                return Result.Failed;
            }
        }
    }
}
