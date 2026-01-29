// ============================================================================
// DetailSheetsCommand.cs
//
// Doel: Start het DetailSheets venster voor het aanmaken van detailsheets.
// Gebruik: Wordt aangeroepen als Revit External Command. Opent een WPF-venster.
// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Deze command opent een moderne UI waarmee de gebruiker snel detailsheets kan
// genereren op basis van geselecteerde views en instellingen.
// De UI is te vinden in Views/DetailSheetsWindow.xaml.
// ============================================================================

using Autodesk.Revit.UI; // Revit UI API
using System; // Standaard .NET functionaliteit
using System.Windows.Interop; // Voor koppeling met Revit-hoofdvenster
using VH_Addin.Views; // Eigen vensters

namespace VH_Tools // Hoofdnamespace voor alle functionaliteit van VH Tools
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)] // Revit: handmatige transactie
    public class DetailSheetsCommand : IExternalCommand // Implementatie van een externe Revit-command
    {
        /// <summary>
        /// Start het DetailSheets venster. Koppelt het venster aan het Revit-hoofdvenster.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            var window = new DetailSheetsWindow(commandData.Application.ActiveUIDocument.Document); // Maak venster aan
            var revitHandle = commandData.Application.MainWindowHandle; // Haal Revit-hoofdvenster op
            new WindowInteropHelper(window) { Owner = revitHandle }; // Koppel venster aan Revit
            window.ShowDialog(); // Toon venster
            return Result.Succeeded; // Command geslaagd
        }
    }
}
