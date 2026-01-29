using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using VH_Addin.Views;

namespace VH_Tools.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class SheetCreatorCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Launch modal window. 
                // All creation logic and transaction are now inside the window class
                // to allow the window to stay open after creation.
                SheetCreatorWindow window = new SheetCreatorWindow(doc);
                window.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
