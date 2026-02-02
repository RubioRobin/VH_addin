using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using VH_Addin.ViewModels;
using VH_Addin.Views;

namespace VH_Addin.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class KozijnGeneratorCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Create ViewModel
            var vm = new KozijnGeneratorViewModel(commandData);
            
            // Create Window
            var window = new KozijnGeneratorWindow();
            window.DataContext = vm;
            
            // Handle closing
            vm.CloseAction = () => window.Close();
            
            // Show as Dialog (Modal) so we block Revit until done/cancelled
            // Or Show() for modeless, but then we need external event handler which is more complex.
            // Given the requirement to Load/Generate, Modal is safer for now unless user needs to interact with Revit while window is open.
            // Let's start with ShowDialog (Modal).
            
            // To parent correctly to Revit window:
            System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(window);
            helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

            window.ShowDialog();

            return Result.Succeeded;
        }
    }
}
