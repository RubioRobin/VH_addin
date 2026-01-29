using System;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace VH_DaglichtPlugin
{
    public class Application : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create ribbon tab
                string tabName = "VH Tools";
                try
                {
                    application.CreateRibbonTab(tabName);
                }
                catch
                {
                    // Tab already exists
                }

                // Create ribbon panel
                RibbonPanel panel = application.CreateRibbonPanel(tabName, "Daglicht");

                // Add Daglichttool button
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                
                PushButtonData buttonData = new PushButtonData(
                    "DaglichtTool",
                    "Daglicht\nBerekening",
                    assemblyPath,
                    "VH_DaglichtPlugin.DaglichtCommand"
                );

                buttonData.ToolTip = "VH Daglichttool - α/β berekening, glasoppervlakte en NEN 55% toets";
                buttonData.LongDescription = 
                    "Berekent α-waaier, β-hoek, glasoppervlakte en voert NEN 55% daglichttoets uit.\n\n" +
                    "Functies:\n" +
                    "• α-berekening met X/A-lijnen\n" +
                    "• β-berekening met echte geometrie\n" +
                    "• Glasoppervlakte (Hout/Alu logica)\n" +
                    "• NEN 55% toets per verblijfsgebied\n" +
                    "• CSV export";

                // Set icon (you can add a 32x32 PNG icon later)
                // buttonData.LargeImage = new BitmapImage(new Uri("pack://application:,,,/VH_DaglichtPlugin;component/Resources/Daglicht_32.png"));

                PushButton button = panel.AddItem(buttonData) as PushButton;
                button.AvailabilityClassName = "VH_DaglichtPlugin.CommandAvailability";

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to load VH Daglichttool:\n{ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }

    public class CommandAvailability : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication applicationData, Autodesk.Revit.DB.CategorySet selectedCategories)
        {
            // Command is available when a document is open
            return applicationData.ActiveUIDocument != null;
        }
    }
}
