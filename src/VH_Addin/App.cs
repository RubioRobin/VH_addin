using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using System.Linq;

namespace VH_Addin
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            string logFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VH_Addins", "startup.log");
            try
            {
                // Ensure log directory exists
                string logDir = Path.GetDirectoryName(logFile);
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                File.AppendAllText(logFile, $"{DateTime.Now}: Starting OnStartup...{Environment.NewLine}");

                // Create Ribbon Tab
                string tabName = "VH Tools";
                try
                {
                    application.CreateRibbonTab(tabName);
                    File.AppendAllText(logFile, $"{DateTime.Now}: Created Tab '{tabName}'{Environment.NewLine}");
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    // Tab likely already exists (e.g. from another addin)
                    File.AppendAllText(logFile, $"{DateTime.Now}: Tab '{tabName}' already exists.{Environment.NewLine}");
                }

                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                // Prepare Icon Safe
                // Use Embedded Resource for maximum reliability
                // Assembly Name in csproj is "VH_Tools", Namespace is "VH_Addin"
                // Default resource name: VH_Tools.Resources.VH_Icon32.png
                ImageSource vhIcon = GetEmbeddedImage("VH_Tools.Resources.VH_Icon32.png");

                // --- PANELS ---
                // Helper to get or create panel
                RibbonPanel GetOrCreatePanel(string tName, string pName)
                {
                    var panels = application.GetRibbonPanels(tName);
                    var existing = panels.FirstOrDefault(p => p.Name == pName);
                    if (existing != null) return existing;
                    return application.CreateRibbonPanel(tName, pName);
                }

                File.AppendAllText(logFile, $"{DateTime.Now}: Creating Panels...{Environment.NewLine}");

                // 1. Sheets
                RibbonPanel sheetsPanel = GetOrCreatePanel(tabName, "Sheets");
                AddPushButton(sheetsPanel, "VH_SheetCreator", "Sheet\nCreator", "VH_Tools.Commands.SheetCreatorCommand", assemblyPath, 
                    "Maak meerdere sheets aan met slimme nummering", "VH Tools: Meerdere sheets tegelijk aanmaken op basis van een startnummer en slimme logica.", vhIcon);
                
                AddPushButton(sheetsPanel, "VH_SheetRenamer", "Sheet\nRenamer", "VH_Tools.Commands.SheetRenamerCommand", assemblyPath, 
                    "Hernoem meerdere sheets tegelijk met zoekfunctie", "VH Tools: Meerdere sheets tegelijk hernoemen met een handige zoek- en filterfunctie.", vhIcon);

                AddPushButton(sheetsPanel, "VH_DetailSheets", "Detail\nSheets", "VH_Tools.DetailSheetsCommand", assemblyPath, 
                    "Maak detailsheets aan met geavanceerde opties", "VH Tools: Detailsheets aanmaken, naamgeving, parameters en meer.", vhIcon);

                AddPushButton(sheetsPanel, "VH_Export", "PDF/DWG\nExport", "VH_Tools.Commands.ExportCommand", assemblyPath,
                    "Exporteer sheets naar PDF en DWG", "VH Tools: Meerdere sheets tegelijk exporteren naar PDF en DWG met automatische formaat-detectie.", vhIcon);

                // 2. Views
                RibbonPanel viewsPanel = GetOrCreatePanel(tabName, "Views");
                AddPushButton(viewsPanel, "VH_ViewDuplicator", "View\nDuplicator", "VH_Tools.ViewDuplicatorCommand", assemblyPath,
                    "Dupliceer views met geavanceerde opties", "VH Tools: Views dupliceren met instellingen en opties.", vhIcon);

                AddPushButton(viewsPanel, "VH_Kozijnstaat", "Kozijnstaat\nGenereren", "VH_Tools.Commands.KozijnstaatCommand", assemblyPath,
                    "Genereer kozijnstaat op Legend View", "Genereert een kozijnstaat met alle geplaatste kozijnen (Assembly Code 31.*) op een Legend View.", vhIcon);

                // 3. Tools
                RibbonPanel toolsPanel = GetOrCreatePanel(tabName, "Tools");
                AddPushButton(toolsPanel, "VH_BulkParameters", "Bulk\nParameters", "VH_Tools.Commands.BulkParametersCommand", assemblyPath,
                    "Bulk edit parameters for multiple elements", "Select multiple elements and update parameter values in bulk.", vhIcon);

                AddPushButton(toolsPanel, "VH_Daglichtberekening", "Daglicht\nBerekening", "VH_Tools.Commands.DaglichtCommand", assemblyPath,
                    "Bereken α, β en glasoppervlakte volgens NEN", "Voer de daglichtberekening uit voor de geselecteerde kozijnen.", vhIcon);

                AddPushButton(toolsPanel, "VH_SpotElevations", "Spot Elev.\nOverride", "VH_Tools.Commands.SpotElevationsCommand", assemblyPath,
                    "Override graphic settings for spot elevations", "Apply graphic overrides (line pattern, weight, transparency, etc.) to all spot elevations in the current view.", vhIcon);

                AddPushButton(toolsPanel, "VH_KozijnGenerator", "Kozijn\nGenerator", "VH_Addin.Commands.KozijnGeneratorCommand", assemblyPath,
                    "Genereer en laad kozijnfamilies", "Selecteer een familie uit een map, configureer afmetingen en parameters, en laad een nieuw type in het project.", vhIcon);

                File.AppendAllText(logFile, $"{DateTime.Now}: OnStartup Complete.{Environment.NewLine}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                try
                {
                    File.AppendAllText(logFile, $"{DateTime.Now}: CRITICAL ERROR in OnStartup: {ex.Message}\n{ex.StackTrace}{Environment.NewLine}");
                    TaskDialog.Show("VH Tools Error", $"Failed to load VH Tools. See log at {logFile}.\n\nError: {ex.Message}");
                }
                catch { } // Last resort swallow
                return Result.Failed;
            }
        }

        private void AddPushButton(RibbonPanel panel, string name, string text, string className, string assembly, string tooltip, string longDesc, ImageSource icon)
        {
            try
            {
                PushButtonData data = new PushButtonData(name, text, assembly, className);
                PushButton button = panel.AddItem(data) as PushButton;
                button.ToolTip = tooltip;
                button.LongDescription = longDesc;
                
                if (icon != null)
                {
                    button.LargeImage = icon;
                    button.Image = icon;
                }
            }
            catch (Exception ex)
            {
                 string log = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VH_Addins", "startup.log");
                 File.AppendAllText(log, $"{DateTime.Now}: ERROR adding button '{name}': {ex.Message}\n{ex.StackTrace}{Environment.NewLine}");
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private BitmapImage GetEmbeddedImage(string resourceName)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                Stream stream = assembly.GetManifestResourceStream(resourceName);
                if (stream == null) return null;

                BitmapImage image = new BitmapImage();
                image.BeginInit();
                image.StreamSource = stream;
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.EndInit();
                image.Freeze();
                return image;
            }
            catch
            {
                return null;
            }
        }
    }
}
