using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace ParaManager
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            string logFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VH_Addins", "paramanager_startup.log");
            try
            {
                // Ensure log directory exists
                string logDir = Path.GetDirectoryName(logFile);
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                File.AppendAllText(logFile, $"{DateTime.Now}: Starting ParaManager OnStartup...{Environment.NewLine}");

                // Target the shared VH Tools tab
                string tabName = "VH Tools";
                try
                {
                    application.CreateRibbonTab(tabName);
                    File.AppendAllText(logFile, $"{DateTime.Now}: Created Tab '{tabName}'{Environment.NewLine}");
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    // Tab likely already exists (e.g. from VH_Addin)
                    File.AppendAllText(logFile, $"{DateTime.Now}: Tab '{tabName}' already exists.{Environment.NewLine}");
                }
                catch (Exception ex)
                {
                    File.AppendAllText(logFile, $"{DateTime.Now}: Error creating tab: {ex.Message}{Environment.NewLine}");
                }

                // Helper to get or create panel
                RibbonPanel GetOrCreatePanel(string tName, string pName)
                {
                    var panels = application.GetRibbonPanels(tName);
                    var existing = panels.FirstOrDefault(p => p.Name == pName);
                    if (existing != null) return existing;
                    return application.CreateRibbonPanel(tName, pName);
                }

                // Create/Get Panel
                RibbonPanel panel = GetOrCreatePanel(tabName, "Parameter Tools");
                File.AppendAllText(logFile, $"{DateTime.Now}: Got/Created Panel 'Parameter Tools'{Environment.NewLine}");

                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                // Add Family Transfer button
                PushButtonData transferButtonData = new PushButtonData(
                    "FamilyTransfer",
                    "Family\nTransfer",
                    assemblyPath,
                    "ParaManager.Commands.FamilyTransferCommand");
                
                transferButtonData.LargeImage = GetEmbeddedImage("ParaManager.Resources.Icons.model_32.png");
                transferButtonData.ToolTip = "Transfer parameters between families";
                transferButtonData.LongDescription = "Copy parameters from a source family to one or more destination families, either directly or via Excel.";

                try
                {
                    panel.AddItem(transferButtonData);
                    File.AppendAllText(logFile, $"{DateTime.Now}: Added FamilyTransfer button{Environment.NewLine}");
                }
                catch (Exception ex)
                {
                     File.AppendAllText(logFile, $"{DateTime.Now}: Error adding button: {ex.Message}{Environment.NewLine}");
                }

                File.AppendAllText(logFile, $"{DateTime.Now}: ParaManager OnStartup Complete.{Environment.NewLine}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                try
                {
                    File.AppendAllText(logFile, $"{DateTime.Now}: CRITICAL ERROR in ParaManager OnStartup: {ex.Message}\n{ex.StackTrace}{Environment.NewLine}");
                }
                catch { }
                return Result.Failed;
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
                
                if (stream != null)
                {
                    BitmapImage image = new BitmapImage();
                    image.BeginInit();
                    image.StreamSource = stream;
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.EndInit();
                    image.Freeze();
                    return image;
                }
            }
            catch { }

            return null;
        }
    }
}
