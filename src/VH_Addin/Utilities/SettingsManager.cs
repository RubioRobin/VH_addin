// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Statische klasse voor het opslaan en laden van instellingen (configuratie)
// ============================================================================

using System; // Standaard .NET functionaliteit
using System.IO; // Bestandsbeheer
using VH_Tools.Models; // Eigen model voor instellingen
using System.Xml.Serialization; // XML serialisatie
// We gebruiken eenvoudige XML-serialisatie om afhankelijkheden te minimaliseren.
// Dit werkt zonder externe packages en is geschikt voor eenvoudige settings.

namespace VH_Tools.Utilities // Hoofdnamespace voor alle hulpfuncties van VH Tools
{
    // Statische klasse voor het opslaan en laden van instellingen (configuratie)
    public static class SettingsManager
    {
        /// <summary>
        /// Geeft het volledige pad naar het settings-bestand terug.
        /// </summary>
        private static string GetSettingsPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData); // AppData-map
            string folder = Path.Combine(appData, "VH_Engineering", "RevitAddin"); // Submap voor settings
            Directory.CreateDirectory(folder); // Zorg dat map bestaat
            return Path.Combine(folder, "kozijnstaat_settings.xml"); // Volledig pad naar settings-bestand
        }

        /// <summary>
        /// Slaat de instellingen op naar een XML-bestand.
        /// </summary>
        public static void SaveSettings(KozijnstaatConfig config)
        {
            try
            {
                string path = GetSettingsPath();
                XmlSerializer xs = new XmlSerializer(typeof(KozijnstaatConfig));
                using (StreamWriter sw = new StreamWriter(path))
                {
                    xs.Serialize(sw, config); // Serialiseer naar XML
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Fout bij opslaan van instellingen: {ex.Message}");
            }
        }

        /// <summary>
        /// Laadt de instellingen uit het XML-bestand.
        /// </summary>
        public static KozijnstaatConfig LoadSettings()
        {
            try
            {
                string path = GetSettingsPath();
                if (File.Exists(path))
                {
                    XmlSerializer xs = new XmlSerializer(typeof(KozijnstaatConfig));
                    using (StreamReader sr = new StreamReader(path))
                    {
                        return (KozijnstaatConfig)xs.Deserialize(sr); // Deserialiseer naar object
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Fout bij laden van instellingen: {ex.Message}");
            }

            return new KozijnstaatConfig(); // Return defaults if load fails
        }
    }
}
