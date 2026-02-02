using System;
using System.IO;
using System.Xml.Serialization;

namespace VH_Addin.Models
{
    public class KozijnSettings
    {
        public string LastUsedFolder { get; set; }

        private static string SettingsPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "VH_Addin",
            "KozijnSettings.xml");

        public static void Save(KozijnSettings settings)
        {
            try
            {
                string dir = Path.GetDirectoryName(SettingsPath);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                XmlSerializer serializer = new XmlSerializer(typeof(KozijnSettings));
                using (StreamWriter writer = new StreamWriter(SettingsPath))
                {
                    serializer.Serialize(writer, settings);
                }
            }
            catch (Exception)
            {
                // Handle or ignore save errors
            }
        }

        public static KozijnSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(KozijnSettings));
                    using (StreamReader reader = new StreamReader(SettingsPath))
                    {
                        return (KozijnSettings)serializer.Deserialize(reader);
                    }
                }
            }
            catch (Exception)
            {
                // Handle or ignore load errors
            }

            return new KozijnSettings();
        }
    }
}
