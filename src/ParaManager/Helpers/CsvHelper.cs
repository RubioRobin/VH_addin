using ParaManager.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;

namespace ParaManager.Helpers
{
    public static class CsvHelper
    {
        /// <summary>
        /// Exports parameters to CSV file
        /// </summary>
        public static bool ExportParametersToCsv(List<ParameterData> parameters, string filePath)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    // Write header
                    writer.WriteLine("Name,Type,Group,Instance/Type,Categories,Description,User Modifiable,Visible,Hide When No Value");

                    // Write data
                    foreach (var param in parameters)
                    {
                        string line = string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\"",
                            EscapeCsv(param.Name),
                            EscapeCsv(param.ParameterType),
                            EscapeCsv(param.Group.ToString()),
                            param.IsInstance ? "Instance" : "Type",
                            EscapeCsv(string.Join("; ", param.Categories)),
                            EscapeCsv(param.Description ?? ""),
                            param.UserModifiable ? "Yes" : "No",
                            param.Visible ? "Yes" : "No",
                            param.HideWhenNoValue ? "Yes" : "No");

                        writer.WriteLine(line);
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Imports parameters from CSV file
        /// </summary>
        public static List<ParameterData> ImportParametersFromCsv(string filePath)
        {
            List<ParameterData> parameters = new List<ParameterData>();

            try
            {
                using (StreamReader reader = new StreamReader(filePath, Encoding.UTF8))
                {
                    // Skip header
                    string headerLine = reader.ReadLine();
                    if (headerLine == null)
                        return parameters;

                    // Read data lines
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var fields = ParseCsvLine(line);
                        if (fields.Count < 9)
                            continue;

                        ParameterData param = new ParameterData
                        {
                            Name = fields[0],
                            Description = fields[5],
                            UserModifiable = fields[6].Equals("Yes", StringComparison.OrdinalIgnoreCase),
                            Visible = fields[7].Equals("Yes", StringComparison.OrdinalIgnoreCase),
                            HideWhenNoValue = fields[8].Equals("Yes", StringComparison.OrdinalIgnoreCase),
                            IsInstance = fields[3].Equals("Instance", StringComparison.OrdinalIgnoreCase)
                        };

                        // Parse categories
                        if (!string.IsNullOrWhiteSpace(fields[4]))
                        {
                            param.Categories = fields[4].Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(c => c.Trim())
                                .Where(c => !string.IsNullOrWhiteSpace(c))
                                .ToList();
                        }

                        // Parse group
#if REVIT2025
                        param.Group = GroupTypeId.IdentityData;
                        if (fields.Count > 2 && !string.IsNullOrEmpty(fields[2]))
                        {
                            var prop = typeof(GroupTypeId).GetProperty(fields[2]);
                            if (prop != null) param.Group = (ForgeTypeId)prop.GetValue(null);
                        }
#else
                        if (Enum.TryParse(fields[2], out BuiltInParameterGroup group))
                        {
                            param.Group = group;
                        }
#endif

                        parameters.Add(param);
                    }
                }
            }
            catch { }

            return parameters;
        }

        /// <summary>
        /// Creates a CSV template for parameter import
        /// </summary>
        public static bool CreateCsvTemplate(string filePath)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    // Write header
                    writer.WriteLine("Name,Type,Group,Instance/Type,Categories,Description,User Modifiable,Visible,Hide When No Value");

                    // Write example row
                    writer.WriteLine("\"Example_Parameter\",\"Text\",\"PG_IDENTITY_DATA\",\"Instance\",\"Walls; Doors\",\"Example description\",\"Yes\",\"Yes\",\"No\"");
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string EscapeCsv(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            // Escape quotes by doubling them
            return value.Replace("\"", "\"\"");
        }

        private static List<string> ParseCsvLine(string line)
        {
            List<string> fields = new List<string>();
            bool inQuotes = false;
            StringBuilder currentField = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        // Escaped quote
                        currentField.Append('"');
                        i++; // Skip next quote
                    }
                    else
                    {
                        // Toggle quote mode
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    // Field separator
                    fields.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }

            // Add last field
            fields.Add(currentField.ToString());

            return fields;
        }
    }
}
