using OfficeOpenXml;
using ParaManager.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;

namespace ParaManager.Helpers
{
    public static class ExcelHelper
    {
        static ExcelHelper()
        {
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Exports parameters to Excel file
        /// </summary>
        public static bool ExportParametersToExcel(List<ParameterData> parameters, string filePath)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // Clear existing worksheets
                    while (package.Workbook.Worksheets.Count > 0)
                    {
                        package.Workbook.Worksheets.Delete(0);
                    }

                    // Create worksheet
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Parameters");

                    // Add headers
                    worksheet.Cells[1, 1].Value = "Name";
                    worksheet.Cells[1, 2].Value = "Type";
                    worksheet.Cells[1, 3].Value = "Group";
                    worksheet.Cells[1, 4].Value = "Instance/Type";
                    worksheet.Cells[1, 5].Value = "Shared";
                    worksheet.Cells[1, 6].Value = "GUID";
                    worksheet.Cells[1, 7].Value = "Categories";
                    worksheet.Cells[1, 8].Value = "Description";
                    worksheet.Cells[1, 9].Value = "User Modifiable";
                    worksheet.Cells[1, 10].Value = "Visible";
                    worksheet.Cells[1, 11].Value = "Hide When No Value";

                    // Style headers
                    using (var range = worksheet.Cells[1, 1, 1, 11])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5A059")); // VH Gold
                        range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }

                    // Add data
                    int row = 2;
                    foreach (var param in parameters)
                    {
                        worksheet.Cells[row, 1].Value = param.Name;
                        worksheet.Cells[row, 2].Value = param.ParameterType.ToString();
                        
#if REVIT2025
                        worksheet.Cells[row, 3].Value = param.GroupName ?? param.Group?.TypeId ?? "";
#else
                        worksheet.Cells[row, 3].Value = param.GroupName ?? param.Group.ToString();
#endif
                        worksheet.Cells[row, 4].Value = param.IsInstance ? "Instance" : "Type";
                        worksheet.Cells[row, 5].Value = param.IsShared ? "Yes" : "No";
                        worksheet.Cells[row, 6].Value = param.GUID ?? "";
                        worksheet.Cells[row, 7].Value = string.Join(", ", param.Categories ?? new List<string>());
                        worksheet.Cells[row, 8].Value = param.Description ?? "";
                        worksheet.Cells[row, 9].Value = param.UserModifiable ? "Yes" : "No";
                        worksheet.Cells[row, 10].Value = param.Visible ? "Yes" : "No";
                        worksheet.Cells[row, 11].Value = param.HideWhenNoValue ? "Yes" : "No";
                        row++;
                    }

                    // Apply Borders and Alignment to Data Range
                    int lastRow = row - 1;
                    if (lastRow > 1)
                    {
                        using (var range = worksheet.Cells[1, 1, lastRow, 11])
                        {
                            range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        }

                        // Center align specific columns (Instance/Type, Shared, Bools)
                        using (var range = worksheet.Cells[2, 4, lastRow, 5]) range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        using (var range = worksheet.Cells[2, 9, lastRow, 11]) range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }

                    // Auto-fit columns
                    worksheet.Cells.AutoFitColumns();
                    
                    // Save
                    package.Save();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Imports parameters from Excel file
        /// </summary>
        public static List<ParameterData> ImportParametersFromExcel(string filePath)
        {
            List<ParameterData> parameters = new List<ParameterData>();

            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                        return parameters;

                    int rowCount = worksheet.Dimension?.Rows ?? 0;
                    
                    // Start from row 2 (skip header)
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string name = worksheet.Cells[row, 1].Value?.ToString();
                        if (string.IsNullOrWhiteSpace(name))
                            continue;

                        ParameterData param = new ParameterData
                        {
                            Name = name,
                            Description = worksheet.Cells[row, 8].Value?.ToString() ?? "",
                            UserModifiable = worksheet.Cells[row, 9].Value?.ToString()?.ToLower() != "no",
                            Visible = worksheet.Cells[row, 10].Value?.ToString()?.ToLower() != "no",
                            HideWhenNoValue = worksheet.Cells[row, 11].Value?.ToString()?.ToLower() == "yes"
                        };

                        // Parse parameter type
                        string typeStr = worksheet.Cells[row, 2].Value?.ToString();
                        if (!string.IsNullOrEmpty(typeStr))
                        {
                            param.ParameterType = typeStr;
                        }

                        // Parse group
                        string groupStr = worksheet.Cells[row, 3].Value?.ToString();
#if REVIT2025
                        param.Group = GroupTypeId.IdentityData;
                        if (!string.IsNullOrEmpty(groupStr))
                        {
                            if (groupStr.Contains(":"))
                                param.Group = new ForgeTypeId(groupStr);
                            else
                            {
                                var prop = typeof(GroupTypeId).GetProperties()
                                    .FirstOrDefault(p => p.Name.Equals(groupStr, StringComparison.OrdinalIgnoreCase));
                                if (prop != null) param.Group = (ForgeTypeId)prop.GetValue(null);
                            }
                        }
#else
                        if (Enum.TryParse(groupStr, out BuiltInParameterGroup group))
                        {
                            param.Group = group;
                        }
#endif

                        // Parse instance/type
                        string instanceTypeStr = worksheet.Cells[row, 4].Value?.ToString();
                        param.IsInstance = instanceTypeStr?.ToLower() == "instance";

                        // Parse Shared
                        string sharedStr = worksheet.Cells[row, 5].Value?.ToString();
                        param.IsShared = sharedStr?.ToLower() == "yes";
                        
                        // Parse GUID
                        param.GUID = worksheet.Cells[row, 6].Value?.ToString();

                        // Parse categories
                        string categoriesStr = worksheet.Cells[row, 7].Value?.ToString();
                        if (!string.IsNullOrWhiteSpace(categoriesStr))
                        {
                            param.Categories = categoriesStr.Split(',')
                                .Select(c => c.Trim())
                                .Where(c => !string.IsNullOrWhiteSpace(c))
                                .ToList();
                        }

                        parameters.Add(param);
                    }
                }
            }
            catch { }

            return parameters;
        }

        /// <summary>
        /// Creates an Excel template for parameter import
        /// </summary>
        public static bool CreateTemplate(string filePath)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // Clear existing worksheets
                    while (package.Workbook.Worksheets.Count > 0)
                    {
                        package.Workbook.Worksheets.Delete(0);
                    }
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Parameters");

                    // Add headers
                    worksheet.Cells[1, 1].Value = "Name";
                    worksheet.Cells[1, 2].Value = "Type";
                    worksheet.Cells[1, 3].Value = "Group";
                    worksheet.Cells[1, 4].Value = "Instance/Type";
                    worksheet.Cells[1, 5].Value = "Shared";
                    worksheet.Cells[1, 6].Value = "GUID";
                    worksheet.Cells[1, 7].Value = "Categories";
                    worksheet.Cells[1, 8].Value = "Description";
                    worksheet.Cells[1, 9].Value = "User Modifiable";
                    worksheet.Cells[1, 10].Value = "Visible";
                    worksheet.Cells[1, 11].Value = "Hide When No Value";

                    // Style headers
                    using (var range = worksheet.Cells[1, 1, 1, 11])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5A059")); // VH Gold
                        range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }

                    // Add example row
                    worksheet.Cells[2, 1].Value = "Example_Parameter";
                    worksheet.Cells[2, 2].Value = "Text";
                    worksheet.Cells[2, 3].Value = "PG_IDENTITY_DATA";
                    worksheet.Cells[2, 4].Value = "Instance";
                    worksheet.Cells[2, 5].Value = "No";
                    worksheet.Cells[2, 6].Value = "";
                    worksheet.Cells[2, 7].Value = "Walls, Doors";
                    worksheet.Cells[2, 8].Value = "Example description";
                    worksheet.Cells[2, 9].Value = "Yes";
                    worksheet.Cells[2, 10].Value = "Yes";
                    worksheet.Cells[2, 11].Value = "No";

                    // Apply Borders and Alignment to Data Range
                    using (var range = worksheet.Cells[1, 1, 2, 11])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    }

                    // Center align specific columns
                    using (var range = worksheet.Cells[2, 4, 2, 5]) range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    using (var range = worksheet.Cells[2, 9, 2, 11]) range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    worksheet.Cells.AutoFitColumns();
                    package.Save();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
