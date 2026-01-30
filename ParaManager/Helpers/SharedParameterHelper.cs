using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ParaManager.Helpers
{
    public static class SharedParameterHelper
    {
        /// <summary>
        /// Merges multiple shared parameter files into one
        /// </summary>
        public static bool MergeSharedParameterFiles(List<string> sourceFiles, string targetFile, Autodesk.Revit.ApplicationServices.Application app)
        {
            try
            {
                // Create or open target file
                if (!File.Exists(targetFile))
                {
                    File.WriteAllText(targetFile, "# This is a Revit shared parameter file.\n# Do not edit manually.\n*META\tVERSION\tMINVERSION\n" +
                        "META\t2\t1\n*GROUP\tID\tNAME\n*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\n");
                }

                // Set as current shared parameter file
                string originalFile = app.SharedParametersFilename;
                app.SharedParametersFilename = targetFile;
                DefinitionFile targetDefFile = app.OpenSharedParameterFile();

                if (targetDefFile == null)
                    return false;

                // Track existing parameters to avoid duplicates
                HashSet<string> existingParams = new HashSet<string>();
                foreach (DefinitionGroup group in targetDefFile.Groups)
                {
                    foreach (Definition def in group.Definitions)
                    {
                        existingParams.Add($"{group.Name}|{def.Name}");
                    }
                }

                // Merge each source file
                foreach (string sourceFile in sourceFiles)
                {
                    if (!File.Exists(sourceFile))
                        continue;

                    app.SharedParametersFilename = sourceFile;
                    DefinitionFile sourceDefFile = app.OpenSharedParameterFile();

                    if (sourceDefFile == null)
                        continue;

                    // Copy groups and parameters
                    foreach (DefinitionGroup sourceGroup in sourceDefFile.Groups)
                    {
                        // Get or create group in target
                        app.SharedParametersFilename = targetFile;
                        targetDefFile = app.OpenSharedParameterFile();
                        
                        DefinitionGroup targetGroup = targetDefFile.Groups.get_Item(sourceGroup.Name);
                        if (targetGroup == null)
                        {
                            targetGroup = targetDefFile.Groups.Create(sourceGroup.Name);
                        }

                        // Copy parameters
                        app.SharedParametersFilename = sourceFile;
                        sourceDefFile = app.OpenSharedParameterFile();
                        DefinitionGroup currentSourceGroup = sourceDefFile.Groups.get_Item(sourceGroup.Name);

                        foreach (Definition sourceDef in currentSourceGroup.Definitions)
                        {
                            string paramKey = $"{sourceGroup.Name}|{sourceDef.Name}";
                            
                            if (!existingParams.Contains(paramKey))
                            {
                                // Switch to target file and create parameter
                                app.SharedParametersFilename = targetFile;
                                targetDefFile = app.OpenSharedParameterFile();
                                targetGroup = targetDefFile.Groups.get_Item(sourceGroup.Name);

                                if (sourceDef is ExternalDefinition extDef)
                                {
                                    ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(
                                        sourceDef.Name,
                                        sourceDef.GetDataType());
                                    
                                    options.Visible = extDef.Visible;
                                    options.UserModifiable = extDef.UserModifiable;
                                    
                                    if (!string.IsNullOrEmpty(extDef.Description))
                                        options.Description = extDef.Description;

                                    targetGroup.Definitions.Create(options);
                                    existingParams.Add(paramKey);
                                }
                            }
                        }
                    }
                }

                // Restore original file
                app.SharedParametersFilename = originalFile;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Exports shared parameters to a readable text format
        /// </summary>
        public static bool ExportSharedParametersToText(string sharedParamFile, string outputFile, Autodesk.Revit.ApplicationServices.Application app)
        {
            try
            {
                string originalFile = app.SharedParametersFilename;
                app.SharedParametersFilename = sharedParamFile;
                DefinitionFile defFile = app.OpenSharedParameterFile();

                if (defFile == null)
                    return false;

                using (StreamWriter writer = new StreamWriter(outputFile))
                {
                    writer.WriteLine("Shared Parameters Export");
                    writer.WriteLine("========================");
                    writer.WriteLine($"Source File: {sharedParamFile}");
                    writer.WriteLine($"Export Date: {DateTime.Now}");
                    writer.WriteLine();

                    foreach (DefinitionGroup group in defFile.Groups)
                    {
                        writer.WriteLine($"Group: {group.Name}");
                        writer.WriteLine(new string('-', 50));

                        foreach (Definition def in group.Definitions)
                        {
                            if (def is ExternalDefinition extDef)
                            {
                                writer.WriteLine($"  Name: {def.Name}");
                                writer.WriteLine($"  GUID: {extDef.GUID}");
                                writer.WriteLine($"  Type: {def.GetDataType().TypeId}");
                                writer.WriteLine($"  Visible: {extDef.Visible}");
                                writer.WriteLine($"  User Modifiable: {extDef.UserModifiable}");
                                
                                if (!string.IsNullOrEmpty(extDef.Description))
                                    writer.WriteLine($"  Description: {extDef.Description}");
                                
                                writer.WriteLine();
                            }
                        }

                        writer.WriteLine();
                    }
                }

                app.SharedParametersFilename = originalFile;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Validates a shared parameter file
        /// </summary>
        public static bool ValidateSharedParameterFile(string filePath, Autodesk.Revit.ApplicationServices.Application app, out string errorMessage)
        {
            errorMessage = string.Empty;

            try
            {
                if (!File.Exists(filePath))
                {
                    errorMessage = "File does not exist.";
                    return false;
                }

                string originalFile = app.SharedParametersFilename;
                app.SharedParametersFilename = filePath;
                DefinitionFile defFile = app.OpenSharedParameterFile();

                if (defFile == null)
                {
                    errorMessage = "Failed to open file. It may be corrupted or in an invalid format.";
                    app.SharedParametersFilename = originalFile;
                    return false;
                }

                // Check for duplicate parameter names within groups
                foreach (DefinitionGroup group in defFile.Groups)
                {
                    HashSet<string> paramNames = new HashSet<string>();
                    foreach (Definition def in group.Definitions)
                    {
                        if (!paramNames.Add(def.Name))
                        {
                            errorMessage = $"Duplicate parameter name '{def.Name}' found in group '{group.Name}'.";
                            app.SharedParametersFilename = originalFile;
                            return false;
                        }
                    }
                }

                app.SharedParametersFilename = originalFile;
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }
        public static ExternalDefinition GetExternalDefinitionByGuid(Autodesk.Revit.ApplicationServices.Application app, string guid)
        {
            try
            {
                DefinitionFile defFile = app.OpenSharedParameterFile();
                if (defFile == null) return null;

                foreach (DefinitionGroup group in defFile.Groups)
                {
                    foreach (Definition def in group.Definitions)
                    {
                        if (def is ExternalDefinition extDef && extDef.GUID.ToString().Equals(guid, StringComparison.OrdinalIgnoreCase))
                        {
                            return extDef;
                        }
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
    }
}
