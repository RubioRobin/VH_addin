using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Input;
using VH_Addin.Models;

namespace VH_Addin.ViewModels
{
    public class KozijnGeneratorViewModel : INotifyPropertyChanged
    {
        private readonly ExternalCommandData _commandData;
        private string _selectedFolder;
        private string _selectedMaterial;
        private string _selectedDivision;
        private double _width = 1000;
        private double _height = 1500;
        private string _vlakverdeling;
        private string _vlakvullingen; // Legacy, kept for binding if needed or just internal
        private string _kozijnType;
        private string _kozijnMerk;

        // Pattern: NLRS_31_WIN_UN_kozijn {material} {division}_gen_VH.rfa
        private static readonly Regex FilePattern = new Regex(@"kozijn\s+(?<material>\w+)\s+(?<division>[\w-]+)_gen_VH\.rfa", RegexOptions.IgnoreCase);

        public event PropertyChangedEventHandler PropertyChanged;
        public Action CloseAction { get; set; }

        public ObservableCollection<string> AvailableMaterials { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> AvailableDivisions { get; set; } = new ObservableCollection<string>();
        
        // New: Dynamic Parameters List
        public ObservableCollection<ParameterInput> DynamicParameters { get; set; } = new ObservableCollection<ParameterInput>();

        // Cache of parsed files: Material -> List of Divisions
        private Dictionary<string, HashSet<string>> _filesCache = new Dictionary<string, HashSet<string>>();
        // Cache: Filename -> List of ParameterInput
        private Dictionary<string, List<ParameterInput>> _parameterListCache = new Dictionary<string, List<ParameterInput>>();

        // Visibility Flags
        private bool _isStep1Visible = true;
        public bool IsStep1Visible
        {
            get => _isStep1Visible;
            set { _isStep1Visible = value; OnPropertyChanged(); }
        }

        private bool _isStep2Visible = false;
        public bool IsStep2Visible
        {
            get => _isStep2Visible;
            set { _isStep2Visible = value; OnPropertyChanged(); }
        }

        public string SelectedFolder
        {
            get => _selectedFolder;
            set
            {
                _selectedFolder = value;
                OnPropertyChanged();
                LoadFiles();
                SaveSettings();
            }
        }

        public string SelectedMaterial
        {
            get => _selectedMaterial;
            set
            {
                _selectedMaterial = value;
                OnPropertyChanged();
                UpdateDivisions();
            }
        }

        public string SelectedDivision
        {
            get => _selectedDivision;
            set
            {
                _selectedDivision = value;
                OnPropertyChanged();
                if (!string.IsNullOrEmpty(value)) Vlakverdeling = value;
            }
        }

        public double Width
        {
            get => _width;
            set { _width = value; OnPropertyChanged(); }
        }

        public double Height
        {
            get => _height;
            set { _height = value; OnPropertyChanged(); }
        }

        public string Vlakverdeling
        {
            get => _vlakverdeling;
            set { _vlakverdeling = value; OnPropertyChanged(); }
        }

        public string Vlakvullingen
        {
            get => _vlakvullingen;
            set { _vlakvullingen = value; OnPropertyChanged(); }
        }

        public string KozijnType
        {
            get => _kozijnType;
            set { _kozijnType = value; OnPropertyChanged(); }
        }

        public string KozijnMerk
        {
            get => _kozijnMerk;
            set { _kozijnMerk = value; OnPropertyChanged(); }
        }

        public ICommand BrowseFolderCommand { get; }
        // public ICommand GenerateCommand { get; } // Legacy
        public ICommand NextStepCommand { get; }
        public ICommand PreviousStepCommand { get; }
        public ICommand FinalizeCommand { get; }

        public KozijnGeneratorViewModel(ExternalCommandData commandData)
        {
            _commandData = commandData;
            BrowseFolderCommand = new RelayCommand(BrowseFolder);
            
            NextStepCommand = new RelayCommand(NextStep);
            PreviousStepCommand = new RelayCommand(PreviousStep);
            FinalizeCommand = new RelayCommand(Finalize);

            var settings = KozijnSettings.Load();
            if (!string.IsNullOrEmpty(settings.LastUsedFolder) && Directory.Exists(settings.LastUsedFolder))
            {
                SelectedFolder = settings.LastUsedFolder;
            }
        }

        private void BrowseFolder(object obj)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SelectedFolder = dialog.SelectedPath;
                }
            }
        }

        private void NextStep(object obj)
        {
            if (string.IsNullOrEmpty(SelectedFolder) || string.IsNullOrEmpty(SelectedMaterial) || string.IsNullOrEmpty(SelectedDivision))
            {
                TaskDialog.Show("Warning", "Selecteer a.u.b. een materiaal en vlakverdeling.");
                return;
            }

            ReadParametersToCollection();

            IsStep1Visible = false;
            IsStep2Visible = true;
        }

        private void PreviousStep(object obj)
        {
            IsStep1Visible = true;
            IsStep2Visible = false;
        }

        private void LoadFiles()
        {
            _filesCache.Clear();
            AvailableMaterials.Clear();
            AvailableDivisions.Clear();

            if (Directory.Exists(SelectedFolder))
            {
                var files = Directory.GetFiles(SelectedFolder, "*.rfa");
                foreach (var file in files)
                {
                    var filename = Path.GetFileName(file);
                    var match = FilePattern.Match(filename);
                    if (match.Success)
                    {
                        string mat = match.Groups["material"].Value.ToLower(); // hout, alu, kunststof
                        string div = match.Groups["division"].Value;         // A1, B1, A1-A1, etc.

                        if (!_filesCache.ContainsKey(mat))
                            _filesCache[mat] = new HashSet<string>();
                        
                        _filesCache[mat].Add(div);
                    }
                }

                foreach (var mat in _filesCache.Keys.OrderBy(k => k))
                {
                    AvailableMaterials.Add(mat);
                }

                if (AvailableMaterials.Any())
                {
                    SelectedMaterial = AvailableMaterials.Contains("hout") ? "hout" : AvailableMaterials.First();
                }
            }
        }

        private void UpdateDivisions()
        {
            AvailableDivisions.Clear();
            if (!string.IsNullOrEmpty(SelectedMaterial) && _filesCache.ContainsKey(SelectedMaterial))
            {
                foreach (var div in _filesCache[SelectedMaterial].OrderBy(d => d))
                {
                    AvailableDivisions.Add(div);
                }

                if (AvailableDivisions.Any())
                    SelectedDivision = AvailableDivisions.First();
            }
        }

        private void SaveSettings()
        {
            var settings = new KozijnSettings { LastUsedFolder = SelectedFolder };
            KozijnSettings.Save(settings);
        }

        public void ReadParametersToCollection()
        {
            DynamicParameters.Clear();

            if (string.IsNullOrEmpty(SelectedFolder) || string.IsNullOrEmpty(SelectedMaterial) || string.IsNullOrEmpty(SelectedDivision))
                return;

            string targetName = $"NLRS_31_WIN_UN_kozijn {SelectedMaterial} {SelectedDivision}_gen_VH.rfa";
            string fullPath = Path.Combine(SelectedFolder, targetName);

            if (!File.Exists(fullPath))
            {
                 if (Directory.Exists(SelectedFolder))
                 {
                    var files = Directory.GetFiles(SelectedFolder, "*.rfa");
                    var match = files.FirstOrDefault(f => Path.GetFileName(f).Equals(targetName, StringComparison.OrdinalIgnoreCase));
                    if (match != null) fullPath = match;
                    else return;
                 }
                 else return;
            }

            // Check cache? For now reload to ensure fresh nested options
            // Optimization: Could cache the LIST of parameters but update options... 
            // Let's just re-read for robustness as it's fast enough.

            try
            {
                var app = _commandData.Application;
                Document familyDoc = app.Application.OpenDocumentFile(fullPath);
                
                try
                {
                    if (familyDoc.IsFamilyDocument)
                    {
                        var familyManager = familyDoc.FamilyManager;
                        
                        // Collect nested Window Families (Types) for Dropdown Options
                        var availableWindowTypes = new FilteredElementCollector(familyDoc)
                            .OfClass(typeof(FamilySymbol))
                            .OfCategory(BuiltInCategory.OST_Windows)
                            .Cast<FamilySymbol>()
                            // Exclude the family itself if somehow included, or just get all nested
                            .Select(s => s.Name)
                            .OrderBy(n => n)
                            .ToList();

                        var fillingParams = familyManager.Parameters
                            .Cast<FamilyParameter>()
                            .Where(p => p.Definition.Name.StartsWith("VH_vlakvulling_", StringComparison.OrdinalIgnoreCase))
                            .Where(p => p.Definition.GetGroupTypeId() == GroupTypeId.Graphics) // Filter for Graphics Group
                            .OrderBy(p => p.Definition.Name)
                            .ToList();

                        var newList = new List<ParameterInput>();

                        if (fillingParams.Any())
                        {
                            FamilyType currentType = familyManager.CurrentType;
                            foreach (var param in fillingParams)
                            {
                                string value = "";
                                if (currentType != null && currentType.HasValue(param))
                                {
                                    if (param.StorageType == StorageType.String) value = currentType.AsString(param);
                                    else if (param.StorageType == StorageType.Integer) value = currentType.AsInteger(param).ToString();
                                    else if (param.StorageType == StorageType.Double) value = currentType.AsDouble(param).ToString();
                                    else if (param.StorageType == StorageType.ElementId)
                                    {
                                        ElementId id = currentType.AsElementId(param);
                                        if (id != ElementId.InvalidElementId)
                                        {
                                           var elem = familyDoc.GetElement(id);
                                           if (elem != null) value = elem.Name;
                                        }
                                    }
                                }
                                
                                var input = new ParameterInput { Name = param.Definition.Name, Value = value };
                                foreach(var opt in availableWindowTypes) input.Options.Add(opt);
                                
                                newList.Add(input);
                            }
                        }

                        // Update list
                        foreach(var p in newList) DynamicParameters.Add(p);
                    }
                }
                finally
                {
                    familyDoc.Close(false);
                }
            }
            catch (Exception) { }
        }

        private void Finalize(object obj)
        {
             if (string.IsNullOrEmpty(SelectedFolder) || string.IsNullOrEmpty(SelectedMaterial) || string.IsNullOrEmpty(SelectedDivision))
            {
                TaskDialog.Show("Warning", "Selecteer a.u.b. een materiaal en vlakverdeling.");
                return;
            }

            string targetName = $"NLRS_31_WIN_UN_kozijn {SelectedMaterial} {SelectedDivision}_gen_VH.rfa";
            string originalPath = Path.Combine(SelectedFolder, targetName);

             if (!File.Exists(originalPath))
            {
                 if (Directory.Exists(SelectedFolder))
                 {
                    var files = Directory.GetFiles(SelectedFolder, "*.rfa");
                    var match = files.FirstOrDefault(f => Path.GetFileName(f).Equals(targetName, StringComparison.OrdinalIgnoreCase));
                    if (match != null) originalPath = match;
                    else 
                    {
                        TaskDialog.Show("Error", $"Kon bestand niet vinden:\n{targetName}");
                        return;
                    }
                 }
                 else return;
            }

            // Strategy: Copy to temp, modify there (create type + set params), then load.
            string tempPath = Path.Combine(Path.GetTempPath(), Path.GetFileName(originalPath));
            
            try
            {
                // 1. Create Temp Copy
                File.Copy(originalPath, tempPath, true);

                var app = _commandData.Application;
                var uidoc = app.ActiveUIDocument;
                var doc = uidoc.Document;

                // 2. Open Temp Document
                Document tempDoc = app.Application.OpenDocumentFile(tempPath);
                
                try
                {
                    using (Transaction t = new Transaction(tempDoc, "Create Window Type"))
                    {
                        t.Start();

                        // 3. Create New Type
                        string newTypeName = $"Merk-{KozijnMerk}_{Width}x{Height}";
                        FamilyManager mgr = tempDoc.FamilyManager;
                        
                        // Check if exists, if so switch to it, else new
                        FamilyType newType = null;
                        foreach(FamilyType ft in mgr.Types)
                        {
                            if (ft.Name.Equals(newTypeName, StringComparison.OrdinalIgnoreCase))
                            {
                                newType = ft;
                                mgr.CurrentType = ft;
                                break;
                            }
                        }
                        
                        if (newType == null)
                        {
                             newType = mgr.NewType(newTypeName);
                        }

                        // 4. Set Parameters (In Family Manager)
                        SetFamilyParam(mgr, tempDoc, "Breedte", Width);
                        SetFamilyParam(mgr, tempDoc, "Width", Width);
                        SetFamilyParam(mgr, tempDoc, "Hoogte", Height);
                        SetFamilyParam(mgr, tempDoc, "Height", Height);
                        SetFamilyParam(mgr, tempDoc, "Vlakverdeling", Vlakverdeling);
                        SetFamilyParam(mgr, tempDoc, "Kozijntype", KozijnType);
                        SetFamilyParam(mgr, tempDoc, "VH_kozijn_merk", KozijnMerk);

                        // Dynamic Params (Nested Types)
                        foreach(var dp in DynamicParameters)
                        {
                            SetFamilyParam(mgr, tempDoc, dp.Name, dp.Value);
                        }

                        t.Commit();
                    }

                    // 5. Load into Project
                    tempDoc.LoadFamily(doc, new FamilyLoadOptions());
                    
                    // Activate the type in the project to be helpful
                    using (Transaction t2 = new Transaction(doc, "Activate Type"))
                    {
                        t2.Start();
                        // Find the loaded family and type
                        string famName = Path.GetFileNameWithoutExtension(originalPath);
                        Family family = new FilteredElementCollector(doc)
                            .OfClass(typeof(Family))
                            .Cast<Family>()
                            .FirstOrDefault(f => f.Name.Equals(famName, StringComparison.OrdinalIgnoreCase));

                        if (family != null)
                        {
                            var stringNewTypeName = $"Merk-{KozijnMerk}_{Width}x{Height}";
                            var symbolId = family.GetFamilySymbolIds()
                                .FirstOrDefault(id => (doc.GetElement(id) as FamilySymbol)?.Name == stringNewTypeName);
                            
                            if (symbolId != null && symbolId != ElementId.InvalidElementId)
                            {
                                var symbol = doc.GetElement(symbolId) as FamilySymbol;
                                if (!symbol.IsActive) symbol.Activate();
                                TaskDialog.Show("Success", $"Family Loaded and Type '{stringNewTypeName}' created.");
                            }
                            else
                            {
                                 TaskDialog.Show("Success", "Family Loaded (Type might need manual activation).");
                            }
                        }
                        t2.Commit();
                    }
                }
                finally
                {
                    tempDoc.Close(false);
                }

                CloseAction?.Invoke();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
            finally
            {
                try 
                { 
                    if (File.Exists(tempPath)) File.Delete(tempPath); 
                } catch { }
            }
        }

        private void SetFamilyParam(FamilyManager mgr, Document doc, string name, object value)
        {
            FamilyParameter param = mgr.get_Parameter(name);
            if (param == null) return;

            if (param.StorageType == StorageType.Double && value is double d)
            {
                mgr.Set(param, UnitUtils.ConvertToInternalUnits(d, UnitTypeId.Millimeters));
            }
            else if (param.StorageType == StorageType.String && value is string s)
            {
                mgr.Set(param, s);
            }
            else if (param.StorageType == StorageType.ElementId && value is string typeName)
            {
                // Find nested symbol in FAMILY doc
                var symbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(x => x.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));
                
                if (symbol != null)
                {
                    mgr.Set(param, symbol.Id);
                }
            }
        }

        // Implements IFamilyLoadOptions
        public class FamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyMe, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInProject, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = true;
                return true;
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
