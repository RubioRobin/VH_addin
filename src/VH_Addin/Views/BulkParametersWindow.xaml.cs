// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Hoofdvenster voor bulk parameterbewerking
// ============================================================================

using Autodesk.Revit.DB; // Revit database objecten
using Autodesk.Revit.UI; // Revit UI API
using System; // Standaard .NET functionaliteit
using System.Collections.Generic; // Lijsten en collecties
using System.Collections.ObjectModel; // Observable collections voor databinding
using System.ComponentModel; // Voor property change notificaties
using System.Linq; // LINQ voor collecties
using System.Windows; // WPF Window
using System.Windows.Controls; // WPF Controls
using System.Windows.Data; // WPF Data binding
using System.Windows.Media; // WPF Kleuren

// Aliassen om ambiguïteit te voorkomen
using WinGrid = System.Windows.Controls.Grid;
using WinControl = System.Windows.Controls.Control;
using WinComboBox = System.Windows.Controls.ComboBox;
using WinTextBox = System.Windows.Controls.TextBox;
using WinCheckBox = System.Windows.Controls.CheckBox;

namespace VH_Tools.Views // Hoofdnamespace voor alle vensters van VH Tools
{
    // Hoofdvenster voor bulk parameterbewerking
    public partial class BulkParametersWindow : Window, INotifyPropertyChanged
    {
        private readonly Document _doc; // Huidig Revit-document
        
        // Static fields for persistence of location
        private static double _lastLeft = double.NaN;
        private static double _lastTop = double.NaN;

        // Data Models voor de UI
        public class CategoryItem
        {
            public string Name { get; set; } // Naam van de categorie
            public BuiltInCategory Category { get; set; } // Revit BuiltInCategory
        }

        public class FamilyItem
        {
            public string Name { get; set; } // Naam van de family/type
            public Family Family { get; set; } // Revit Family object (voor loadable families)
            public bool IsSystemFamily { get; set; } // True als het een system family is
        }

        public class TypeItem
        {
            public string Name { get; set; }
            public ElementType ElementType { get; set; } // Kan FamilySymbol of system type zijn (WallType, FloorType, etc.)
        }

        public class ParamRow : INotifyPropertyChanged
        {
            public string Name { get; set; }
            public StorageType StorageType { get; set; }
            public string Kind { get; set; } // "yesno", "material", "default"
            
            private string _value;
            public string Value 
            { 
                get => _value; 
                set { _value = value; OnPropertyChanged(nameof(Value)); OnPropertyChanged(nameof(ValueFontStyle)); OnPropertyChanged(nameof(ValueForeground)); } 
            }
            
            public List<Parameter> Parameters { get; set; }
            public bool Common { get; set; }

            // UI Properties
            public System.Windows.Visibility IsYesNoVisible => Kind == "yesno" ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            public System.Windows.Visibility IsMaterialVisible => Kind == "material" ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            public System.Windows.Visibility IsDefaultVisible => (Kind != "yesno" && Kind != "material") ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            public Brush LabelForeground => Common ? (Brush)new BrushConverter().ConvertFromString("#2B2B2B") : Brushes.Gray;
            public double Opacity => Common ? 1.0 : 0.6;

            private bool? _yesNoValue;
            public bool? YesNoValue 
            { 
                get => _yesNoValue; 
                set 
                { 
                    _yesNoValue = value; 
                    OnPropertyChanged(nameof(YesNoValue)); 
                    OnPropertyChanged(nameof(YesNoLabel)); 
                    OnPropertyChanged(nameof(YesNoForeground)); 
                    OnPropertyChanged(nameof(IsIndeterminate));
                } 
            }

            public string YesNoLabel => IsIndeterminate ? "Varieert" : null;
            public Brush YesNoForeground => IsIndeterminate ? Brushes.Gray : Brushes.Black;
            public bool IsIndeterminate => YesNoValue == null;

            private Material _selectedMaterial;
            public Material SelectedMaterial 
            { 
                get => _selectedMaterial; 
                set { _selectedMaterial = value; OnPropertyChanged(nameof(SelectedMaterial)); } 
            }

            public FontStyle ValueFontStyle => Value == "<varies>" ? FontStyles.Italic : FontStyles.Normal;
            public Brush ValueForeground => Value == "<varies>" ? Brushes.Gray : Brushes.Black;

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // State
        private List<ParamRow> _currentParamRows = new List<ParamRow>();
        private List<Material> _allMaterials;
        public List<Material> Materials => _allMaterials;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        public BulkParametersWindow(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            DataContext = this;

            // Restore location
            if (!double.IsNaN(_lastLeft))
            {
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.Left = _lastLeft;
                this.Top = _lastTop;
            }

            // this.Title = "Bulk parameters"; // title is already set in XAML
            LoadCategories();
            SetupEvents();

            this.Closing += (s, e) => {
                _lastLeft = this.Left;
                _lastTop = this.Top;
            };
        }

        private long GetIdValue(ElementId id)
        {
            if (id == null) return -1;
#if REVIT2024 || REVIT2025
            return id.Value;
#else
            return (long)id.IntegerValue;
#endif
        }

        private void LoadCategories()
        {
            var catDict = new Dictionary<long, CategoryItem>();

            // Verzamel categorieën van loadable families (FamilySymbol)
            var symbolCollector = new FilteredElementCollector(_doc).OfClass(typeof(FamilySymbol));
            foreach (Element elem in symbolCollector)
            {
                var cat = elem.Category;
                if (cat != null)
                {
                    long catId = GetIdValue(cat.Id);
                    if (!catDict.ContainsKey(catId))
                    {
                        catDict[catId] = new CategoryItem 
                        { 
                            Name = cat.Name, 
                            Category = (BuiltInCategory)catId 
                        };
                    }
                }
            }

            // Verzamel categorieën van system families (ElementType maar niet FamilySymbol)
            var typeCollector = new FilteredElementCollector(_doc)
                .OfClass(typeof(ElementType))
                .Where(e => !(e is FamilySymbol)); // Exclude FamilySymbols (already collected)

            foreach (Element elem in typeCollector)
            {
                var cat = elem.Category;
                if (cat != null)
                {
                    long catId = GetIdValue(cat.Id);
                    if (!catDict.ContainsKey(catId))
                    {
                        catDict[catId] = new CategoryItem 
                        { 
                            Name = cat.Name, 
                            Category = (BuiltInCategory)catId 
                        };
                    }
                }
            }

            lstCategories.ItemsSource = catDict.Values.OrderBy(c => c.Name).ToList();

            _allMaterials = new FilteredElementCollector(_doc)
                .OfClass(typeof(Material))
                .Cast<Material>()
                .OrderBy(m => m.Name)
                .ToList();
            
            OnPropertyChanged(nameof(Materials));
        }

        private void SetupEvents()
        {
            lstCategories.SelectionChanged += LstCategories_SelectionChanged;
            lstFamilies.SelectionChanged += LstFamilies_SelectionChanged;
            lstTypes.SelectionChanged += LstTypes_SelectionChanged;
            txtParamSearch.TextChanged += TxtParamSearch_TextChanged;
        }

        private void LstCategories_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedCat = lstCategories.SelectedItem as CategoryItem;
            if (selectedCat == null)
            {
                lstFamilies.ItemsSource = null;
                lstTypes.ItemsSource = null;
                lstParams.ItemsSource = null;
                return;
            }

            var families = new List<FamilyItem>();

            // Verzamel loadable families (FamilySymbol)
            var symbolCollector = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(selectedCat.Category);

            var loadableFamilies = symbolCollector
                .Cast<FamilySymbol>()
                .Select(s => s.Family)
                .Distinct(new FamilyEqualityComparer(this))
                .Select(f => new FamilyItem { Name = f.Name, Family = f, IsSystemFamily = false });

            families.AddRange(loadableFamilies);

            // Verzamel system families (ElementType maar niet FamilySymbol)
            var typeCollector = new FilteredElementCollector(_doc)
                .OfClass(typeof(ElementType))
                .OfCategory(selectedCat.Category)
                .Where(e => !(e is FamilySymbol));

            // Groepeer system types op family name (bijv. "Basic Wall")
            var systemFamilies = typeCollector
                .Cast<ElementType>()
                .GroupBy(t => t.FamilyName)
                .Select(g => new FamilyItem { Name = g.Key, Family = null, IsSystemFamily = true });

            families.AddRange(systemFamilies);

            lstFamilies.ItemsSource = families.OrderBy(f => f.Name).ToList();
            lstTypes.ItemsSource = null;
            lstParams.ItemsSource = null;
        }

        private void LstFamilies_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedFamilies = lstFamilies.SelectedItems.Cast<FamilyItem>().ToList();
            if (!selectedFamilies.Any())
            {
                lstTypes.ItemsSource = null;
                lstParams.ItemsSource = null;
                return;
            }

            var selectedCat = lstCategories.SelectedItem as CategoryItem;
            var types = new List<TypeItem>();

            foreach (var famItem in selectedFamilies)
            {
                try
                {
                    if (famItem.IsSystemFamily)
                    {
                        // Voor system families: verzamel alle types met deze FamilyName
                        var systemTypes = new FilteredElementCollector(_doc)
                            .OfClass(typeof(ElementType))
                            .OfCategory(selectedCat.Category)
                            .Cast<ElementType>()
                            .Where(t => t.FamilyName == famItem.Name && !(t is FamilySymbol));

                        foreach (var type in systemTypes)
                        {
                            types.Add(new TypeItem { Name = type.Name, ElementType = type });
                        }
                    }
                    else
                    {
                        // Voor loadable families: gebruik bestaande logica
                        var family = famItem.Family;
                        var symbolIds = family.GetFamilySymbolIds();
                        foreach (var id in symbolIds)
                        {
                            var symbol = _doc.GetElement(id) as FamilySymbol;
                            if (symbol != null)
                            {
                                types.Add(new TypeItem { Name = symbol.Name, ElementType = symbol });
                            }
                        }
                    }
                }
                catch { }
            }

            lstTypes.ItemsSource = types.OrderBy(t => t.Name).ToList();
            lstParams.ItemsSource = null;
        }

        private void LstTypes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedTypes = lstTypes.SelectedItems.Cast<TypeItem>().ToList();
            if (!selectedTypes.Any())
            {
                lstParams.ItemsSource = null;
                return;
            }

            InitializeParams(selectedTypes.Select(t => t.ElementType).ToList());
        }

        private void InitializeParams(List<ElementType> elementTypes)
        {
            try
            {
                if (!elementTypes.Any())
                {
                    _currentParamRows.Clear();
                    lstParams.ItemsSource = null;
                    return;
                }
                
                _currentParamRows = GetEditableParams(elementTypes);
                FilterAndRenderParams();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fout bij laden parameters: {ex.Message}");
            }
        }

        private List<ParamRow> GetEditableParams(List<ElementType> elementTypes)
        {
            if (!elementTypes.Any()) return new List<ParamRow>();

            try
            {
                var paramMap = new Dictionary<string, List<Parameter>>();
                int count = elementTypes.Count;

                for (int i = 0; i < count; i++)
                {
                    var elemType = elementTypes[i];
                    foreach (Parameter p in elemType.Parameters)
                    {
                        if (p == null || p.IsReadOnly) continue;
                        var defName = p.Definition?.Name;
                        if (string.IsNullOrEmpty(defName)) continue;
                        if (!paramMap.ContainsKey(defName)) paramMap[defName] = new List<Parameter>(new Parameter[count]);
                        paramMap[defName][i] = p;
                    }
                }

                var rows = new List<ParamRow>();
                foreach (var kvp in paramMap)
                {
                    var validParams = kvp.Value.Where(p => p != null).ToList();
                    if (!validParams.Any()) continue;

                    var firstP = validParams.First();
                    var kind = DetectKind(firstP);
                    var values = validParams.Select(p => GetParamValueString(p)).ToList();
                    var distinctValues = values.Distinct().ToList();
                    string displayValue = (distinctValues.Count == 1) ? distinctValues[0] : "<varies>";

                    var row = new ParamRow
                    {
                        Name = kvp.Key,
                        StorageType = firstP.StorageType,
                        Kind = kind,
                        Value = displayValue,
                        Parameters = validParams,
                        Common = (validParams.Count == count)
                    };

                    // Initial state for editors
                    if (kind == "yesno")
                    {
                        row.YesNoValue = (displayValue == "<varies>") ? (bool?)null : (displayValue.ToLower() == "ja" || displayValue.ToLower() == "yes" || displayValue == "1");
                    }
                    else if (kind == "material")
                    {
                        if (displayValue != "<varies>")
                        {
                            row.SelectedMaterial = _allMaterials.FirstOrDefault(m => m.Name == displayValue);
                        }
                    }

                    rows.Add(row);
                }

                return rows.OrderBy(r => r.Name).ToList();
            }
            catch { return new List<ParamRow>(); }
        }

        private string GetParamValueString(Parameter p)
        {
            if (p.StorageType == StorageType.String) return p.AsString() ?? "";
            string val = p.AsValueString();
            if (!string.IsNullOrEmpty(val)) return val;
            if (p.StorageType == StorageType.Integer) return p.AsInteger().ToString();
            if (p.StorageType == StorageType.Double) return p.AsDouble().ToString("F2");
            if (p.StorageType == StorageType.ElementId) 
            {
#if REVIT2024 || REVIT2025
                return p.AsElementId().Value.ToString();
#else
                return p.AsElementId().IntegerValue.ToString();
#endif
            }
            return "";
        }

        private string DetectKind(Parameter p)
        {
            if (p.StorageType == StorageType.Integer)
            {
                try
                {
#if REVIT2023 || REVIT2024 || REVIT2025
                    if (p.Definition.GetDataType() == SpecTypeId.Boolean.YesNo) return "yesno";
#endif
                } catch { }

                string valStr = p.AsValueString()?.ToLower();
                if (valStr == "yes" || valStr == "no" || valStr == "ja" || valStr == "nee") return "yesno";
                string name = p.Definition.Name.ToLower();
                if (name.Contains("yes/no") || name.Contains("ja/nee")) return "yesno";
            }

            if (p.StorageType == StorageType.ElementId)
            {
                var eid = p.AsElementId();
                if (eid != ElementId.InvalidElementId)
                {
                    var el = _doc.GetElement(eid);
                    if (el is Material) return "material";
                }
            }
            return "default";
        }

        private void FilterAndRenderParams()
        {
            string search = txtParamSearch.Text?.ToLower() ?? "";
            var visibleRows = _currentParamRows
                .Where(r => string.IsNullOrEmpty(search) || r.Name.ToLower().Contains(search))
                .ToList();
            lstParams.ItemsSource = visibleRows;
        }

        private void TxtParamSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterAndRenderParams();
        }

        private void TitleBar_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left) this.DragMove();
        }

        private void BtnYesNo_Click(object sender, RoutedEventArgs e)
        {
            if (sender is WinCheckBox chk && chk.Tag is ParamRow row)
            {
                row.YesNoValue = chk.IsChecked;
            }
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is WinTextBox txt && txt.Text == "<varies>")
            {
                txt.Text = "";
                // Property change in Value will update font style/color via binding
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Weet je zeker dat je de parameters wilt aanpassen?", "Bevestigen", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;

            int changes = 0;
            using (Transaction t = new Transaction(_doc, "Bulk Parameter Edit"))
            {
                t.Start();
                foreach (var row in _currentParamRows)
                {
                    try
                    {
                        if (row.Kind == "yesno")
                        {
                            if (row.YesNoValue == null) continue;
                            int val = row.YesNoValue.Value ? 1 : 0;
                            foreach (var p in row.Parameters) p.Set(val);
                            changes++;
                        }
                        else if (row.Kind == "material")
                        {
                            if (row.SelectedMaterial == null) continue;
                            foreach (var p in row.Parameters) p.Set(row.SelectedMaterial.Id);
                            changes++;
                        }
                        else
                        {
                            if (row.Value == "<varies>" || string.IsNullOrEmpty(row.Value)) continue;
                            foreach (var p in row.Parameters)
                            {
                                if (p.StorageType == StorageType.String) p.Set(row.Value);
                                else p.SetValueString(row.Value);
                            }
                            changes++;
                        }
                    }
                    catch { }
                }
                t.Commit();
            }

            if (changes > 0)
            {
                MessageBox.Show($"{changes} parameters aangepast.");
                DialogResult = true;
                Close();
            }
            else MessageBox.Show("Geen wijzigingen.");
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => Close();

        private class FamilyEqualityComparer : IEqualityComparer<Family>
        {
            private readonly BulkParametersWindow _parent;
            public FamilyEqualityComparer(BulkParametersWindow parent) { _parent = parent; }
            public bool Equals(Family x, Family y) => _parent.GetIdValue(x.Id) == _parent.GetIdValue(y.Id);
            public int GetHashCode(Family obj) => _parent.GetIdValue(obj.Id).GetHashCode();
        }
    }
}
