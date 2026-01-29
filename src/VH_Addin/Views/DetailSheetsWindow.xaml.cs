// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Hoofdvenster voor het aanmaken van detailsheets
// Volledige herimplementatie gebaseerd op Dynamo script VH_DetailSheets_V1.0.dyn
// ============================================================================

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace VH_Addin.Views
{
    public partial class DetailSheetsWindow : Window, INotifyPropertyChanged
    {
        private readonly Document _doc;
        private bool _isBusy = false;

        // Static fields for persistence of location
        private static double _lastLeft = double.NaN;
        private static double _lastTop = double.NaN;

        // Models
        public class ViewItem
        {
            public string Name { get; set; }
            public View View { get; set; }
            public string Level { get; set; }
            public string Template { get; set; }
            public Dictionary<string, string> Parameters { get; set; }
            public bool IsChecked { get; set; } = false;
        }

        public class TitleBlockItem
        {
            public string Name { get; set; }
            public FamilySymbol Symbol { get; set; }
        }

        public class PlanRow : INotifyPropertyChanged
        {
            private bool _selected;
            private string _sheetnr;
            private string _sheetName;
            private string _tb;
            private bool _plaatsen;

            public bool Selected { get => _selected; set { _selected = value; OnPropertyChanged(nameof(Selected)); } }
            public string GeneratedNumber { get; set; }
            public string Sheetnr { get => _sheetnr; set { _sheetnr = value; OnPropertyChanged(nameof(Sheetnr)); } }
            public string SheetName { get => _sheetName; set { _sheetName = value; OnPropertyChanged(nameof(SheetName)); } }
            public string ViewName { get; set; }
            public string TitleBlock { get => _tb; set { _tb = value; OnPropertyChanged(nameof(TitleBlock)); } }
            public bool PlaceView { get => _plaatsen; set { _plaatsen = value; OnPropertyChanged(nameof(PlaceView)); } }
            
            public View View { get; set; }
            public string OriginalName { get; set; }
            public int Counter { get; set; }
            public bool UserNameSet { get; set; } = false;
            public ElementId SheetId { get; set; } = ElementId.InvalidElementId;
            public string InitName { get; set; }
            public string InitSheetnr { get; set; }
            public string InitTB { get; set; }

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // Collections
        public ObservableCollection<ViewItem> AllViews { get; set; } = new ObservableCollection<ViewItem>();
        public ObservableCollection<TitleBlockItem> TitleBlocks { get; set; } = new ObservableCollection<TitleBlockItem>();
        public ObservableCollection<PlanRow> PlanRowsH { get; set; } = new ObservableCollection<PlanRow>();
        public ObservableCollection<PlanRow> PlanRowsV { get; set; } = new ObservableCollection<PlanRow>();

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        public DetailSheetsWindow(Document doc)
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

            LoadTitleBlocks();
            LoadViews();
            SetupEventHandlers();

            this.Closing += (s, e) => {
                _lastLeft = this.Left;
                _lastTop = this.Top;
            };
            
            // Init default values
            tbDatum.Text = DateTime.Now.ToString("dd-MM-yyyy");
            tbNumberTemplate.Text = "{view}";
            tbNameTemplate.Text = "{view}";
            tbOffsetX.Text = "0";
            tbOffsetY.Text = "250";
        }

        private void SetupEventHandlers()
        {
            // Tab 1 Events
            btnSelectAll.Click += (s, e) => SetAllViewCheckboxes(true);
            btnSelectNone.Click += (s, e) => SetAllViewCheckboxes(false);
            btnAddHorizontal.Click += (s, e) => AddToCategory("H");
            btnAddVertical.Click += (s, e) => AddToCategory("V");
            tbSearch.TextChanged += (s, e) => RefreshViewList();
            cmbFilterField.SelectionChanged += (s, e) => RefreshFilterValues();
            cmbFilterValue.SelectionChanged += (s, e) => RefreshViewList();

            // Name templates events
            tbNumberTemplate.TextChanged += (s, e) => RecalcAllRows();
            tbNameTemplate.TextChanged += (s, e) => RecalcAllRows();
            tbPrefix.TextChanged += (s, e) => RecalcAllRows();
            tbFind.TextChanged += (s, e) => RecalcAllRows();
            tbReplace.TextChanged += (s, e) => RecalcAllRows();
            tbSuffix.TextChanged += (s, e) => RecalcAllRows();

            // Tab 2 Events
            btnResultsSelectAll.Click += (s, e) => SetResultsSelection(true);
            btnResultsSelectNone.Click += (s, e) => SetResultsSelection(false);
            btnDeleteSelected.Click += (s, e) => DeleteSelectedRows();
            btnApplyFindReplace.Click += (s, e) => ApplyResultFindReplace();
            btnResetNames.Click += (s, e) => ResetResultNames();

            gridHorizontal.ItemsSource = PlanRowsH;
            gridVertical.ItemsSource = PlanRowsV;
        }

        // Custom Tab Logic to match ViewDuplicator
        private void BtnTabSelectie_Click(object sender, RoutedEventArgs e)
        {
            SetTabActive(btnTabSelectie, tabContentSelectie);
            SetTabInactive(btnTabResultaten, tabContentResultaten);
        }

        private void BtnTabResultaten_Click(object sender, RoutedEventArgs e)
        {
            SetTabActive(btnTabResultaten, tabContentResultaten);
            SetTabInactive(btnTabSelectie, tabContentSelectie);
        }

        private void SetTabActive(Button btn, System.Windows.Controls.Grid content)
        {
            btn.Style = (Style)FindResource("PillTabButtonActive");
            content.Visibility = System.Windows.Visibility.Visible;
        }

        private void SetTabInactive(Button btn, System.Windows.Controls.Grid content)
        {
            btn.Style = (Style)FindResource("PillTabButton");
            content.Visibility = System.Windows.Visibility.Collapsed;
        }

        #region Loading Data

        private void LoadTitleBlocks()
        {
            var collector = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsElementType()
                .Cast<FamilySymbol>()
                .OrderBy(tb => (tb.FamilyName ?? "") + " : " + tb.Name);

            TitleBlocks.Clear();
            foreach (var tb in collector)
            {
                var label = string.IsNullOrEmpty(tb.FamilyName) ? tb.Name : $"{tb.FamilyName} : {tb.Name}";
                TitleBlocks.Add(new TitleBlockItem { Name = label, Symbol = tb });
            }
            cmbTitleBlock.ItemsSource = TitleBlocks;
            cmbTitleBlock.DisplayMemberPath = "Name";
            if (TitleBlocks.Count > 0) cmbTitleBlock.SelectedIndex = 0;
        }

        private void LoadViews()
        {
            var views = new FilteredElementCollector(_doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && !(v is ViewSheet) && v.ViewType != ViewType.ThreeD)
                .OrderBy(v => v.Name);

            AllViews.Clear();
            foreach (var v in views)
            {
                if (v.ViewType == ViewType.Schedule || v.ViewType == ViewType.Legend || v.ViewType == ViewType.AreaPlan) continue;
                if (v.Name.ToLower().Contains("system browser") || v.Name.ToLower().Contains("project view")) continue;

                var level = v.GenLevel?.Name ?? string.Empty;
                var template = v.ViewTemplateId != ElementId.InvalidElementId ? _doc.GetElement(v.ViewTemplateId)?.Name : string.Empty;
                
                var paramDict = new Dictionary<string, string>();
                foreach (Parameter p in v.Parameters)
                {
                    try
                    {
                        var val = p.AsValueString() ?? p.AsString() ?? string.Empty;
                        paramDict[p.Definition.Name] = val;
                    }
                    catch { }
                }

                AllViews.Add(new ViewItem { Name = v.Name, View = v, Level = level, Template = template, Parameters = paramDict });
            }

            // Populate Filter Fields
            var paramNames = new HashSet<string>();
            foreach (var vi in AllViews)
            {
                foreach (var k in vi.Parameters.Keys) paramNames.Add(k);
            }

            cmbFilterField.Items.Clear();
            cmbFilterField.Items.Add("(Geen)");
            cmbFilterField.Items.Add("Level");
            cmbFilterField.Items.Add("Template");
            foreach (var name in paramNames.OrderBy(n => n))
            {
                if (name != "Level" && name != "Template") cmbFilterField.Items.Add(name);
            }
            cmbFilterField.SelectedIndex = 0;
            RefreshViewList();
        }

        private void RefreshFilterValues()
        {
            string field = cmbFilterField.SelectedItem as string;
            if (string.IsNullOrEmpty(field) || field == "(Geen)")
            {
                cmbFilterValue.Items.Clear();
                cmbFilterValue.IsEnabled = false;
                RefreshViewList();
                return;
            }

            cmbFilterValue.IsEnabled = true;
            var values = new HashSet<string>();
            foreach (var vi in AllViews)
            {
                string val = string.Empty;
                if (field == "Level") val = vi.Level;
                else if (field == "Template") val = vi.Template;
                else if (vi.Parameters.ContainsKey(field)) val = vi.Parameters[field];

                if (!string.IsNullOrEmpty(val)) values.Add(val);
            }

            cmbFilterValue.Items.Clear();
            cmbFilterValue.Items.Add("(Alle)");
            foreach (var v in values.OrderBy(x => x)) cmbFilterValue.Items.Add(v);
            cmbFilterValue.SelectedIndex = 0;
            RefreshViewList();
        }

        private void RefreshViewList()
        {
            viewsPanel.Children.Clear();
            string search = tbSearch.Text?.ToLower().Trim() ?? string.Empty;
            string field = cmbFilterField.SelectedItem as string;
            string filterVal = cmbFilterValue.SelectedItem as string;

            foreach (var vi in AllViews)
            {
                // Filter search
                if (!string.IsNullOrEmpty(search) && !vi.Name.ToLower().Contains(search)) continue;

                // Filter field
                if (!string.IsNullOrEmpty(field) && field != "(Geen)" && !string.IsNullOrEmpty(filterVal) && filterVal != "(Alle)")
                {
                    string val = string.Empty;
                    if (field == "Level") val = vi.Level;
                    else if (field == "Template") val = vi.Template;
                    else if (vi.Parameters.ContainsKey(field)) val = vi.Parameters[field];

                    if (val != filterVal) continue;
                }

                var cb = new CheckBox { Content = vi.Name, Tag = vi, IsChecked = vi.IsChecked, Margin = new Thickness(0, 2, 0, 2), HorizontalAlignment = HorizontalAlignment.Left };
                cb.Click += (s, e) => vi.IsChecked = cb.IsChecked ?? false;
                viewsPanel.Children.Add(cb);
            }
        }

        #endregion

        #region Logic & Actions

        private void SetAllViewCheckboxes(bool value)
        {
            foreach (var child in viewsPanel.Children)
            {
                if (child is CheckBox cb && cb.Tag is ViewItem vi)
                {
                    cb.IsChecked = value;
                    vi.IsChecked = value;
                }
            }
        }

        private void AddToCategory(string cat)
        {
            var selected = AllViews.Where(v => v.IsChecked).ToList();
            if (!selected.Any()) return;

            var targetList = (cat == "H") ? PlanRowsH : PlanRowsV;
            var existingIds = new HashSet<int>(targetList.Select(r => r.View.Id.IntegerValue));
            var tbLabel = (cmbTitleBlock.SelectedItem as TitleBlockItem)?.Name ?? string.Empty;

            foreach (var vi in selected)
            {
                if (existingIds.Contains(vi.View.Id.IntegerValue)) continue;

                var row = new PlanRow
                {
                    Selected = false,
                    View = vi.View,
                    ViewName = vi.Name,
                    OriginalName = vi.Name,
                    TitleBlock = tbLabel,
                    PlaceView = chkPlaceViews.IsChecked ?? true,
                    Sheetnr = string.Empty,
                    SheetName = string.Empty
                };
                targetList.Add(row);
            }
            RecalcAllRows();
            // Switch to results tab (simulated click)
            BtnTabResultaten_Click(null, null);
        }

        private void RecalcAllRows()
        {
            RecalcCategory(PlanRowsH);
            RecalcCategory(PlanRowsV);
        }

        private void RecalcCategory(ObservableCollection<PlanRow> rows)
        {
            int counter = 1;
            string numTpl = tbNumberTemplate.Text;
            string nameTpl = tbNameTemplate.Text;
            string prefix = tbPrefix.Text ?? "";
            string find = tbFind.Text ?? "";
            string repl = tbReplace.Text ?? "";
            string suf = tbSuffix.Text ?? "";
            string datum = tbDatum.Text ?? "";

            foreach (var row in rows)
            {
                row.Counter = counter++;
                string baseName = RenderTemplate(nameTpl, row.View, row.OriginalName, row.Counter, datum);
                string baseNum = RenderTemplate(numTpl, row.View, row.OriginalName, row.Counter, datum);

                // Apply base rules
                string finalName = ApplyRules(baseName, prefix, suf, find, repl);
                string finalNum = ApplyRules(baseNum, prefix, suf, find, repl);

                row.GeneratedNumber = finalNum;
                if (string.IsNullOrEmpty(row.Sheetnr) || row.Sheetnr == row.InitSheetnr) row.Sheetnr = finalNum;
                if (!row.UserNameSet) row.SheetName = finalName;

                if (row.InitName == null) // First time
                {
                    row.InitName = finalName;
                    row.InitSheetnr = finalNum;
                    row.InitTB = row.TitleBlock;
                }
            }
        }

        private string ApplyRules(string val, string prefix, string suffix, string find, string repl)
        {
            if (string.IsNullOrEmpty(val)) return string.Empty;
            string res = val;
            if (!string.IsNullOrEmpty(find)) res = res.Replace(find, repl);
            res = $"{prefix}{res}{suffix}";
            return res;
        }

        private string RenderTemplate(string tpl, View v, string origName, int counter, string date)
        {
            if (string.IsNullOrEmpty(tpl)) return origName;
            
            string res = tpl;
            res = Regex.Replace(res, @"\{view\}", origName ?? "", RegexOptions.IgnoreCase);
            res = Regex.Replace(res, @"\{n\}", counter.ToString(), RegexOptions.IgnoreCase);
            res = Regex.Replace(res, @"\{n:(\d+)\}", m => counter.ToString().PadLeft(int.Parse(m.Groups[1].Value), '0'), RegexOptions.IgnoreCase);
            res = Regex.Replace(res, @"\{date\}", date ?? "", RegexOptions.IgnoreCase);
            res = Regex.Replace(res, @"\{level\}", v.GenLevel?.Name ?? "", RegexOptions.IgnoreCase);
            res = Regex.Replace(res, @"\{project\}", _doc.ProjectInformation.Name ?? _doc.Title ?? "", RegexOptions.IgnoreCase);
            
            // Param support {param:Name}
            res = Regex.Replace(res, @"\{param:([^}]+)\}", m => {
                var p = v.LookupParameter(m.Groups[1].Value);
                if (p != null) return p.AsValueString() ?? p.AsString() ?? "";
                return "";
            }, RegexOptions.IgnoreCase);

            return res.Trim();
        }

        private void SetResultsSelection(bool val)
        {
            foreach (var r in PlanRowsH) r.Selected = val;
            foreach (var r in PlanRowsV) r.Selected = val;
        }

        private void DeleteSelectedRows()
        {
            var toDelH = PlanRowsH.Where(r => r.Selected).ToList();
            var toDelV = PlanRowsV.Where(r => r.Selected).ToList();
            foreach (var r in toDelH) PlanRowsH.Remove(r);
            foreach (var r in toDelV) PlanRowsV.Remove(r);
            RecalcAllRows();
        }

        private void ApplyResultFindReplace()
        {
            string find = fxTxtFind.Text;
            string repl = fxTxtReplace.Text ?? "";
            if (string.IsNullOrEmpty(find)) return;

            foreach (var r in PlanRowsH.Concat(PlanRowsV).Where(x => x.Selected))
            {
                r.SheetName = r.SheetName.Replace(find, repl);
                r.UserNameSet = true;
            }
        }

        private void ResetResultNames()
        {
            foreach (var r in PlanRowsH.Concat(PlanRowsV).Where(x => x.Selected))
            {
                r.SheetName = r.InitName;
                r.Sheetnr = r.InitSheetnr;
                r.TitleBlock = r.InitTB;
                r.UserNameSet = false;
            }
        }

        #endregion

        #region Execution

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (_isBusy) return;
            var allRows = PlanRowsH.Concat(PlanRowsV).ToList();
            if (!allRows.Any())
            {
                TaskDialog.Show("DetailSheets", "Geen details om te verwerken.");
                return;
            }

            _isBusy = true;
            try
            {
                using (var t = new Transaction(_doc, "Maak Detail Sheets"))
                {
                    t.Start();
                    
                    var usedNumbers = new HashSet<string>(new FilteredElementCollector(_doc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .Select(s => s.SheetNumber));

                    foreach (var row in allRows)
                    {
                        // Get TB symbol
                        var tbItem = TitleBlocks.FirstOrDefault(x => x.Name == row.TitleBlock) ?? TitleBlocks.FirstOrDefault();
                        if (tbItem == null) continue;

                        // Reserve unique number
                        string number = row.Sheetnr;
                        if (usedNumbers.Contains(number))
                        {
                            int i = 1;
                            while (usedNumbers.Contains($"{number} ({i})")) i++;
                            number = $"{number} ({i})";
                        }
                        usedNumbers.Add(number);

                        // Create Sheet
                        if (!tbItem.Symbol.IsActive) tbItem.Symbol.Activate();
                        var sheet = ViewSheet.Create(_doc, tbItem.Symbol.Id);
                        sheet.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(number);
                        sheet.get_Parameter(BuiltInParameter.SHEET_NAME).Set(row.SheetName);
                        
                        // Set Parameters
                        SetSheetParameters(sheet, row.View);

                        // Place View
                        if (row.PlaceView)
                        {
                            double ox = ParseDouble(tbOffsetX.Text) / 304.8;
                            double oy = ParseDouble(tbOffsetY.Text) / 304.8;
                            PlaceViewOnSheet(sheet, row.View, ox, oy);
                        }
                    }

                    t.Commit();
                }
                TaskDialog.Show("DetailSheets", "Klaar! Sheets zijn aangemaakt.");
                this.Close();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Er is een fout opgetreden: {ex.Message}");
            }
            finally { _isBusy = false; }
        }

        private void SetSheetParameters(ViewSheet sheet, View sourceView)
        {
            // Issue Date
            sheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE)?.Set(tbDatum.Text ?? "");

            // Sheet parameters
            SetParam(sheet, "VH_tekening engineer", tbTekenaar.Text);
            SetParam(sheet, "VH_tekening fase", tbTekFase.Text);
            SetParam(sheet, "VH_tekening status", tbTekStatus.Text);
            
            // Project/Shared parameters (some might need to be copied from view or from UI)
            SetParam(sheet, "VH_bouwdeel", tbBouwdeel.Text);
            SetParam(sheet, "VH_categorie", tbCategorie.Text);
            SetParam(sheet, "VH_fase", tbFase.Text);
            SetParam(sheet, "VH_onderdeel", tbOnderdeel.Text);

            // Copy from view if UI is empty
            CopyIfEmpty(sheet, sourceView, "VH_bouwdeel");
            CopyIfEmpty(sheet, sourceView, "VH_categorie");
            CopyIfEmpty(sheet, sourceView, "VH_fase");
            CopyIfEmpty(sheet, sourceView, "VH_onderdeel");
        }

        private void CopyIfEmpty(ViewSheet s, View v, string paramName)
        {
            var sp = s.LookupParameter(paramName);
            if (sp == null || sp.IsReadOnly || !string.IsNullOrEmpty(sp.AsString())) return;
            var vp = v.LookupParameter(paramName);
            if (vp != null) sp.Set(vp.AsValueString() ?? vp.AsString() ?? "");
        }

        private void SetParam(Element e, string name, string val)
        {
            if (string.IsNullOrEmpty(val)) return;
            var p = e.LookupParameter(name);
            if (p != null && !p.IsReadOnly) p.Set(val);
        }

        private void PlaceViewOnSheet(ViewSheet sheet, View v, double dx, double dy)
        {
            if (!Viewport.CanAddViewToSheet(_doc, sheet.Id, v.Id)) return;
            var vp = Viewport.Create(_doc, sheet.Id, v.Id, XYZ.Zero);
            _doc.Regenerate();

            // Center on titleblock
            XYZ center = XYZ.Zero;
            var tbs = new FilteredElementCollector(_doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements();
            if (tbs.Any())
            {
                var bb = tbs.First().get_BoundingBox(sheet);
                if (bb != null) center = (bb.Max + bb.Min) / 2.0;
            }
            else
            {
                var outline = sheet.Outline;
                if (outline != null) center = new XYZ((outline.Max.U + outline.Min.U) / 2.0, (outline.Max.V + outline.Min.V) / 2.0, 0);
            }

            XYZ target = new XYZ(center.X + dx, center.Y + dy, 0);
            try { vp.SetBoxCenter(target); }
            catch { ElementTransformUtils.MoveElement(_doc, vp.Id, target - vp.GetBoxCenter()); }
        }

        private double ParseDouble(string s)
        {
            if (double.TryParse(s?.Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double res)) return res;
            return 0;
        }

        #endregion

        #region Window Chrome
        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => this.Close();
        #endregion
    }
}
