// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Hoofdvenster voor het dupliceren van views
// ============================================================================

using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;

namespace VH_Addin.Views
{
    public class TabData
    {
        public Button Button { get; set; }
        public System.Windows.Controls.Grid ContentGrid { get; set; }
        public List<int> ViewIds { get; set; }
        public string TabName { get; set; }
    }

    public partial class ViewDuplicatorWindow : Window
    {
        private readonly Document _doc;
        private List<ViewItem> _allViews;
        private List<CheckBox> _checkBoxes = new List<CheckBox>();
        private Dictionary<int, bool> _selectedViews = new Dictionary<int, bool>();
        private int? _mainLastIndex = null;
        private SolidColorBrush _previewBrush = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#6B7280"));

        private List<Autodesk.Revit.DB.View> _viewTemplates;
        private List<Element> _scopeBoxes;

        private Button _activeTabButton;
        private System.Windows.Controls.Grid _mainContentGrid;
        private Dictionary<string, TabData> _savedTabs = new Dictionary<string, TabData>();

        // Static fields for persistence of location
        private static double _lastLeft = double.NaN;
        private static double _lastTop = double.NaN;

        public ViewDuplicatorWindow(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            
            // Restore location
            if (!double.IsNaN(_lastLeft))
            {
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.Left = _lastLeft;
                this.Top = _lastTop;
            }

            LoadMainContent();
            CreateMainTab();
            LoadData();
            SetupEventHandlers();
            RefreshList();

            this.Closing += (s, e) => {
                _lastLeft = this.Left;
                _lastTop = this.Top;
            };
        }

        private void LoadMainContent()
        {
            try
            {
                // Load the main content XAML into tabContent Grid
                var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var assemblyDir = Path.GetDirectoryName(assemblyLocation);
                var xamlPath = Path.Combine(assemblyDir, "ViewDuplicatorMainContent.xaml");
                
                if (!File.Exists(xamlPath))
                {
                    throw new FileNotFoundException($"ViewDuplicatorMainContent.xaml not found at: {xamlPath}");
                }
                
                using (var stream = new FileStream(xamlPath, FileMode.Open, FileAccess.Read))
                {
                    _mainContentGrid = (System.Windows.Controls.Grid)XamlReader.Load(stream);
                }
                tabContent.Children.Add(_mainContentGrid);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading ViewDuplicatorMainContent.xaml:\n{ex.Message}\n\nAssembly location: {System.Reflection.Assembly.GetExecutingAssembly().Location}", 
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        private T FindControl<T>(string name) where T : FrameworkElement
        {
            return _mainContentGrid.FindName(name) as T;
        }

        // Main controls (found via FindControl)
        private ComboBox cmbP => FindControl<ComboBox>("cmbP");
        private ComboBox cmbV => FindControl<ComboBox>("cmbV");
        private TextBox txtS => FindControl<TextBox>("txtS");
        private ComboBox cmbTpl => FindControl<ComboBox>("cmbTpl");
        private ComboBox cmbScp => FindControl<ComboBox>("cmbScp");
        private ComboBox cmbDup => FindControl<ComboBox>("cmbDup");
        private TextBox txtPre => FindControl<TextBox>("txtPre");
        private TextBox txtFind => FindControl<TextBox>("txtFind");
        private TextBox txtRepl => FindControl<TextBox>("txtRepl");
        private TextBox txtOld => FindControl<TextBox>("txtOld");
        private TextBox txtNew => FindControl<TextBox>("txtNew");
        private TextBox txtCount => FindControl<TextBox>("txtCount");
        private TextBox txtSet => FindControl<TextBox>("txtSet");
        private ListBox lst => FindControl<ListBox>("lst");
        private Button btnAll => FindControl<Button>("btnAll");
        private Button btnNone => FindControl<Button>("btnNone");
        private Button btnSaveSet => FindControl<Button>("btnSaveSet");

        private void CreateMainTab()
        {
            var mainTab = new Button
            {
                Content = "Hoofd",
                Style = (Style)FindResource("PillTabButtonActive")
            };
            mainTab.Click += (s, e) => SwitchToTab(mainTab, null, null);
            tabButtons.Children.Add(mainTab);
            _activeTabButton = mainTab;
            // remove button hidden on main tab
            btnRemoveTab.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void SwitchToTab(Button tabButton, System.Windows.Controls.Grid contentGrid, string tabName)
        {
            // Reset previous active tab
            if (_activeTabButton != null)
            {
                _activeTabButton.Style = (Style)FindResource("PillTabButton");
            }

            // Set new active tab
            _activeTabButton = tabButton;
            _activeTabButton.Style = (Style)FindResource("PillTabButtonActive");

            // Switch content
            tabContent.Children.Clear();
            if (contentGrid != null)
                tabContent.Children.Add(contentGrid);
            else
                tabContent.Children.Add(_mainContentGrid);

            // Always hide top remove button; delete is now at bottom for set tabs
            btnRemoveTab.Visibility = System.Windows.Visibility.Collapsed;
            btnRemoveTab.Click -= BtnRemoveTab_Click;
            
            // Update Execute button: change to Delete when on set tab
            if (string.IsNullOrEmpty(tabName))
            {
                // Main tab
                btnExecute.Content = "Uitvoeren";
                btnExecute.Tag = null;
            }
            else
            {
                // Set tab
                btnExecute.Content = "Verwijderen";
                btnExecute.Tag = tabName;
            }
        }

        private void BtnRemoveTab_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button b && b.Tag is string tabName)
            {
                RemoveSetTab(tabName);
            }
        }

        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            // Check if this is a set tab (btnExecute.Tag contains tab name)
            string tabName = btnExecute.Tag as string;
            if (!string.IsNullOrEmpty(tabName))
            {
                // Set tab: delete it
                RemoveSetTab(tabName);
            }
            else
            {
                // Main tab: execute duplication
                BtnRun_Click(sender, e);
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            SetAllChecks(true);
        }

        private void BtnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            SetAllChecks(false);
        }

        private void LoadData()
        {
            // Load all views
            _allViews = new FilteredElementCollector(_doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .Where(v => !IsExcluded(v))
                .Select(v => new ViewItem { Name = v.Name, View = v })
                .OrderBy(v => v.Name)
                .ToList();

            // Load string parameters
            var paramNames = GetStringParameters();
            cmbP.Items.Add("<Geen>");
            foreach (var p in paramNames)
                cmbP.Items.Add(p);
            cmbP.SelectedIndex = 0;

            cmbV.Items.Add("<Alle>");
            cmbV.SelectedIndex = 0;

            // Load view templates
            _viewTemplates = new FilteredElementCollector(_doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .Where(v => v.IsTemplate)
                .OrderBy(v => v.Name)
                .ToList();

            cmbTpl.Items.Add("<Geen>");
            foreach (var vt in _viewTemplates)
                cmbTpl.Items.Add(vt.Name);
            cmbTpl.SelectedIndex = 0;

            // Load scope boxes
            _scopeBoxes = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                .WhereElementIsNotElementType()
                .OrderBy(s => s.Name)
                .ToList();

            cmbScp.Items.Add("<Geen>");
            foreach (var sb in _scopeBoxes)
                cmbScp.Items.Add(sb.Name);
            cmbScp.SelectedIndex = 0;

            // Load duplicate options
            cmbDup.Items.Add("Duplicate");
            cmbDup.Items.Add("With Detailing");
            cmbDup.Items.Add("As Dependent");
            cmbDup.SelectedIndex = 0;

            txtCount.Text = "1";
        }

        private bool IsExcluded(Autodesk.Revit.DB.View v)
        {
            if (v.IsTemplate) return true;
            if (v is View3D || v is ViewDrafting || v is ViewSchedule || v is ViewSheet) return true;
            
            try
            {
                if (v.ViewType == ViewType.ThreeD || 
                    v.ViewType == ViewType.AreaPlan || 
                    v.ViewType == ViewType.Legend || 
                    v.ViewType == ViewType.DraftingView)
                    return true;
            }
            catch { }

            string name = (v.Name ?? "").ToLower();
            return name.Contains("system browser") || name.Contains("project view");
        }

        private List<string> GetStringParameters()
        {
            var paramSet = new HashSet<string> { "VH_categorie", "Comments", "View Subcategory", "Type Mark", "Discipline" };
            
            int count = 0;
            foreach (var item in _allViews.Take(200))
            {
                try
                {
                    foreach (Parameter p in item.View.Parameters)
                    {
                        if (p != null && p.Definition != null && p.StorageType == StorageType.String)
                            paramSet.Add(p.Definition.Name);
                    }
                }
                catch { }
                count++;
            }

            return paramSet.OrderBy(p => p).ToList();
        }

        private string GetParameterValue(Autodesk.Revit.DB.View view, string paramName)
        {
            try
            {
                Parameter p = view.LookupParameter(paramName);
                if (p != null)
                {
                    if (p.StorageType == StorageType.String)
                        return p.AsString();
                    else
                        return p.AsValueString();
                }
            }
            catch { }
            return "";
        }

        private void SetupEventHandlers()
        {
            txtS.TextChanged += (s, e) => RefreshList();
            cmbP.SelectionChanged += (s, e) => { RefreshParameterValues(); RefreshList(); };
            cmbV.SelectionChanged += (s, e) => RefreshList();

            txtPre.TextChanged += (s, e) => UpdatePreviews();
            txtFind.TextChanged += (s, e) => UpdatePreviews();
            txtRepl.TextChanged += (s, e) => UpdatePreviews();
            txtOld.TextChanged += (s, e) => UpdatePreviews();
            txtNew.TextChanged += (s, e) => UpdatePreviews();
            txtCount.TextChanged += (s, e) => UpdatePreviews();

            btnAll.Click += (s, e) => SetAllChecks(true);
            btnNone.Click += (s, e) => SetAllChecks(false);
            btnSaveSet.Click += BtnSaveSet_Click;
            // btnRun is now btnExecute and hooked up in XAML
        }

        private void RefreshParameterValues()
        {
            cmbV.Items.Clear();
            cmbV.Items.Add("<Alle>");

            string paramName = cmbP.SelectedItem?.ToString();
            if (paramName != null && paramName != "<Geen>")
            {
                var values = new HashSet<string>();
                foreach (var item in _allViews.Take(500))
                {
                    string val = GetParameterValue(item.View, paramName);
                    if (!string.IsNullOrEmpty(val))
                        values.Add(val);
                }

                foreach (var v in values.OrderBy(x => x))
                    cmbV.Items.Add(v);
            }
            cmbV.SelectedIndex = 0;
        }

        private void RefreshList()
        {
            lst.Items.Clear();
            _checkBoxes.Clear();

            string search = (txtS.Text ?? "").ToLower().Trim();
            string paramName = cmbP.SelectedItem?.ToString();
            string paramValue = cmbV.SelectedItem?.ToString();

            foreach (var item in _allViews)
            {
                // Filter by parameter
                if (paramName != null && paramName != "<Geen>")
                {
                    string val = GetParameterValue(item.View, paramName);
                    if (paramValue != null && paramValue != "<Alle>" && val != paramValue)
                        continue;
                }

                // Filter by search text
                if (!string.IsNullOrEmpty(search) && !item.Name.ToLower().Contains(search))
                    continue;

                AddViewRow(item);
            }
        }

        private void AddViewRow(ViewItem item)
        {
            System.Windows.Controls.Grid row = new System.Windows.Controls.Grid 
            { 
                Margin = new Thickness(0),
                Background = new SolidColorBrush(Colors.Transparent)
            };
            row.ColumnDefinitions.Add(new ColumnDefinition());

            CheckBox chk = new CheckBox
            {
                Margin = new Thickness(8, 8, 8, 8),
                VerticalAlignment = VerticalAlignment.Center,
                IsChecked = false
            };

            StackPanel cont = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center
            };

            TextBlock lbl = new TextBlock
            {
                Text = item.Name,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            TextBlock preview = new TextBlock
            {
                Margin = new Thickness(12, 0, 0, 0),
                Foreground = _previewBrush,
                IsHitTestVisible = false,
                FontStyle = FontStyles.Italic
            };

            cont.Children.Add(lbl);
            cont.Children.Add(preview);
            chk.Content = cont;
            row.Children.Add(chk);

            chk.Tag = new CheckBoxData { Name = item.Name, View = item.View, Preview = preview, Row = row };
            row.Tag = chk;

            chk.Click += OnCheckBoxClick;
            
            _checkBoxes.Add(chk);
            lst.Items.Add(row);
        }

        private void OnCheckBoxClick(object sender, RoutedEventArgs e)
        {
            CheckBox chk = sender as CheckBox;
            var data = chk.Tag as CheckBoxData;
            bool val = chk.IsChecked ?? false;
            int cur = lst.Items.IndexOf(data.Row);

            if (Keyboard.Modifiers == ModifierKeys.Shift && _mainLastIndex != null && cur >= 0)
            {
                int a = Math.Min(_mainLastIndex.Value, cur);
                int b = Math.Max(_mainLastIndex.Value, cur);
                
                for (int idx = a; idx <= b; idx++)
                {
                    System.Windows.Controls.Grid rowGrid = lst.Items[idx] as System.Windows.Controls.Grid;
                    CheckBox cbox = rowGrid.Tag as CheckBox;
                    cbox.IsChecked = val;
                    
                    var d = cbox.Tag as CheckBoxData;
                    _selectedViews[d.View.Id.IntegerValue] = val;
                    d.Preview.Text = val ? GetPreviewText(d.Name) : "";
                }
            }
            else
            {
                _selectedViews[data.View.Id.IntegerValue] = val;
                data.Preview.Text = val ? GetPreviewText(data.Name) : "";
            }

            if (cur >= 0)
                _mainLastIndex = cur;
        }

        private void UpdatePreviews()
        {
            foreach (var chk in _checkBoxes)
            {
                var data = chk.Tag as CheckBoxData;
                if (chk.IsChecked ?? false)
                    data.Preview.Text = GetPreviewText(data.Name);
            }
        }

        private string GetPreviewText(string name)
        {
            string baseName = RenameBase(name);
            int count = ParseCount(txtCount.Text);
            return "→ " + baseName + (count <= 1 ? "" : $" ×{count}");
        }

        private string RenameBase(string name)
        {
            string prefix = txtPre.Text ?? "";
            string find = txtFind.Text ?? "";
            string replace = txtRepl.Text ?? "";
            string suffix = txtNew.Text ?? "";

            string result = name ?? "";
            if (!string.IsNullOrEmpty(find))
                result = result.Replace(find, replace);

            return Clean(prefix + result + suffix) ?? name;
        }

        private string Clean(string name)
        {
            return Regex.Replace(name ?? "", @"[^A-Za-z0-9_\-\s\.\(\)\[\]]", "_").Trim();
        }

        private int ParseCount(string text, int defaultValue = 1)
        {
            if (int.TryParse((text ?? "").Trim(), out int n))
                return Math.Max(1, Math.Min(200, n));
            return defaultValue;
        }

        private void SetAllChecks(bool value)
        {
            foreach (var chk in _checkBoxes)
            {
                chk.IsChecked = value;
                var data = chk.Tag as CheckBoxData;
                _selectedViews[data.View.Id.IntegerValue] = value;
                data.Preview.Text = value ? GetPreviewText(data.Name) : "";
            }
        }

        private void BtnSaveSet_Click(object sender, RoutedEventArgs e)
        {
            var checkedBoxes = _checkBoxes.Where(c => c.IsChecked ?? false).ToList();
            if (!checkedBoxes.Any())
            {
                MessageBox.Show("Vink minimaal één view aan om een set op te slaan.", "Set opslaan");
                return;
            }

            string tabName = (txtSet.Text ?? "").Trim();
            if (string.IsNullOrEmpty(tabName))
                tabName = $"Set {DateTime.Now:yyyy-MM-dd HH:mm:ss}";

            // Create set tab with selected views
            var viewIds = checkedBoxes.Select(c => (c.Tag as CheckBoxData).View.Id.IntegerValue).ToList();
            
            var defaults = new Dictionary<string, object>
            {
                { "tpl_index", cmbTpl.SelectedIndex },
                { "scp_index", cmbScp.SelectedIndex },
                { "dup_index", cmbDup.SelectedIndex },
                { "pre", txtPre.Text ?? "" },
                { "find", txtFind.Text ?? "" },
                { "repl", txtRepl.Text ?? "" },
                { "old", txtOld.Text ?? "" },
                { "new", txtNew.Text ?? "" },
                { "count", ParseCount(txtCount.Text) }
            };

            CreateSetTab(tabName, viewIds, defaults);
        }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            // Check if we have sets or main tab selections
            List<DuplicationJob> jobs = new List<DuplicationJob>();

            // Collect jobs from set tabs
            if (_setRows.Any())
            {
                foreach (var kvp in _setRows)
                {
                    foreach (var row in kvp.Value.Where(r => r.Checkbox.IsChecked ?? false))
                    {
                        Autodesk.Revit.DB.View view = _doc.GetElement(new ElementId(row.ViewId)) as Autodesk.Revit.DB.View;
                        if (view != null)
                        {
                            jobs.Add(new DuplicationJob
                            {
                                SourceView = view,
                                NewName = row.NewNameTextBox.Text,
                                TemplateIndex = row.TemplateCombo.SelectedIndex,
                                ScopeBoxIndex = row.ScopeBoxCombo.SelectedIndex,
                                DuplicateIndex = row.DuplicateCombo.SelectedIndex,
                                Copies = row.CopiesCount
                            });
                        }
                    }
                }
            }

            // Collect jobs from main tab if no sets or sets are empty
            if (!jobs.Any())
            {
                var selectedCheckboxes = _checkBoxes.Where(c => c.IsChecked ?? false).ToList();
                if (!selectedCheckboxes.Any())
                {
                    MessageBox.Show("Selecteer minimaal één view om te dupliceren.", "Geen selectie");
                    return;
                }

                int tplIndex = cmbTpl.SelectedIndex;
                int scpIndex = cmbScp.SelectedIndex;
                int dupIndex = cmbDup.SelectedIndex;
                int copies = ParseCount(txtCount.Text);

                foreach (var chk in selectedCheckboxes)
                {
                    var data = chk.Tag as CheckBoxData;
                    string baseName = RenameBase(data.View.Name);

                    jobs.Add(new DuplicationJob
                    {
                        SourceView = data.View,
                        NewName = baseName,
                        TemplateIndex = tplIndex,
                        ScopeBoxIndex = scpIndex,
                        DuplicateIndex = dupIndex,
                        Copies = copies
                    });
                }
            }

            if (!jobs.Any())
            {
                MessageBox.Show("Geen views geselecteerd.", "Geen selectie");
                return;
            }

            // Execute duplication
            try
            {
                HashSet<string> existingNames = new HashSet<string>(
                    new FilteredElementCollector(_doc)
                        .OfClass(typeof(Autodesk.Revit.DB.View))
                        .Cast<Autodesk.Revit.DB.View>()
                        .Select(v => v.Name)
                );

                using (Transaction trans = new Transaction(_doc, "Duplicate Views"))
                {
                    trans.Start();

                    foreach (var job in jobs)
                    {
                        Autodesk.Revit.DB.View template = job.TemplateIndex > 0 && job.TemplateIndex <= _viewTemplates.Count 
                            ? _viewTemplates[job.TemplateIndex - 1] : null;
                        Element scopeBox = job.ScopeBoxIndex > 0 && job.ScopeBoxIndex <= _scopeBoxes.Count 
                            ? _scopeBoxes[job.ScopeBoxIndex - 1] : null;
                        
                        ViewDuplicateOption dupOption = job.DuplicateIndex switch
                        {
                            1 => ViewDuplicateOption.WithDetailing,
                            2 => ViewDuplicateOption.AsDependent,
                            _ => ViewDuplicateOption.Duplicate
                        };

                        for (int i = 0; i < job.Copies; i++)
                        {
                            string uniqueName = GetRevitStyleUniqueName(job.NewName, existingNames);

                            try
                            {
                                ElementId newId;
                                
                                // Try AsDependent first, fall back to normal Duplicate if it fails
                                if (dupOption == ViewDuplicateOption.AsDependent)
                                {
                                    try
                                    {
                                        newId = job.SourceView.Duplicate(ViewDuplicateOption.AsDependent);
                                    }
                                    catch
                                    {
                                        // Fallback to regular duplicate
                                        newId = job.SourceView.Duplicate(ViewDuplicateOption.Duplicate);
                                    }
                                }
                                else
                                {
                                    newId = job.SourceView.Duplicate(dupOption);
                                }
                                
                                Autodesk.Revit.DB.View newView = _doc.GetElement(newId) as Autodesk.Revit.DB.View;
                                
                                if (newView != null)
                                {
                                    // Set name
                                    try
                                    {
                                        Parameter nameParam = newView.get_Parameter(BuiltInParameter.VIEW_NAME);
                                        if (nameParam != null && !nameParam.IsReadOnly)
                                            nameParam.Set(uniqueName);
                                        else
                                            newView.Name = uniqueName;
                                    }
                                    catch
                                    {
                                        newView.Name = uniqueName;
                                    }

                                    // Apply view template
                                    if (template != null)
                                    {
                                        try
                                        {
                                            if (newView.ViewTemplateId == ElementId.InvalidElementId || 
                                                newView.ViewTemplateId.IntegerValue == -1)
                                            {
                                                newView.ViewTemplateId = template.Id;
                                            }
                                        }
                                        catch { }
                                    }

                                    // Apply scope box
                                    if (scopeBox != null)
                                    {
                                        try
                                        {
                                            Parameter scopeParam = newView.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP);
                                            if (scopeParam == null)
                                                scopeParam = newView.LookupParameter("Scope Box");
                                            
                                            if (scopeParam != null && !scopeParam.IsReadOnly)
                                                scopeParam.Set(scopeBox.Id);
                                        }
                                        catch { }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error duplicating view: {ex.Message}");
                            }
                        }
                    }

                    trans.Commit();
                }

                MessageBox.Show("Views succesvol gedupliceerd!", "Voltooid", MessageBoxButton.OK, MessageBoxImage.Information);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fout bij dupliceren: {ex.Message}", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string GetRevitStyleUniqueName(string baseName, HashSet<string> existing)
        {
            string name = (baseName ?? "").Trim();
            if (string.IsNullOrEmpty(name))
                name = "View";

            if (!existing.Contains(name))
            {
                existing.Add(name);
                return name;
            }

            int i = 1;
            while (true)
            {
                string candidate = $"{name} ({i:D2})";
                if (!existing.Contains(candidate))
                {
                    existing.Add(candidate);
                    return candidate;
                }
                i++;
            }
        }

        // ========== SETS FUNCTIONALITY ==========
        private Dictionary<string, List<SetViewRow>> _setRows = new Dictionary<string, List<SetViewRow>>();
        private int? _setLastIndex = null;

        private void CreateSetTab(string tabName, List<int> viewIds, Dictionary<string, object> defaults)
        {
            // Check if tab already exists
            Button existingButton = null;
            System.Windows.Controls.Grid existingContent = null;
            
            foreach (var child in tabButtons.Children)
            {
                if (child is Button btn && btn.Content.ToString() == tabName)
                {
                    existingButton = btn;
                    if (_savedTabs.ContainsKey(tabName))
                        existingContent = _savedTabs[tabName].ContentGrid;
                    break;
                }
            }

            // Create new pill button if it doesn't exist
            if (existingButton == null)
            {
                var newButton = new Button
                {
                    Content = tabName,
                    Style = (Style)FindResource("PillTabButton")
                };
                
                var content = CreateSetTabContent(tabName, viewIds, defaults);
                _savedTabs[tabName] = new TabData { Button = newButton, ContentGrid = content, ViewIds = viewIds, TabName = tabName };
                
                newButton.Click += (s, e) => SwitchToTab(newButton, content, tabName);
                tabButtons.Children.Add(newButton);
                SwitchToTab(newButton, content, tabName);
            }
            else
            {
                SwitchToTab(existingButton, existingContent, tabName);
            }
        }

        private System.Windows.Controls.Grid CreateSetTabContent(string tabName, List<int> viewIds, Dictionary<string, object> defaults)
        {
            var mainGrid = new System.Windows.Controls.Grid();
            mainGrid.Margin = new Thickness(0);  // No margin - content grid already has padding

            // Column definitions (match main content exactly)
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(12) });
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(520) });

            // Row definitions (match main content exactly)
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // Filters/Settings row
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });  // Views row
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // Selection row

            // Left: Views list with header row (on row 1, not row 0!)
            var viewsGroup = new GroupBox { Header = "Views in set", Margin = new Thickness(0, 0, 0, 4), VerticalAlignment = VerticalAlignment.Top };
            var viewsContainer = new System.Windows.Controls.Grid();
            viewsContainer.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            viewsContainer.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            
            // Header row
            var headerGrid = new System.Windows.Controls.Grid { Margin = new Thickness(4, 0, 4, 6), MinHeight = 32 };
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(24) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.2, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            
            AddHeaderLabel(headerGrid, 0, " ");
            AddHeaderLabel(headerGrid, 1, "Huidige naam");
            AddHeaderLabel(headerGrid, 2, "Nieuwe naam");
            AddHeaderLabel(headerGrid, 3, "Viewtemplate");
            AddHeaderLabel(headerGrid, 4, "Scope Box");
            AddHeaderLabel(headerGrid, 5, "Dupliceer");
            
            System.Windows.Controls.Grid.SetRow(headerGrid, 0);
            viewsContainer.Children.Add(headerGrid);
            
            var listBox = new ListBox { MinHeight = 480 };
            VirtualizingPanel.SetIsVirtualizing(listBox, true);
            VirtualizingPanel.SetVirtualizationMode(listBox, VirtualizationMode.Recycling);
            ScrollViewer.SetCanContentScroll(listBox, true);
            System.Windows.Controls.Grid.SetRow(listBox, 1);
            viewsContainer.Children.Add(listBox);
            
            viewsGroup.Content = viewsContainer;
            System.Windows.Controls.Grid.SetRow(viewsGroup, 1);  // Row 1, not 0!
            System.Windows.Controls.Grid.SetColumn(viewsGroup, 0);
            mainGrid.Children.Add(viewsGroup);

            // Right: Controls
            var rightPanel = CreateSetRightPanel(tabName, listBox, defaults);
            System.Windows.Controls.Grid.SetRow(rightPanel, 1);  // Row 1 to align with Views
            System.Windows.Controls.Grid.SetColumn(rightPanel, 2);
            mainGrid.Children.Add(rightPanel);

            // Bottom: Selection buttons (like main tab)
            var selectionGroup = new GroupBox { Header = "Selectie", Margin = new Thickness(0, 8, 0, 0) };
            var selectionPanel = new StackPanel { Orientation = Orientation.Horizontal };
            
            var btnSelectAll = new Button 
            { 
                Content = "Alles", 
                Margin = new Thickness(0, 0, 8, 0),
                Padding = new Thickness(12, 6, 12, 6),
                Background = Brushes.White,
                Foreground = Brushes.Black,
                BorderBrush = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#D1D5DB")),
                BorderThickness = new Thickness(1, 1, 1, 1),
                Cursor = Cursors.Hand
            };
            
            // Apply rounded corners
            var btnAllTemplate = new ControlTemplate(typeof(Button));
            var btnAllFactory = new FrameworkElementFactory(typeof(Border));
            btnAllFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            btnAllFactory.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
            btnAllFactory.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
            btnAllFactory.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Button.PaddingProperty));
            btnAllFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(4));
            var btnAllContent = new FrameworkElementFactory(typeof(ContentPresenter));
            btnAllContent.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            btnAllContent.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            btnAllFactory.AppendChild(btnAllContent);
            btnAllTemplate.VisualTree = btnAllFactory;
            btnSelectAll.Template = btnAllTemplate;
            
            var btnSelectNone = new Button 
            { 
                Content = "Geen",
                Padding = new Thickness(12, 6, 12, 6),
                Background = Brushes.White,
                Foreground = Brushes.Black,
                BorderBrush = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#D1D5DB")),
                BorderThickness = new Thickness(1, 1, 1, 1),
                Cursor = Cursors.Hand
            };
            
            // Apply rounded corners
            var btnNoneTemplate = new ControlTemplate(typeof(Button));
            var btnNoneFactory = new FrameworkElementFactory(typeof(Border));
            btnNoneFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            btnNoneFactory.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
            btnNoneFactory.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
            btnNoneFactory.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Button.PaddingProperty));
            btnNoneFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(4));
            var btnNoneContent = new FrameworkElementFactory(typeof(ContentPresenter));
            btnNoneContent.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            btnNoneContent.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            btnNoneFactory.AppendChild(btnNoneContent);
            btnNoneTemplate.VisualTree = btnNoneFactory;
            btnSelectNone.Template = btnNoneTemplate;
            
            btnSelectAll.Click += (s, e) => SetSetTabAllChecks(tabName, true);
            btnSelectNone.Click += (s, e) => SetSetTabAllChecks(tabName, false);
            selectionPanel.Children.Add(btnSelectAll);
            selectionPanel.Children.Add(btnSelectNone);
            selectionGroup.Content = selectionPanel;
            System.Windows.Controls.Grid.SetRow(selectionGroup, 2);
            System.Windows.Controls.Grid.SetColumn(selectionGroup, 0);
            mainGrid.Children.Add(selectionGroup);

            // Populate views
            var rows = new List<SetViewRow>();
            foreach (int vid in viewIds)
            {
                Autodesk.Revit.DB.View view = _doc.GetElement(new ElementId(vid)) as Autodesk.Revit.DB.View;
                if (view != null)
                {
                    var row = CreateSetViewRow(view, defaults, listBox);
                    rows.Add(row);
                }
            }
            _setRows[tabName] = rows;

            return mainGrid;
        }

        private StackPanel CreateSetRightPanel(string tabName, ListBox listBox, Dictionary<string, object> defaults)
        {
            var panel = new StackPanel { Orientation = Orientation.Vertical, VerticalAlignment = VerticalAlignment.Stretch };

            // Instellingen dupliceren
            var settingsGroup = new GroupBox { Header = "Instellingen dupliceren", Margin = new Thickness(0, 0, 0, 8), VerticalAlignment = VerticalAlignment.Top };
            var settingsGrid = new System.Windows.Controls.Grid();
            for (int i = 0; i < 4; i++) settingsGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            settingsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            settingsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var cmbTplSet = new ComboBox { Margin = new Thickness(0, 0, 0, 8) };
            var cmbScpSet = new ComboBox { Margin = new Thickness(0, 0, 0, 8) };
            var cmbDupSet = new ComboBox { Margin = new Thickness(0, 0, 0, 8) };
            
            cmbTplSet.Items.Add("<Geen>");
            foreach (var vt in _viewTemplates) cmbTplSet.Items.Add(vt.Name);
            cmbTplSet.SelectedIndex = (int)defaults["tpl_index"];

            cmbScpSet.Items.Add("<Geen>");
            foreach (var sb in _scopeBoxes) cmbScpSet.Items.Add(sb.Name);
            cmbScpSet.SelectedIndex = (int)defaults["scp_index"];

            cmbDupSet.Items.Add("Duplicate");
            cmbDupSet.Items.Add("With Detailing");
            cmbDupSet.Items.Add("As Dependent");
            cmbDupSet.SelectedIndex = (int)defaults["dup_index"];

            AddLabelAndControl(settingsGrid, 0, "Viewtemplate:", cmbTplSet);
            AddLabelAndControl(settingsGrid, 1, "Scope Box:", cmbScpSet);
            AddLabelAndControl(settingsGrid, 2, "Dupliceer:", cmbDupSet);

            var btnApplySettings = new Button { Content = "Toepassen", Style = (Style)Resources["BtnGold"], HorizontalAlignment = HorizontalAlignment.Right };
            btnApplySettings.Click += (s, e) => ApplySettingsToSelected(tabName, cmbTplSet.SelectedIndex, cmbScpSet.SelectedIndex, cmbDupSet.SelectedIndex);
            System.Windows.Controls.Grid.SetRow(btnApplySettings, 3);
            System.Windows.Controls.Grid.SetColumn(btnApplySettings, 1);
            settingsGrid.Children.Add(btnApplySettings);

            settingsGroup.Content = settingsGrid;
            panel.Children.Add(settingsGroup);

            // Naamgeving
            var namingGroup = new GroupBox { Header = "Naamgeving", Margin = new Thickness(0, 8, 0, 8) };
            var namingGrid = new System.Windows.Controls.Grid();
            for (int i = 0; i < 5; i++) namingGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            namingGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            namingGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var txtPreSet = new TextBox { Text = defaults["pre"].ToString(), Margin = new Thickness(0, 0, 0, 8) };
            var txtFindSet = new TextBox { Text = defaults["find"].ToString(), Margin = new Thickness(0, 0, 0, 8) };
            var txtReplSet = new TextBox { Text = defaults["repl"].ToString(), Margin = new Thickness(0, 0, 0, 8) };
            var txtNewSet = new TextBox { Text = defaults["new"].ToString(), Margin = new Thickness(0, 0, 0, 8) };

            AddLabelAndControl(namingGrid, 0, "Prefix:", txtPreSet);
            AddLabelAndControl(namingGrid, 1, "Find:", txtFindSet);
            AddLabelAndControl(namingGrid, 2, "Replace:", txtReplSet);
            AddLabelAndControl(namingGrid, 3, "Suffix:", txtNewSet);

            var btnApplyNaming = new Button { Content = "Toepassen", Style = (Style)Resources["BtnGold"], HorizontalAlignment = HorizontalAlignment.Right };
            btnApplyNaming.Click += (s, e) => ApplyNamingToSelected(tabName, txtPreSet.Text, txtFindSet.Text, txtReplSet.Text, txtNewSet.Text);
            System.Windows.Controls.Grid.SetRow(btnApplyNaming, 4);
            System.Windows.Controls.Grid.SetColumn(btnApplyNaming, 1);
            namingGrid.Children.Add(btnApplyNaming);

            namingGroup.Content = namingGrid;
            panel.Children.Add(namingGroup);

            // Kopieën
            var copiesGroup = new GroupBox { Header = "Kopieën", Margin = new Thickness(0, 8, 0, 8) };
            var copiesGrid = new System.Windows.Controls.Grid();
            for (int i = 0; i < 2; i++) copiesGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            copiesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            copiesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var txtCountSet = new TextBox { Text = defaults["count"].ToString(), Margin = new Thickness(0, 0, 0, 8) };
            AddLabelAndControl(copiesGrid, 0, "Aantal kopieën:", txtCountSet);

            var buttonsPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var btnApplyCopies = new Button { Content = "Toepassen", Style = (Style)Resources["BtnGold"] };
            btnApplyCopies.Click += (s, e) => ApplyCopiesToSelected(tabName, txtCountSet.Text);
            buttonsPanel.Children.Add(btnApplyCopies);
            System.Windows.Controls.Grid.SetRow(buttonsPanel, 1);
            System.Windows.Controls.Grid.SetColumn(buttonsPanel, 1);
            copiesGrid.Children.Add(buttonsPanel);

            copiesGroup.Content = copiesGrid;
            panel.Children.Add(copiesGroup);

            return panel;
        }

        private void AddLabelAndControl(System.Windows.Controls.Grid grid, int row, string labelText, UIElement control)
        {
            var label = new Label { Content = labelText, VerticalAlignment = VerticalAlignment.Center };
            System.Windows.Controls.Grid.SetRow(label, row);
            System.Windows.Controls.Grid.SetColumn(label, 0);
            grid.Children.Add(label);

            System.Windows.Controls.Grid.SetRow(control, row);
            System.Windows.Controls.Grid.SetColumn(control, 1);
            grid.Children.Add(control);
        }

        private void AddHeaderLabel(System.Windows.Controls.Grid grid, int column, string text)
        {
            var label = new TextBlock 
            { 
                Text = text, 
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(4, 4, 4, 4),
                TextWrapping = TextWrapping.Wrap,
                FontSize = 12
            };
            System.Windows.Controls.Grid.SetColumn(label, column);
            grid.Children.Add(label);
        }

        private SetViewRow CreateSetViewRow(Autodesk.Revit.DB.View view, Dictionary<string, object> defaults, ListBox listBox)
        {
            var row = new System.Windows.Controls.Grid { Margin = new Thickness(4, 6, 4, 6) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(24) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.2, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var chk = new CheckBox { IsChecked = true, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(2, 0, 2, 0) };
            System.Windows.Controls.Grid.SetColumn(chk, 0);
            row.Children.Add(chk);

            var txtOldName = new TextBlock { Text = view.Name, Margin = new Thickness(2, 3, 2, 3), TextTrimming = TextTrimming.CharacterEllipsis, VerticalAlignment = VerticalAlignment.Center };
            System.Windows.Controls.Grid.SetColumn(txtOldName, 1);
            row.Children.Add(txtOldName);

            // FlatText style for new name textbox
            var txtNewName = new TextBox 
            { 
                Text = view.Name, 
                Margin = new Thickness(2, 3, 2, 3),
                Height = 26,
                Padding = new Thickness(8, 0, 8, 0),
                HorizontalContentAlignment = HorizontalAlignment.Left,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderBrush = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#D1D5DB")),
                BorderThickness = new Thickness(1),
                Background = Brushes.White
            };
            System.Windows.Controls.Grid.SetColumn(txtNewName, 2);
            row.Children.Add(txtNewName);

            var cmbTpl = new ComboBox { Margin = new Thickness(2) };
            cmbTpl.Items.Add("<Geen>");
            foreach (var vt in _viewTemplates) cmbTpl.Items.Add(vt.Name);
            cmbTpl.SelectedIndex = (int)defaults["tpl_index"];
            System.Windows.Controls.Grid.SetColumn(cmbTpl, 3);
            row.Children.Add(cmbTpl);

            var cmbScp = new ComboBox { Margin = new Thickness(2) };
            cmbScp.Items.Add("<Geen>");
            foreach (var sb in _scopeBoxes) cmbScp.Items.Add(sb.Name);
            cmbScp.SelectedIndex = (int)defaults["scp_index"];
            System.Windows.Controls.Grid.SetColumn(cmbScp, 4);
            row.Children.Add(cmbScp);

            var cmbDup = new ComboBox { Margin = new Thickness(2) };
            cmbDup.Items.Add("Duplicate");
            cmbDup.Items.Add("With Detailing");
            cmbDup.Items.Add("As Dependent");
            cmbDup.SelectedIndex = (int)defaults["dup_index"];
            System.Windows.Controls.Grid.SetColumn(cmbDup, 5);
            row.Children.Add(cmbDup);

            listBox.Items.Add(row);

            var setRow = new SetViewRow
            {
                ViewId = view.Id.IntegerValue,
                OriginalName = view.Name,
                Checkbox = chk,
                NewNameTextBox = txtNewName,
                TemplateCombo = cmbTpl,
                ScopeBoxCombo = cmbScp,
                DuplicateCombo = cmbDup,
                CopiesCount = (int)defaults["count"]
            };

            // Row click toggles checkbox
            row.MouseLeftButtonUp += (s, e) => 
            {
                chk.IsChecked = !chk.IsChecked;
                e.Handled = true;
            };

            // Shift-click for range selection
            chk.Click += (s, e) =>
            {
                int cur = listBox.Items.IndexOf(row);
                if (Keyboard.Modifiers == ModifierKeys.Shift && _setLastIndex.HasValue && cur >= 0)
                {
                    int a = Math.Min(_setLastIndex.Value, cur);
                    int b = Math.Max(_setLastIndex.Value, cur);
                    bool val = chk.IsChecked ?? false;
                    
                    for (int idx = a; idx <= b; idx++)
                    {
                        var itemRow = listBox.Items[idx] as System.Windows.Controls.Grid;
                        if (itemRow != null)
                        {
                            foreach (var child in itemRow.Children)
                            {
                                if (child is CheckBox cb)
                                {
                                    cb.IsChecked = val;
                                    break;
                                }
                            }
                        }
                    }
                }
                if (cur >= 0) _setLastIndex = cur;
            };

            return setRow;
        }

        private void ApplySettingsToSelected(string tabName, int tplIdx, int scpIdx, int dupIdx)
        {
            if (!_setRows.ContainsKey(tabName)) return;
            
            foreach (var row in _setRows[tabName].Where(r => r.Checkbox.IsChecked ?? false))
            {
                row.TemplateCombo.SelectedIndex = tplIdx;
                row.ScopeBoxCombo.SelectedIndex = scpIdx;
                row.DuplicateCombo.SelectedIndex = dupIdx;
            }
        }

        private void ApplyNamingToSelected(string tabName, string prefix, string find, string replace, string suffix)
        {
            if (!_setRows.ContainsKey(tabName)) return;

            foreach (var row in _setRows[tabName].Where(r => r.Checkbox.IsChecked ?? false))
            {
                string baseName = row.OriginalName;
                if (!string.IsNullOrEmpty(find))
                    baseName = baseName.Replace(find, replace ?? "");
                string renamed = Clean(prefix + baseName + suffix);
                int copies = row.CopiesCount;
                row.NewNameTextBox.Text = renamed + (copies <= 1 ? "" : $"  (×{copies})");
            }
        }

        private void ApplyCopiesToSelected(string tabName, string countText)
        {
            if (!_setRows.ContainsKey(tabName)) return;
            int count = ParseCount(countText);
            
            foreach (var row in _setRows[tabName].Where(r => r.Checkbox.IsChecked ?? false))
            {
                row.CopiesCount = count;
            }
        }

        private void SetSetTabAllChecks(string tabName, bool value)
        {
            if (!_setRows.ContainsKey(tabName)) return;

            foreach (var row in _setRows[tabName])
            {
                row.Checkbox.IsChecked = value;
            }
        }

        private void RemoveSetTab(string tabName)
        {
            if (_savedTabs.ContainsKey(tabName))
            {
                var tabData = _savedTabs[tabName];
                tabButtons.Children.Remove(tabData.Button);
                _savedTabs.Remove(tabName);
                _setRows.Remove(tabName);
                
                // Switch back to main tab
                if (tabButtons.Children.Count > 0)
                {
                    var mainButton = tabButtons.Children[0] as Button;
                    SwitchToTab(mainButton, null, null);
                }
                btnRemoveTab.Visibility = System.Windows.Visibility.Collapsed;
            }
        }
    }

    internal class ViewItem
    {
        public string Name { get; set; }
        public Autodesk.Revit.DB.View View { get; set; }
    }

    internal class CheckBoxData
    {
        public string Name { get; set; }
        public Autodesk.Revit.DB.View View { get; set; }
        public TextBlock Preview { get; set; }
        public System.Windows.Controls.Grid Row { get; set; }
    }

    internal class SetViewRow
    {
        public int ViewId { get; set; }
        public string OriginalName { get; set; }
        public CheckBox Checkbox { get; set; }
        public TextBox NewNameTextBox { get; set; }
        public ComboBox TemplateCombo { get; set; }
        public ComboBox ScopeBoxCombo { get; set; }
        public ComboBox DuplicateCombo { get; set; }
        public int CopiesCount { get; set; }
    }

    internal class DuplicationJob
    {
        public Autodesk.Revit.DB.View SourceView { get; set; }
        public string NewName { get; set; }
        public int TemplateIndex { get; set; }
        public int ScopeBoxIndex { get; set; }
        public int DuplicateIndex { get; set; }
        public int Copies { get; set; }
    }
}
