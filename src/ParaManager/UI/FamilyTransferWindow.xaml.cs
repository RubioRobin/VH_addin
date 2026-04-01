using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using ParaManager.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;

namespace ParaManager.UI
{
    public partial class FamilyTransferWindow : Window, IExternalEventHandler
    {
        private UIDocument _uidoc;
        private Document _doc;
        private Document _sourceFamilyDoc;
        private List<FamilyParameterData> _sourceParameters;
        private List<Family> _destinationFamilies;
        private List<Document> _destinationDocs;
        private HashSet<string> _originallyOpenFamilies = new HashSet<string>();
        
        private ExternalEvent _externalEvent;
        private List<FamilyParameterData> _pendingParams;

        public FamilyTransferWindow(UIDocument uidoc)
        {
            InitializeComponent();
            MessageBox.Show("DEBUG: NIEUWE ParaManager build geladen op " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "ParaManager Debug", MessageBoxButton.OK, MessageBoxImage.Information);
            _uidoc = uidoc;
            _doc = _uidoc.Document;
            _sourceParameters = new List<FamilyParameterData>();
            _destinationFamilies = new List<Family>();
            _destinationDocs = new List<Document>();
            
            _externalEvent = ExternalEvent.Create(this);

            // Set Revit as owner
            WindowInteropHelper helper = new WindowInteropHelper(this);
            helper.Owner = Autodesk.Windows.ComponentManager.ApplicationWindow;

            // Track all currently open families
            foreach (Document d in _doc.Application.Documents)
            {
                if (d.IsFamilyDocument)
                    _originallyOpenFamilies.Add(d.Title);
            }

            // If active document is a family, set it as source automatically
            if (_doc.IsFamilyDocument)
            {
                txtSourceFamily.Text = _doc.Title;
                _sourceFamilyDoc = _doc;
                LoadSourceParameters(_sourceFamilyDoc);
            }
        }

        public void Execute(UIApplication app)
        {
            if (_pendingParams != null)
            {
                DoTransferInternal(_pendingParams);
                _pendingParams = null;
            }
        }

        public string GetName()
        {
            return "Family Parameter Transfer";
        }

        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            UpdateUiMode();
        }

        private void UpdateUiMode()
        {
            // Ensure UI elements are loaded
            if (grpSourceFamily == null || grpExcelSource == null || grpDestination == null || btnAction == null) return;

            if (rbModeTransfer.IsChecked == true)
            {
                // Transfer Mode
                grpSourceFamily.Visibility = System.Windows.Visibility.Visible;
                grpExcelSource.Visibility = System.Windows.Visibility.Collapsed;
                grpDestination.Visibility = System.Windows.Visibility.Visible;
                grpDestination.Opacity = 1;
                grpDestination.IsEnabled = true;
                btnAction.Content = "TRANSFER PARAMETERS";
                
                // If switching back to transfer, we might want to clear excel source text or keep it separate?
                // For now, let's just show family source
                if (txtSourceFamily.Text.StartsWith("Excel:")) txtSourceFamily.Text = "";
            }
            else if (rbModeImport.IsChecked == true)
            {
                // Import Mode
                grpSourceFamily.Visibility = System.Windows.Visibility.Collapsed;
                grpExcelSource.Visibility = System.Windows.Visibility.Visible;
                grpDestination.Visibility = System.Windows.Visibility.Visible;
                btnAction.Content = "IMPORT TO FAMILIES";
            }
            else if (rbModeExport.IsChecked == true)
            {
                // Export Mode
                grpSourceFamily.Visibility = System.Windows.Visibility.Visible;
                grpExcelSource.Visibility = System.Windows.Visibility.Collapsed;
                // We hide or disable destination for export
                grpDestination.Visibility = System.Windows.Visibility.Collapsed;
                btnAction.Content = "EXPORT TO EXCEL";
            }
        }
        private void SelectSourceFamily_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Hide();

                try
                {
                    var reference = _uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select a family instance to load parameters from");
                    if (reference != null)
                    {
                        Element el = _doc.GetElement(reference);
                        Family family = null;

                        if (el is FamilyInstance fi)
                        {
                            family = fi.Symbol.Family;
                        }
                        else if (el is Family f)
                        {
                            family = f;
                        }

                        if (family != null)
                        {
                            txtSourceFamily.Text = family.Name;
                            
                            if (_sourceFamilyDoc != null && _sourceFamilyDoc.IsFamilyDocument && _sourceFamilyDoc.Title != _doc.Title)
                            {
                                _sourceFamilyDoc.Close(false);
                            }

                            _sourceFamilyDoc = _doc.EditFamily(family);
                            LoadSourceParameters(_sourceFamilyDoc);
                        }
                        else
                        {
                            MessageBox.Show("Please select a Family or Family Instance.", "Selection Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    // User canceled selection
                }
                finally
                {
                    this.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error selecting source family: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Show();
            }
        }

        private void LoadSourceParameters(Document familyDoc)
        {
            try
            {
                _sourceParameters.Clear();

                if (familyDoc != null && familyDoc.IsFamilyDocument)
                {
                    FamilyManager famMgr = familyDoc.FamilyManager;

                    foreach (FamilyParameter param in famMgr.Parameters)
                    {
                        FamilyParameterData paramData = new FamilyParameterData
                        {
                            Name = param.Definition.Name,
                            ParameterType = ParaManager.Helpers.ParameterTypeHelper.GetParameterTypeString(param.Definition),
                            IsInstance = param.IsInstance,
                            IsShared = param.IsShared,
                            GUID = param.IsShared ? param.GUID.ToString() : null,
                            Formula = param.Formula ?? "",
#if REVIT2025
                            Group = param.Definition.GetGroupTypeId()
#else
                            Group = param.Definition.ParameterGroup
#endif
                        };

                        _sourceParameters.Add(paramData);
                    }
                }

                RenderParameterList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading source parameters: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FilterParameters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RenderParameterList();
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RenderParameterList();
        }

        private void RenderParameterList()
        {
            if (lvParameters == null) return;

            // Populate filter dropdowns if needed
            PopulateFilterDropdowns();

            if (_sourceParameters.Count == 0)
            {
                lvParameters.ItemsSource = null;
                UpdateSelectionCount();
                return;
            }

            // Apply all filters
            var filteredParams = _sourceParameters.AsEnumerable();

            // Search filter
            if (txtSearch != null && !string.IsNullOrWhiteSpace(txtSearch.Text))
            {
                string searchText = txtSearch.Text.ToLower();
                filteredParams = filteredParams.Where(p => p.Name.ToLower().Contains(searchText));
            }

            // Scope filter (Instance/Type/Shared/Family)
            if (cmbFilter != null && cmbFilter.SelectedItem is ComboBoxItem selectedItem)
            {
                string filter = selectedItem.Content.ToString();
                switch (filter)
                {
                    case "Instance Parameters":
                        filteredParams = filteredParams.Where(p => p.IsInstance);
                        break;
                    case "Type Parameters":
                        filteredParams = filteredParams.Where(p => !p.IsInstance);
                        break;
                    case "Shared Parameters":
                        filteredParams = filteredParams.Where(p => p.IsShared);
                        break;
                    case "Family Parameters":
                        filteredParams = filteredParams.Where(p => !p.IsShared);
                        break;
                }
            }

            // Type filter
            if (cmbTypeFilter != null && cmbTypeFilter.SelectedItem is ComboBoxItem typeItem)
            {
                string typeFilter = typeItem.Content.ToString();
                if (typeFilter != "All Types")
                {
                    filteredParams = filteredParams.Where(p => p.ParameterType == typeFilter);
                }
            }

            // Group filter
            if (cmbGroupFilter != null && cmbGroupFilter.SelectedItem is ComboBoxItem groupItem)
            {
                string groupFilter = groupItem.Content.ToString();
                if (groupFilter != "All Groups")
                {
                    filteredParams = filteredParams.Where(p => p.GroupName == groupFilter);
                }
            }

            lvParameters.ItemsSource = filteredParams.ToList();
            UpdateSelectionCount();
        }

        private void PopulateFilterDropdowns()
        {
            if (_sourceParameters.Count == 0) return;

            // Populate Type filter
            if (cmbTypeFilter != null && cmbTypeFilter.Items.Count <= 1)
            {
                var types = _sourceParameters.Select(p => p.ParameterType).Distinct().OrderBy(t => t);
                foreach (var type in types)
                {
                    cmbTypeFilter.Items.Add(new ComboBoxItem { Content = type });
                }
            }

            // Populate Group filter
            if (cmbGroupFilter != null && cmbGroupFilter.Items.Count <= 1)
            {
                var groups = _sourceParameters.Select(p => p.GroupName).Distinct().OrderBy(g => g);
                foreach (var group in groups)
                {
                    cmbGroupFilter.Items.Add(new ComboBoxItem { Content = group });
                }
            }
        }

        private void UpdateSelectionCount()
        {
            if (txtSelectionCount == null) return;

            int selectedCount = _sourceParameters.Count(p => p.IsSelected);
            int totalCount = _sourceParameters.Count;
            txtSelectionCount.Text = $"{selectedCount} / {totalCount} selected";

            // Update header checkbox state
            if (chkSelectAllHeader != null)
            {
                var visibleParams = lvParameters?.ItemsSource as IEnumerable<FamilyParameterData>;
                if (visibleParams != null && visibleParams.Any())
                {
                    bool allSelected = visibleParams.All(p => p.IsSelected);
                    bool noneSelected = visibleParams.All(p => !p.IsSelected);
                    chkSelectAllHeader.IsChecked = allSelected ? true : (noneSelected ? false : null);
                }
            }
        }

        private void ParameterCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            UpdateSelectionCount();
        }

        private void ListViewItem_PreviewMouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // Toggle selection when clicking on the row (but not directly on the checkbox as it handles itself)
            if (e.OriginalSource is CheckBox) return;

            if (sender is ListViewItem item && item.DataContext is FamilyParameterData paramData)
            {
                paramData.IsSelected = !paramData.IsSelected;
                UpdateSelectionCount();
            }
        }

        private void SelectAllHeader_Checked(object sender, RoutedEventArgs e)
        {
            SetVisibleParametersSelection(true);
        }

        private void SelectAllHeader_Unchecked(object sender, RoutedEventArgs e)
        {
            SetVisibleParametersSelection(false);
        }

        private void SetVisibleParametersSelection(bool selected)
        {
            var visibleParams = lvParameters?.ItemsSource as IEnumerable<FamilyParameterData>;
            if (visibleParams != null)
            {
                foreach (var param in visibleParams)
                {
                    param.IsSelected = selected;
                }
                lvParameters.Items.Refresh();
                UpdateSelectionCount();
            }
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            SetVisibleParametersSelection(true);
            if (chkSelectAllHeader != null) chkSelectAllHeader.IsChecked = true;
        }

        private void SelectNone_Click(object sender, RoutedEventArgs e)
        {
            SetVisibleParametersSelection(false);
            if (chkSelectAllHeader != null) chkSelectAllHeader.IsChecked = false;
        }

        private void InvertSelection_Click(object sender, RoutedEventArgs e)
        {
            var visibleParams = lvParameters?.ItemsSource as IEnumerable<FamilyParameterData>;
            if (visibleParams != null)
            {
                foreach (var param in visibleParams)
                {
                    param.IsSelected = !param.IsSelected;
                }
                lvParameters.Items.Refresh();
                UpdateSelectionCount();
            }
        }

        private void SavePreset_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedNames = _sourceParameters.Where(p => p.IsSelected).Select(p => p.Name).ToList();
                if (selectedNames.Count == 0)
                {
                    MessageBox.Show("No parameters selected to save.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                Microsoft.Win32.SaveFileDialog saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "JSON Files (*.json)|*.json",
                    Title = "Save Parameter Selection Preset",
                    FileName = "ParameterPreset.json"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var preset = new ParameterSelectionPreset
                    {
                        Name = Path.GetFileNameWithoutExtension(saveDialog.FileName),
                        SelectedParameterNames = selectedNames,
                        CreatedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    };

                    string json = System.Text.Json.JsonSerializer.Serialize(preset, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(saveDialog.FileName, json);
                    MessageBox.Show($"Preset saved with {selectedNames.Count} parameters.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving preset: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadPreset_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog openDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "JSON Files (*.json)|*.json",
                    Title = "Load Parameter Selection Preset"
                };

                if (openDialog.ShowDialog() == true)
                {
                    string json = File.ReadAllText(openDialog.FileName);
                    var preset = System.Text.Json.JsonSerializer.Deserialize<ParameterSelectionPreset>(json);

                    if (preset?.SelectedParameterNames == null || preset.SelectedParameterNames.Count == 0)
                    {
                        MessageBox.Show("No parameters found in preset.", "Info", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // Apply preset
                    int matchedCount = 0;
                    foreach (var param in _sourceParameters)
                    {
                        bool shouldSelect = preset.SelectedParameterNames.Contains(param.Name);
                        param.IsSelected = shouldSelect;
                        if (shouldSelect) matchedCount++;
                    }

                    lvParameters.Items.Refresh();
                    UpdateSelectionCount();
                    MessageBox.Show($"Preset applied: {matchedCount} of {preset.SelectedParameterNames.Count} parameters matched.", "Preset Loaded", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading preset: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearFilters_Click(object sender, RoutedEventArgs e)
        {
            // Reset search
            if (txtSearch != null) txtSearch.Text = "";

            // Reset scope filter
            if (cmbFilter != null) cmbFilter.SelectedIndex = 0;

            // Reset type filter (keep only first item)
            if (cmbTypeFilter != null)
            {
                while (cmbTypeFilter.Items.Count > 1)
                    cmbTypeFilter.Items.RemoveAt(1);
                cmbTypeFilter.SelectedIndex = 0;
            }

            // Reset group filter (keep only first item)
            if (cmbGroupFilter != null)
            {
                while (cmbGroupFilter.Items.Count > 1)
                    cmbGroupFilter.Items.RemoveAt(1);
                cmbGroupFilter.SelectedIndex = 0;
            }

            RenderParameterList();
        }

        private void AddDestinationFamilies_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Hide();

                try
                {
                    var references = _uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select family instances as targets");
                    if (references != null && references.Count > 0)
                    {
                        foreach (var reference in references)
                        {
                            Element el = _doc.GetElement(reference);
                            Family family = null;

                            if (el is FamilyInstance fi)
                            {
                                family = fi.Symbol.Family;
                            }
                            else if (el is Family f)
                            {
                                family = f;
                            }

                            if (family != null && !_destinationFamilies.Any(df => df.Id == family.Id))
                            {
                                _destinationFamilies.Add(family);
                                lstDestinationFamilies.Items.Add(family.Name);
                            }
                        }
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                }
                finally
                {
                    this.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding destination families: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Show();
            }
        }

        private void AddOpenFamilies_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFamilies = _doc.Application.Documents.Cast<Document>()
                    .Where(d => d.IsFamilyDocument && d.Title != _doc.Title)
                    .ToList();

                if (openFamilies.Count == 0)
                {
                    MessageBox.Show("No other family documents are currently open.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                foreach (var famDoc in openFamilies)
                {
                    if (!_destinationDocs.Any(d => d.Title == famDoc.Title))
                    {
                        _destinationDocs.Add(famDoc);
                        _originallyOpenFamilies.Add(famDoc.Title);
                        lstDestinationFamilies.Items.Add($"(Open) {famDoc.Title}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding open families: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RemoveDestinationFamily_Click(object sender, RoutedEventArgs e)
        {
            if (lstDestinationFamilies.SelectedIndex >= 0)
            {
                int index = lstDestinationFamilies.SelectedIndex;
                string itemText = lstDestinationFamilies.Items[index].ToString();

                if (itemText.StartsWith("(Open) "))
                {
                    string title = itemText.Replace("(Open) ", "");
                    _destinationDocs.RemoveAll(d => d.Title == title);
                }
                else
                {
                    _destinationFamilies.RemoveAll(f => f.Name == itemText);
                }
                
                lstDestinationFamilies.Items.RemoveAt(index);
            }
        }

        private void ExecuteAction_Click(object sender, RoutedEventArgs e)
        {
            if (rbModeExport.IsChecked == true)
            {
                ExportToExcel_Logic();
            }
            else
            {
                // Transfer or Import Mode -> Same core transfer logic
                TransferParameters_Logic();
            }
        }

        private void TransferParameters_Logic()
        {
            var selectedParams = GetSelectedParameters();
            if (selectedParams.Count == 0)
            {
                MessageBox.Show("Please select at least one parameter.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_destinationFamilies.Count == 0 && _destinationDocs.Count == 0)
            {
                MessageBox.Show("Please add at least one destination family document or instance.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _pendingParams = selectedParams;
            _externalEvent.Raise();
        }

        private List<FamilyParameterData> GetSelectedParameters()
        {
            // Use the IsSelected property from the data model
            return _sourceParameters.Where(p => p.IsSelected).ToList();
        }

        private void DoTransferInternal(List<FamilyParameterData> selectedParams)
        {
            try
            {
                bool overwrite = chkOverwriteExisting.IsChecked == true;
                bool skipErrors = chkSkipErrors.IsChecked == true;

                int successCount = 0;
                int failCount = 0;
                List<string> errors = new List<string>();

                // 1. Process Families picked from the model (will be edited, loaded back, and closed if not originally open)
                foreach (Family family in _destinationFamilies)
                {
                    try
                    {
                        Document destDoc = _doc.EditFamily(family);
                        if (destDoc != null)
                        {
                            bool paramSuccess = ProcessFamilyDoc(destDoc, selectedParams, overwrite, skipErrors, errors, family.Name);
                            
                            // Load back if we are in a project
                            if (!_doc.IsFamilyDocument)
                            {
                                destDoc.LoadFamily(_doc, new FamilyLoadOptions());
                            }

                            if (!_originallyOpenFamilies.Contains(family.Name))
                            {
                                destDoc.Close(false);
                            }

                            if (paramSuccess) successCount++;
                            else failCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        errors.Add($"{family.Name}: {ex.Message}");
                    }
                }

                // 2. Process Families that are already open in the editor
                foreach (Document destDoc in _destinationDocs)
                {
                    try
                    {
                        bool paramSuccess = ProcessFamilyDoc(destDoc, selectedParams, overwrite, skipErrors, errors, destDoc.Title);
                        if (paramSuccess) successCount++;
                        else failCount++;
                        
                        // We do NOT close or load back automatically here, 
                        // as the user is working on these documents.
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        errors.Add($"{destDoc.Title}: {ex.Message}");
                    }
                }

                string resultMessage = $"Transfer completed!\n\nSuccessful: {successCount}\nFailed: {failCount}";
                if (errors.Count > 0)
                {
                    resultMessage += $"\n\nErrors:\n{string.Join("\n", errors.Take(10))}";
                    if (errors.Count > 10)
                        resultMessage += $"\n... and {errors.Count - 10} more errors";
                }

                MessageBox.Show(resultMessage, "Transfer Results", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error transferring parameters: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ProcessFamilyDoc(Document destDoc, List<FamilyParameterData> selectedParams, bool overwrite, bool skipErrors, List<string> errors, string familyName)
        {
            try
            {
                using (Transaction trans = new Transaction(destDoc, "Transfer Parameters"))
                {
                    trans.Start();

                    FamilyManager famMgr = destDoc.FamilyManager;

                    foreach (var param in selectedParams)
                    {
                        try
                        {
                            FamilyParameter existingParam = famMgr.get_Parameter(param.Name);

                            if (existingParam != null && !overwrite)
                                continue;

                            if (existingParam == null)
                            {
                                ForgeTypeId paramTypeId = ParaManager.Helpers.ParameterTypeHelper.GetForgeTypeId(param.ParameterType);
#if REVIT2025
                                ForgeTypeId groupTypeId = param.Group;
#else
                                ForgeTypeId groupTypeId = GroupTypeId.IdentityData;
                                try
                                {
                                    var method = typeof(ParameterUtils).GetMethod("GetGroupTypeId", new[] { typeof(BuiltInParameterGroup) });
                                    if (method != null)
                                    {
                                        groupTypeId = (ForgeTypeId)method.Invoke(null, new object[] { param.Group });
                                    }
                                }
                                catch { }
#endif

                                // HANDLE SHARED PARAMETERS
                                if (param.IsShared && !string.IsNullOrEmpty(param.GUID))
                                {
                                    ExternalDefinition extDef = ParaManager.Helpers.SharedParameterHelper.GetExternalDefinitionByGuid(destDoc.Application, param.GUID);
                                    if (extDef != null)
                                    {
                                        famMgr.AddParameter(extDef, groupTypeId, param.IsInstance);
                                    }
                                    else
                                    {
                                        errors.Add($"{familyName}: Shared Parameter '{param.Name}' (GUID: {param.GUID}) not found in current Shared Parameter File. Skipped.");
                                        continue;
                                    }
                                }
                                else
                                {
                                    // FAMILY PARAMETER
#if REVIT2025
                                    famMgr.AddParameter(param.Name, param.Group, paramTypeId, param.IsInstance);
#else
                                    famMgr.AddParameter(param.Name, groupTypeId, paramTypeId, param.IsInstance);
#endif
                                }
                            }
                            else if (overwrite)
                            {
                                if (!string.IsNullOrEmpty(param.Formula))
                                    famMgr.SetFormula(existingParam, param.Formula);
                            }
                        }
                        catch (Exception paramEx)
                        {
                            if (!skipErrors)
                                throw;
                            errors.Add($"{familyName}: {param.Name} - {paramEx.Message}");
                        }
                    }

                    trans.Commit();
                    return true;
                }
            }
            catch (Exception ex)
            {
                errors.Add($"{familyName}: {ex.Message}");
                return false;
            }
        }



        private void TitleBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ClickCount == 1)
            {
                DragMove();
            }
        }

        private void ExportToExcel_Logic()
        {
            try
            {
                var parameters = GetSelectedParameters();
                if (parameters.Count == 0)
                {
                    MessageBox.Show("Please select parameters to export.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Microsoft.Win32.SaveFileDialog saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    Title = "Export Family Parameters to Excel",
                    FileName = "FamilyParameters.xlsx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    // Convert to ParameterData for ExcelHelper
                    List<ParameterData> data = parameters.Select(p => new ParameterData
                    {
                        Name = p.Name,
                        ParameterType = p.ParameterType,
                        IsInstance = p.IsInstance,
#if REVIT2025
                        Group = p.Group,
#else
                        Group = p.Group,
#endif
                        IsShared = p.IsShared,
                        GUID = p.GUID
                    }).ToList();

                    ParaManager.Helpers.ExcelHelper.ExportParametersToExcel(data, saveDialog.FileName);
                    MessageBox.Show("Export successful!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BrowseExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog openDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    Title = "Load Parameters from Excel"
                };

                if (openDialog.ShowDialog() == true)
                {
                    var data = ParaManager.Helpers.ExcelHelper.ImportParametersFromExcel(openDialog.FileName);
                    if (data == null || data.Count == 0)
                    {
                        MessageBox.Show("No parameters found in the Excel file.", "Import Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // Clear any loaded family source to avoid confusion
                    _sourceFamilyDoc = null;
                    txtSourceFamily.Text = "";

                    txtSourceExcel.Text = Path.GetFileName(openDialog.FileName);
                    _sourceParameters.Clear();

                    foreach (var p in data)
                    {
                        _sourceParameters.Add(new FamilyParameterData
                        {
                            Name = p.Name,
                            ParameterType = p.ParameterType,
                            IsInstance = p.IsInstance,
#if REVIT2025
                            Group = p.Group,
#else
                            Group = p.Group,
#endif
                            IsShared = p.IsShared,
                            GUID = p.GUID
                        });
                    }

                    RenderParameterList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading from Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    /// <summary>
    /// Helper class to handle family loading options without UI prompts
    /// </summary>
    public class FamilyLoadOptions : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;
        }

        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = true;
            return true;
        }
    }
}
