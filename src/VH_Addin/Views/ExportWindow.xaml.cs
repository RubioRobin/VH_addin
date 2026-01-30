using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using VH_Tools.Utilities;
using VH_Addin.Models;
using System.Text.Json;
using System.Diagnostics;

namespace VH_Addin.Views
{
    public partial class ExportWindow : Window, INotifyPropertyChanged
    {
        private readonly Document _doc;
        private string _targetFolder;
        private ExportSettings _settings;
        private List<ViewSheetSet> _printSets;
        private List<ExportDWGSettings> _dwgSetups;
        private System.ComponentModel.ICollectionView _sheetsView;

        public class SheetItem : INotifyPropertyChanged
        {
            private bool _isSelected = true;
            public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); } }
            public ViewSheet Sheet { get; set; }
            public string SheetNumber => Sheet.SheetNumber;
            public string Name => Sheet.Name;
            public string DetectedSize { get; set; }
            public string Orientation { get; set; }
            public double WidthMM { get; set; }
            public double HeightMM { get; set; }

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
        }

        public class NamingPart : INotifyPropertyChanged
        {
            public bool IsParameter { get; set; }
            public bool IsDate { get; set; }
            public string Name { get; set; } // Parameter name or static text
            public string DisplayName => IsParameter ? $"[{Name}]" : (IsDate ? "[Datum:JJMMDD]" : Name);

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
        }

        public class ParameterItem
        {
            public string Name { get; set; }
            public ElementId Id { get; set; }
        }

        public ObservableCollection<SheetItem> Sheets { get; set; } = new ObservableCollection<SheetItem>();
        public ObservableCollection<SheetItem> OrderedSheets { get; set; } = new ObservableCollection<SheetItem>();
        public ObservableCollection<NamingPart> NamingRule { get; set; } = new ObservableCollection<NamingPart>();
        public List<ParameterItem> AvailableParameters { get; set; } = new List<ParameterItem>();

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));

        public ExportWindow(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            DataContext = this;
            
            LoadSettings();
            LoadSheets();
            LoadPrintSets();
            LoadDwgSetups();
            LoadParameters();
            
            // Apply Settings to UI
            ApplySettingsToUI();
            
            UpdatePreview();

            _sheetsView = System.Windows.Data.CollectionViewSource.GetDefaultView(Sheets);
            _sheetsView.Filter = SheetFilter;
        }

        private void LoadSettings()
        {
            try
            {
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VH_Engineering");
                string file = Path.Combine(folder, "ExportSettings.json");
                
                if (File.Exists(file))
                {
                    string json = File.ReadAllText(file);
                    _settings = JsonSerializer.Deserialize<ExportSettings>(json);
                }
            }
            catch { }

            if (_settings == null) _settings = new ExportSettings();

            // Defaults
            if (string.IsNullOrEmpty(_settings.LastExportFolder) || !Directory.Exists(_settings.LastExportFolder))
                _settings.LastExportFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            _targetFolder = _settings.LastExportFolder;

            // Naming Rule Default
            if (_settings.NamingRule == null || !_settings.NamingRule.Any())
            {
                _settings.NamingRule = new List<string> { "[Sheet Number]", "_", "[Sheet Name]" };
            }
        }

        private void SaveSettings()
        {
            try
            {
                // Update settings object from UI state
                _settings.LastExportFolder = _targetFolder;
                _settings.OpenFolderAfterExport = chkOpenFolder.IsChecked == true;
                _settings.AutoSeparator = chkAutoSeparator.IsChecked == true;
                _settings.Separator = txtAutoSeparator.Text;
                _settings.ExportPdf = chkExportPdf.IsChecked == true;
                _settings.ExportDwg = chkExportDwg.IsChecked == true;
                _settings.CombinePdf = chkCombinePdf.IsChecked == true;
                _settings.ColorMode = cmbColorMode.SelectedIndex;
                _settings.RasterQuality = cmbRasterQuality.SelectedIndex;
                
                if (cmbDwgSetup.SelectedItem is ExportDWGSettings dwgSet)
                {
                    _settings.LastDwgSetupName = dwgSet.Name;
                }

                _settings.NamingRule.Clear();
                foreach (var part in NamingRule)
                {
                    if (part.IsParameter) _settings.NamingRule.Add($"[{part.Name}]");
                    else if (part.IsDate) _settings.NamingRule.Add("[Datum:JJMMDD]");
                    else _settings.NamingRule.Add(part.Name);
                }

                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VH_Engineering");
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                string file = Path.Combine(folder, "ExportSettings.json");
                
                string json = JsonSerializer.Serialize(_settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(file, json);
            }
            catch (Exception ex)
            {
                // Log or ignore
            }
        }

        private void ApplySettingsToUI()
        {
            txtFolderPath.Text = _targetFolder;
            chkOpenFolder.IsChecked = _settings.OpenFolderAfterExport;
            
            chkAutoSeparator.IsChecked = _settings.AutoSeparator;
            txtAutoSeparator.Text = _settings.Separator;
            
            chkExportPdf.IsChecked = _settings.ExportPdf;
            chkExportDwg.IsChecked = _settings.ExportDwg;
            chkCombinePdf.IsChecked = _settings.CombinePdf;
            
            cmbColorMode.SelectedIndex = _settings.ColorMode;
            cmbRasterQuality.SelectedIndex = _settings.RasterQuality;

            // Reconstruct Naming Rule
            NamingRule.Clear();
            foreach (string s in _settings.NamingRule)
            {
                if (s.StartsWith("[Datum:")) NamingRule.Add(new NamingPart { IsDate = true, Name = "Date" });
                else if (s.StartsWith("[") && s.EndsWith("]")) NamingRule.Add(new NamingPart { IsParameter = true, Name = s.Trim('[', ']') });
                else NamingRule.Add(new NamingPart { IsParameter = false, Name = s });
            }
            listNamingRule.ItemsSource = NamingRule;
        }

        private void LoadPrintSets()
        {
            _printSets = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheetSet))
                .Cast<ViewSheetSet>()
                .OrderBy(s => s.Name)
                .ToList();

            foreach (var set in _printSets)
            {
                cmbPrintSets.Items.Add(new ComboBoxItem { Content = set.Name, Tag = set });
            }
        }

        private void LoadDwgSetups()
        {
            _dwgSetups = new FilteredElementCollector(_doc)
                .OfClass(typeof(ExportDWGSettings))
                .Cast<ExportDWGSettings>()
                .OrderBy(s => s.Name)
                .ToList();

            cmbDwgSetup.ItemsSource = _dwgSetups;
            
            // Select last used or default
            if (!string.IsNullOrEmpty(_settings.LastDwgSetupName))
            {
                var match = _dwgSetups.FirstOrDefault(s => s.Name == _settings.LastDwgSetupName);
                if (match != null) cmbDwgSetup.SelectedItem = match;
                else if (_dwgSetups.Any()) cmbDwgSetup.SelectedIndex = 0;
            }
            else if (_dwgSetups.Any())
            {
                cmbDwgSetup.SelectedIndex = 0;
            }
        }

        private bool SheetFilter(object obj)
        {
            if (obj is SheetItem item)
            {
                // Print Set Filter
                if (cmbPrintSets.SelectedItem is ComboBoxItem cbi && cbi.Tag is ViewSheetSet vss)
                {
                    if (!vss.Views.Contains(item.Sheet)) return false;
                }

                // Search Filter
                string search = txtSearch.Text;
                if (string.IsNullOrEmpty(search) || search == "Zoeken op nummer of naam...") return true;
                
                return item.SheetNumber.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       item.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            return false;
        }

        private void CmbPrintSets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _sheetsView?.Refresh();
            // Auto-select visible sheets? Maybe optional. For now just filter list.
        }

        private void LoadParameters()
        {
            var sample = new FilteredElementCollector(_doc).OfClass(typeof(ViewSheet)).FirstOrDefault() as ViewSheet;
            if (sample == null) return;

            var paramList = new List<ParameterItem>();
            foreach (Parameter p in sample.Parameters)
            {
                if (p.Definition == null) continue;
                if (string.IsNullOrEmpty(p.Definition.Name)) continue;
                paramList.Add(new ParameterItem { Name = p.Definition.Name, Id = p.Id });
            }
            
            AvailableParameters = paramList
                .GroupBy(x => x.Name)
                .Select(g => g.First())
                .OrderBy(x => x.Name)
                .ToList();
                
            listAvailableParams.ItemsSource = AvailableParameters;
        }

        private void LoadSheets()
        {
            var sheetList = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .OrderBy(s => s.SheetNumber)
                .ToList();

            foreach (var s in sheetList)
            {
                var item = new SheetItem { Sheet = s };
                DetectSheetSize(item);
                Sheets.Add(item);
            }
            dgSheets.ItemsSource = Sheets;
        }

        private void DetectSheetSize(SheetItem item)
        {
            var tb = new FilteredElementCollector(_doc, item.Sheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType()
                .FirstOrDefault();

            if (tb != null)
            {
                var bb = tb.get_BoundingBox(item.Sheet);
                if (bb != null)
                {
                    double wFeet = bb.Max.X - bb.Min.X;
                    double hFeet = bb.Max.Y - bb.Min.Y;
                    
                    double wMM = Math.Round(wFeet * 304.8);
                    double hMM = Math.Round(hFeet * 304.8);

                    item.WidthMM = wMM;
                    item.HeightMM = hMM;

                    double max = Math.Max(wMM, hMM);
                    double min = Math.Min(wMM, hMM);

                    item.Orientation = (wMM > hMM) ? "Land." : "Port.";

                    if (MatchSize(max, min, 1189, 841)) item.DetectedSize = "A0";
                    else if (MatchSize(max, min, 841, 594)) item.DetectedSize = "A1";
                    else if (MatchSize(max, min, 594, 420)) item.DetectedSize = "A2";
                    else if (MatchSize(max, min, 420, 297)) item.DetectedSize = "A3";
                    else if (MatchSize(max, min, 297, 210)) item.DetectedSize = "A4";
                    else item.DetectedSize = $"{max}x{min}";
                }
            }
            else
            {
                item.DetectedSize = "N/A";
                item.Orientation = "N/A";
            }
        }

        private bool MatchSize(double w, double h, double targetW, double targetH)
        {
            const double tol = 20.0;
            return Math.Abs(w - targetW) < tol && Math.Abs(h - targetH) < tol;
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Selecteer de map voor de export";
                dialog.SelectedPath = _targetFolder;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _targetFolder = dialog.SelectedPath;
                    txtFolderPath.Text = _targetFolder;
                }
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var selected = Sheets.Where(s => s.IsSelected).ToList();
            
            // Check visibility - filter out non-visible (filtered) sheets if needed?
            // Usually user expects only visible filtered sheets to be exported if they did "Select All"
            // But IsSelected property persists even if filtered out.
            // Let's assume IsSelected is the source of truth, but we should probably respect the filter if logic demands.
            // For now, stick to IsSelected property.

            if (!selected.Any())
            {
                MessageBox.Show("Selecteer minimaal één sheet.");
                return;
            }

            if (!Directory.Exists(_targetFolder))
            {
                MessageBox.Show("De geselecteerde map bestaat niet.");
                return;
            }

             SaveSettings();

            bool doPdf = chkExportPdf.IsChecked == true;
            bool doDwg = chkExportDwg.IsChecked == true;
            bool doCombine = chkCombinePdf.IsChecked == true;

            if (!doPdf && !doDwg) return;

            panelProgress.Visibility = System.Windows.Visibility.Visible;
            btnExport.IsEnabled = false;
            pbProgress.Maximum = selected.Count + (doCombine && doPdf ? 1 : 0);
            pbProgress.Value = 0;

            try
            {
                int pCount = 0;
                
                if (doPdf && doCombine)
                {
                    txtProgress.Text = "Bezig met samenvoegen tot één PDF...";
                    DoEvents();
                    ExportSheetsToCombinedPdf(OrderedSheets.Any() ? OrderedSheets.ToList() : selected); // Use ordered if populated
                    pCount++;
                    pbProgress.Value = pCount;
                }

                foreach (var item in selected)
                {
                    txtProgress.Text = $"Exporteren: {item.SheetNumber} - {item.Name}";
                    DoEvents();

                    if (doPdf && !doCombine) ExportSheetToPdf(item);
                    if (doDwg) ExportSheetToDwg(item);

                    pCount++;
                    pbProgress.Value = pCount;
                }

                if (chkOpenFolder.IsChecked == true)
                {
                    Process.Start("explorer.exe", _targetFolder);
                }

                MessageBox.Show("Export succesvol afgerond!");
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fout bij export: {ex.Message}\n\nStack: {ex.StackTrace}");
            }
            finally
            {
                btnExport.IsEnabled = true;
                panelProgress.Visibility = System.Windows.Visibility.Collapsed;
            }
        }

        private void ExportSheetToPdf(SheetItem item)
        {
            PDFExportOptions options = new PDFExportOptions();
            options.FileName = GetConfiguredFileName(item);
            
            if (item.DetectedSize == "A0") options.PaperFormat = ExportPaperFormat.ISO_A0;
            else if (item.DetectedSize == "A1") options.PaperFormat = ExportPaperFormat.ISO_A1;
            else if (item.DetectedSize == "A2") options.PaperFormat = ExportPaperFormat.ISO_A2;
            else if (item.DetectedSize == "A3") options.PaperFormat = ExportPaperFormat.ISO_A3;
            else if (item.DetectedSize == "A4") options.PaperFormat = ExportPaperFormat.ISO_A4;
            else options.PaperFormat = ExportPaperFormat.Default;

            options.HideUnreferencedViewTags = chkHideUnreferencedTags.IsChecked == true;
            options.MaskCoincidentLines = chkMaskCoincidentLines.IsChecked == true;
            options.StopOnError = false;

            if (cmbColorMode.SelectedIndex == 1) options.ColorDepth = ColorDepthType.BlackLine;
            else if (cmbColorMode.SelectedIndex == 2) options.ColorDepth = ColorDepthType.GrayScale;
            else options.ColorDepth = ColorDepthType.Color;

            if (cmbRasterQuality.SelectedIndex == 0) options.RasterQuality = RasterQualityType.Low;
            else if (cmbRasterQuality.SelectedIndex == 1) options.RasterQuality = RasterQualityType.Medium;
            else if (cmbRasterQuality.SelectedIndex == 2) options.RasterQuality = RasterQualityType.High;
            else options.RasterQuality = RasterQualityType.Presentation;

            options.AlwaysUseRaster = rbRaster.IsChecked == true;

            if (rbCenter.IsChecked == true) options.PaperPlacement = PaperPlacementType.Center;
            else options.PaperPlacement = PaperPlacementType.LowerLeft;

            if (rbZoomFit.IsChecked == true) options.ZoomType = ZoomType.FitToPage;
            else
            {
                options.ZoomType = ZoomType.Zoom;
                if (int.TryParse(txtZoomPercent.Text, out int zoom)) options.ZoomPercentage = zoom;
                else options.ZoomPercentage = 100;
            }

            _doc.Export(_targetFolder, new List<ElementId> { item.Sheet.Id }, options);
        }

        private void ExportSheetsToCombinedPdf(List<SheetItem> items)
        {
            PDFExportOptions options = new PDFExportOptions();
            options.Combine = true;
            options.FileName = GetConfiguredFileName(items.First());
            options.PaperFormat = ExportPaperFormat.Default;
            options.PaperOrientation = PageOrientationType.Auto;
            options.HideUnreferencedViewTags = chkHideUnreferencedTags.IsChecked == true;
            options.MaskCoincidentLines = chkMaskCoincidentLines.IsChecked == true;
            options.StopOnError = false;

            if (cmbColorMode.SelectedIndex == 1) options.ColorDepth = ColorDepthType.BlackLine;
            else if (cmbColorMode.SelectedIndex == 2) options.ColorDepth = ColorDepthType.GrayScale;
            else options.ColorDepth = ColorDepthType.Color;

            if (cmbRasterQuality.SelectedIndex == 0) options.RasterQuality = RasterQualityType.Low;
            else if (cmbRasterQuality.SelectedIndex == 1) options.RasterQuality = RasterQualityType.Medium;
            else if (cmbRasterQuality.SelectedIndex == 2) options.RasterQuality = RasterQualityType.High;
            else options.RasterQuality = RasterQualityType.Presentation;

            options.AlwaysUseRaster = rbRaster.IsChecked == true;
            if (rbCenter.IsChecked == true) options.PaperPlacement = PaperPlacementType.Center;
            else options.PaperPlacement = PaperPlacementType.LowerLeft;

            if (rbZoomFit.IsChecked == true) options.ZoomType = ZoomType.FitToPage;
            else
            {
                options.ZoomType = ZoomType.Zoom;
                if (int.TryParse(txtZoomPercent.Text, out int zoom)) options.ZoomPercentage = zoom;
                else options.ZoomPercentage = 100;
            }

            var ids = items.Select(i => i.Sheet.Id).ToList();
            _doc.Export(_targetFolder, ids, options);
        }

        private void ExportSheetToDwg(SheetItem item)
        {
            DWGExportOptions options = null;
            
            // Use selected setup
            if (cmbDwgSetup.SelectedItem is ExportDWGSettings setup)
            {
                options = setup.GetDWGExportOptions();
            }
            
            // Fallback
            if (options == null) options = DWGExportOptions.GetPredefinedOptions(_doc, "");
            if (options == null) options = new DWGExportOptions();

            _doc.Export(_targetFolder, GetConfiguredFileName(item), new List<ElementId> { item.Sheet.Id }, options);
        }

        private void DoEvents()
        {
            System.Windows.Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, new Action(delegate { }));
        }

        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void BtnTab_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button btn = sender as System.Windows.Controls.Button;
            
            gridTabSelection.Visibility = System.Windows.Visibility.Collapsed;
            gridTabSettings.Visibility = System.Windows.Visibility.Collapsed;
            gridTabNaming.Visibility = System.Windows.Visibility.Collapsed;
            gridTabOrder.Visibility = System.Windows.Visibility.Collapsed;
            gridTabDwgSettings.Visibility = System.Windows.Visibility.Collapsed;
            
            btnTabSelection.Opacity = 0.5; btnTabSelection.FontWeight = FontWeights.Normal;
            btnTabSettings.Opacity = 0.5; btnTabSettings.FontWeight = FontWeights.Normal;
            btnTabNaming.Opacity = 0.5; btnTabNaming.FontWeight = FontWeights.Normal;
            btnTabOrder.Opacity = 0.5; btnTabOrder.FontWeight = FontWeights.Normal;
            btnTabDwgSettings.Opacity = 0.5; btnTabDwgSettings.FontWeight = FontWeights.Normal;

            if (btn == btnTabSelection)
            {
                gridTabSelection.Visibility = System.Windows.Visibility.Visible;
                btnTabSelection.Opacity = 1.0; btnTabSelection.FontWeight = FontWeights.Bold;
            }
            else if (btn == btnTabSettings)
            {
                gridTabSettings.Visibility = System.Windows.Visibility.Visible;
                btnTabSettings.Opacity = 1.0; btnTabSettings.FontWeight = FontWeights.Bold;
            }
            else if (btn == btnTabNaming)
            {
                gridTabNaming.Visibility = System.Windows.Visibility.Visible;
                btnTabNaming.Opacity = 1.0; btnTabNaming.FontWeight = FontWeights.Bold;
            }
            else if (btn == btnTabOrder)
            {
                gridTabOrder.Visibility = System.Windows.Visibility.Visible;
                btnTabOrder.Opacity = 1.0; btnTabOrder.FontWeight = FontWeights.Bold;
            }
            else if (btn == btnTabDwgSettings)
            {
                gridTabDwgSettings.Visibility = System.Windows.Visibility.Visible;
                btnTabDwgSettings.Opacity = 1.0; btnTabDwgSettings.FontWeight = FontWeights.Bold;
            }
        }

        private void ChkCombinePdf_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;
            bool combine = chkCombinePdf.IsChecked == true;
            btnTabOrder.Visibility = combine ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            if (combine)
            {
                var selected = Sheets.Where(s => s.IsSelected).ToList();
                var newOrdered = new ObservableCollection<SheetItem>();
                foreach (var item in OrderedSheets) if (selected.Contains(item)) newOrdered.Add(item);
                foreach (var item in selected) if (!newOrdered.Contains(item)) newOrdered.Add(item);
                OrderedSheets = newOrdered;
                listOrderedSheets.ItemsSource = OrderedSheets;
            }
        }

        private void BtnMoveSheetUp_Click(object sender, RoutedEventArgs e)
        {
            int index = listOrderedSheets.SelectedIndex;
            if (index > 0)
            {
                var item = OrderedSheets[index];
                OrderedSheets.RemoveAt(index);
                OrderedSheets.Insert(index - 1, item);
                listOrderedSheets.SelectedIndex = index - 1;
            }
        }

        private void BtnMoveSheetDown_Click(object sender, RoutedEventArgs e)
        {
            int index = listOrderedSheets.SelectedIndex;
            if (index >= 0 && index < OrderedSheets.Count - 1)
            {
                var item = OrderedSheets[index];
                OrderedSheets.RemoveAt(index);
                OrderedSheets.Insert(index + 1, item);
                listOrderedSheets.SelectedIndex = index + 1;
            }
        }

        private string GetConfiguredFileName(SheetItem item)
        {
            List<string> values = new List<string>();
            foreach (var part in NamingRule)
            {
                string partValue = "";
                if (part.IsParameter)
                {
                    Parameter p = item.Sheet.LookupParameter(part.Name);
                    if (p != null) partValue = p.AsValueString() ?? p.AsString() ?? "";
                }
                else if (part.IsDate) partValue = DateTime.Now.ToString("yyMMdd");
                else partValue = part.Name;

                if (!string.IsNullOrEmpty(partValue)) values.Add(partValue);
            }

            string separator = (chkAutoSeparator.IsChecked == true) ? txtAutoSeparator.Text : "";
            string fileName = string.Join(separator, values);

            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars) fileName = fileName.Replace(c, '_');

            return string.IsNullOrWhiteSpace(fileName) ? "Untitled" : fileName;
        }

        private void UpdatePreview()
        {
            var sample = Sheets.FirstOrDefault();
            if (sample == null) return;
            txtNamingPreview.Text = GetConfiguredFileName(sample) + (chkExportPdf.IsChecked == true ? ".pdf" : ".dwg");
        }

        private void BtnAddParam_Click(object sender, RoutedEventArgs e)
        {
            if (listAvailableParams.SelectedItem is ParameterItem pi)
            {
                NamingRule.Add(new NamingPart { IsParameter = true, Name = pi.Name });
                UpdatePreview();
            }
        }

        private void BtnAddDate_Click(object sender, RoutedEventArgs e)
        {
            NamingRule.Add(new NamingPart { IsDate = true, Name = "Date" });
            UpdatePreview();
        }

        private void BtnAddText_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(txtSeparator.Text))
            {
                NamingRule.Add(new NamingPart { IsParameter = false, Name = txtSeparator.Text });
                UpdatePreview();
            }
        }

        private void BtnRemovePart_Click(object sender, RoutedEventArgs e)
        {
            if (listNamingRule.SelectedItem is NamingPart np)
            {
                NamingRule.Remove(np);
                UpdatePreview();
            }
        }

        private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            int index = listNamingRule.SelectedIndex;
            if (index > 0)
            {
                var item = NamingRule[index];
                NamingRule.RemoveAt(index);
                NamingRule.Insert(index - 1, item);
                listNamingRule.SelectedIndex = index - 1;
                UpdatePreview();
            }
        }

        private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            int index = listNamingRule.SelectedIndex;
            if (index >= 0 && index < NamingRule.Count - 1)
            {
                var item = NamingRule[index];
                NamingRule.RemoveAt(index);
                NamingRule.Insert(index + 1, item);
                listNamingRule.SelectedIndex = index + 1;
                UpdatePreview();
            }
        }

        private void ExportFormat_Changed(object sender, RoutedEventArgs e)
        {
            if (IsLoaded) UpdatePreview();
        }

        private void ChkExportDwg_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;
            bool doDwg = chkExportDwg.IsChecked == true;
            btnTabDwgSettings.Visibility = doDwg ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            if (IsLoaded) UpdatePreview();
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e) { _sheetsView?.Refresh(); }

        private void TxtSearch_GotFocus(object sender, RoutedEventArgs e)
        {
            if (txtSearch.Text == "Zoeken op nummer of naam...") { txtSearch.Text = ""; txtSearch.Foreground = System.Windows.Media.Brushes.Black; }
        }

        private void TxtSearch_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtSearch.Text)) { txtSearch.Text = "Zoeken op nummer of naam..."; txtSearch.Foreground = System.Windows.Media.Brushes.Gray; }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e) { foreach (var s in Sheets) s.IsSelected = true; }
        
        private void BtnSelectNone_Click(object sender, RoutedEventArgs e) { foreach (var s in Sheets) s.IsSelected = false; }
        
        private void BtnCancel_Click(object sender, RoutedEventArgs e) => Close();
    }
}
