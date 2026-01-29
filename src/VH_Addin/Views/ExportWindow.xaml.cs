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

namespace VH_Addin.Views
{
    public partial class ExportWindow : Window, INotifyPropertyChanged
    {
        private readonly Document _doc;
        private string _targetFolder;
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
            
            LoadSheets();
            
            // Restore last used folder from a simple text file
            try
            {
                string settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VH_Engineering", "ExportFolder.txt");
                if (File.Exists(settingsPath))
                    _targetFolder = File.ReadAllText(settingsPath);
            }
            catch { }

            if (string.IsNullOrEmpty(_targetFolder) || !Directory.Exists(_targetFolder))
                _targetFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            
            txtFolderPath.Text = _targetFolder;

            LoadParameters();
            
            // Default naming rule: [Sheet Number]_[Sheet Name]
            NamingRule.Add(new NamingPart { IsParameter = true, Name = "Sheet Number" });
            NamingRule.Add(new NamingPart { IsParameter = false, Name = "_" });
            NamingRule.Add(new NamingPart { IsParameter = true, Name = "Sheet Name" });
            listNamingRule.ItemsSource = NamingRule;
            
            UpdatePreview();

            _sheetsView = System.Windows.Data.CollectionViewSource.GetDefaultView(Sheets);
            _sheetsView.Filter = SheetFilter;
        }

        private bool SheetFilter(object obj)
        {
            if (obj is SheetItem item)
            {
                string search = txtSearch.Text;
                if (string.IsNullOrEmpty(search) || search == "Zoeken op nummer of naam...") return true;
                
                return item.SheetNumber.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       item.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            return false;
        }

        private void LoadParameters()
        {
            // Get a sample sheet to see available parameters
            var sample = new FilteredElementCollector(_doc).OfClass(typeof(ViewSheet)).FirstOrDefault() as ViewSheet;
            if (sample == null) return;

            var paramList = new List<ParameterItem>();
            foreach (Parameter p in sample.Parameters)
            {
                if (p.Definition == null) continue;
                if (string.IsNullOrEmpty(p.Definition.Name)) continue;
                paramList.Add(new ParameterItem { Name = p.Definition.Name, Id = p.Id });
            }
            
            // Filter out duplicates by name
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
            // Find TitleBlock on this sheet
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

                    // Swap if needed to get larger dimension as width for orientation check
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
            const double tol = 20.0; // mm tolerance
            return Math.Abs(w - targetW) < tol && Math.Abs(h - targetH) < tol;
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Selecteer de map voor de export";
                dialog.UseDescriptionForTitle = true;
                dialog.SelectedPath = _targetFolder;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _targetFolder = dialog.SelectedPath;
                    txtFolderPath.Text = _targetFolder;
                    
                    // Save setting
                    try
                    {
                        string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VH_Engineering");
                        if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                        File.WriteAllText(Path.Combine(folder, "ExportFolder.txt"), _targetFolder);
                    }
                    catch { }
                }
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var selected = Sheets.Where(s => s.IsSelected).ToList();
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
                
                // Handle Combined PDF first if requested
                if (doPdf && doCombine)
                {
                    txtProgress.Text = "Bezig met samenvoegen tot één PDF...";
                    DoEvents();
                    ExportSheetsToCombinedPdf(OrderedSheets.ToList());
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

                MessageBox.Show("Export succesvol afgerond!");
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fout bij export: {ex.Message}");
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
            
            // Map detected size to Revit PaperFormat if possible
            // ISO names are common in Revit
            if (item.DetectedSize == "A0") options.PaperFormat = ExportPaperFormat.ISO_A0;
            else if (item.DetectedSize == "A1") options.PaperFormat = ExportPaperFormat.ISO_A1;
            else if (item.DetectedSize == "A2") options.PaperFormat = ExportPaperFormat.ISO_A2;
            else if (item.DetectedSize == "A3") options.PaperFormat = ExportPaperFormat.ISO_A3;
            else if (item.DetectedSize == "A4") options.PaperFormat = ExportPaperFormat.ISO_A4;
            else options.PaperFormat = ExportPaperFormat.Default;

            options.HideUnreferencedViewTags = chkHideUnreferencedTags.IsChecked == true;
            options.MaskCoincidentLines = chkMaskCoincidentLines.IsChecked == true;
            options.StopOnError = false;

            // Appearance: Colors
            if (cmbColorMode.SelectedIndex == 1) options.ColorDepth = ColorDepthType.BlackLine;
            else if (cmbColorMode.SelectedIndex == 2) options.ColorDepth = ColorDepthType.GrayScale;
            else options.ColorDepth = ColorDepthType.Color;

            // Appearance: Raster Quality
            if (cmbRasterQuality.SelectedIndex == 0) options.RasterQuality = RasterQualityType.Low;
            else if (cmbRasterQuality.SelectedIndex == 1) options.RasterQuality = RasterQualityType.Medium;
            else if (cmbRasterQuality.SelectedIndex == 2) options.RasterQuality = RasterQualityType.High;
            else options.RasterQuality = RasterQualityType.Presentation;

            // Hidden Line Views: Vector vs Raster
            options.AlwaysUseRaster = rbRaster.IsChecked == true;

            // Paper Placement
            if (rbCenter.IsChecked == true)
            {
                options.PaperPlacement = PaperPlacementType.Center;
            }
            else
            {
                options.PaperPlacement = PaperPlacementType.LowerLeft;
                // Note: Revit API PDFExportOptions may have OriginOffset or similar 
                // but it's often read-only or depends on exact version.
                // We'll set the placement and hope for the best for now.
            }

            // Zoom
            if (rbZoomFit.IsChecked == true)
            {
                options.ZoomType = ZoomType.FitToPage;
            }
            else
            {
                options.ZoomType = ZoomType.Zoom;
                if (int.TryParse(txtZoomPercent.Text, out int zoom))
                {
                    options.ZoomPercentage = zoom;
                }
                else
                {
                    options.ZoomPercentage = 100;
                }
            }

            _doc.Export(_targetFolder, new List<ElementId> { item.Sheet.Id }, options);
        }

        private void ExportSheetsToCombinedPdf(List<SheetItem> items)
        {
            PDFExportOptions options = new PDFExportOptions();
            options.Combine = true;
            options.FileName = GetConfiguredFileName(items.First());
            
            // For mixed sizes and orientations
            options.PaperFormat = ExportPaperFormat.Default;
            // Note: In some versions it's PaperOrientation, in 2025 it's PaperOrientation
            // PageOrientationType is the correct enum for 2025
            options.PaperOrientation = PageOrientationType.Auto;
            
            options.HideUnreferencedViewTags = chkHideUnreferencedTags.IsChecked == true;
            options.MaskCoincidentLines = chkMaskCoincidentLines.IsChecked == true;
            options.StopOnError = false;

            // Appearance: Colors
            if (cmbColorMode.SelectedIndex == 1) options.ColorDepth = ColorDepthType.BlackLine;
            else if (cmbColorMode.SelectedIndex == 2) options.ColorDepth = ColorDepthType.GrayScale;
            else options.ColorDepth = ColorDepthType.Color;

            // Appearance: Raster Quality
            if (cmbRasterQuality.SelectedIndex == 0) options.RasterQuality = RasterQualityType.Low;
            else if (cmbRasterQuality.SelectedIndex == 1) options.RasterQuality = RasterQualityType.Medium;
            else if (cmbRasterQuality.SelectedIndex == 2) options.RasterQuality = RasterQualityType.High;
            else options.RasterQuality = RasterQualityType.Presentation;

            // Hidden Line Views: Vector vs Raster
            options.AlwaysUseRaster = rbRaster.IsChecked == true;

            // Paper Placement
            if (rbCenter.IsChecked == true) options.PaperPlacement = PaperPlacementType.Center;
            else options.PaperPlacement = PaperPlacementType.LowerLeft;

            // Zoom
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
            DWGExportOptions options = DWGExportOptions.GetPredefinedOptions(_doc, ""); 
            if (options == null) options = new DWGExportOptions();
            
            // Apply UI settings
            // Version
            if (cmbDwgVersion.SelectedIndex == 1) options.FileVersion = ACADVersion.R2018;
            else if (cmbDwgVersion.SelectedIndex == 2) options.FileVersion = ACADVersion.R2013;
            else if (cmbDwgVersion.SelectedIndex == 3) options.FileVersion = ACADVersion.R2010;
            else options.FileVersion = ACADVersion.Default;

            // Coordinates
            options.SharedCoords = cmbDwgCoordinates.SelectedIndex == 1;

            // Units
            if (cmbDwgUnits.SelectedIndex == 1) options.TargetUnit = ExportUnit.Meter;
            else if (cmbDwgUnits.SelectedIndex == 2) options.TargetUnit = ExportUnit.Centimeter;
            else options.TargetUnit = ExportUnit.Millimeter;

            // Colors
            options.Colors = ExportColorMode.IndexColors;
            
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
            
            // Reset all
            gridTabSelection.Visibility = System.Windows.Visibility.Collapsed;
            gridTabSettings.Visibility = System.Windows.Visibility.Collapsed;
            gridTabNaming.Visibility = System.Windows.Visibility.Collapsed;
            
            btnTabSelection.Opacity = 0.5; btnTabSelection.FontWeight = FontWeights.Normal;
            btnTabSettings.Opacity = 0.5; btnTabSettings.FontWeight = FontWeights.Normal;
            btnTabNaming.Opacity = 0.5; btnTabNaming.FontWeight = FontWeights.Normal;
            btnTabOrder.Opacity = 0.5; btnTabOrder.FontWeight = FontWeights.Normal;
            btnTabDwgSettings.Opacity = 0.5; btnTabDwgSettings.FontWeight = FontWeights.Normal;

            if (btn == btnTabSelection)
            {
                gridTabSelection.Visibility = System.Windows.Visibility.Visible;
                btnTabSelection.Opacity = 1.0;
                btnTabSelection.FontWeight = FontWeights.Bold;
            }
            else if (btn == btnTabSettings)
            {
                gridTabSettings.Visibility = System.Windows.Visibility.Visible;
                btnTabSettings.Opacity = 1.0;
                btnTabSettings.FontWeight = FontWeights.Bold;
            }
            else if (btn == btnTabNaming)
            {
                gridTabNaming.Visibility = System.Windows.Visibility.Visible;
                btnTabNaming.Opacity = 1.0;
                btnTabNaming.FontWeight = FontWeights.Bold;
            }
            else if (btn == btnTabOrder)
            {
                gridTabOrder.Visibility = System.Windows.Visibility.Visible;
                btnTabOrder.Opacity = 1.0;
                btnTabOrder.FontWeight = FontWeights.Bold;
            }
            else if (btn == btnTabDwgSettings)
            {
                gridTabDwgSettings.Visibility = System.Windows.Visibility.Visible;
                btnTabDwgSettings.Opacity = 1.0;
                btnTabDwgSettings.FontWeight = FontWeights.Bold;
            }
        }

        private void ChkCombinePdf_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;
            
            bool combine = chkCombinePdf.IsChecked == true;
            btnTabOrder.Visibility = combine ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            
            if (combine)
            {
                // Sync OrderedSheets with currently selected sheets
                var selected = Sheets.Where(s => s.IsSelected).ToList();
                
                // Keep existing order for already present sheets, add new ones at the end
                var newOrdered = new ObservableCollection<SheetItem>();
                foreach (var item in OrderedSheets)
                {
                    if (selected.Contains(item)) newOrdered.Add(item);
                }
                foreach (var item in selected)
                {
                    if (!newOrdered.Contains(item)) newOrdered.Add(item);
                }
                
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
                    if (p != null)
                    {
                        partValue = p.AsValueString() ?? p.AsString() ?? "";
                    }
                }
                else if (part.IsDate)
                {
                    partValue = DateTime.Now.ToString("yyMMdd");
                }
                else
                {
                    partValue = part.Name;
                }

                if (!string.IsNullOrEmpty(partValue))
                {
                    values.Add(partValue);
                }
            }

            string separator = (chkAutoSeparator.IsChecked == true) ? txtAutoSeparator.Text : "";
            string fileName = string.Join(separator, values);

            // Sanitize filename
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                fileName = fileName.Replace(c, '_');
            }

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

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            _sheetsView?.Refresh();
        }

        private void TxtSearch_GotFocus(object sender, RoutedEventArgs e)
        {
            if (txtSearch.Text == "Zoeken op nummer of naam...")
            {
                txtSearch.Text = "";
                txtSearch.Foreground = System.Windows.Media.Brushes.Black;
            }
        }

        private void TxtSearch_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtSearch.Text))
            {
                txtSearch.Text = "Zoeken op nummer of naam...";
                txtSearch.Foreground = System.Windows.Media.Brushes.Gray;
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e) { foreach (var s in Sheets) s.IsSelected = true; }
        private void BtnSelectNone_Click(object sender, RoutedEventArgs e) { foreach (var s in Sheets) s.IsSelected = false; }
        private void BtnCancel_Click(object sender, RoutedEventArgs e) => Close();
    }
}
