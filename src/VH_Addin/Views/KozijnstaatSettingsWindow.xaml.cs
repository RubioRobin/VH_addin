// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Venster voor het instellen van de kozijnstaat-configuratie
// ============================================================================

using System;
using System.Windows;
using VH_Tools.Models;
using VH_Tools.Utilities;
using System.Linq;

namespace VH_Addin.Views
{
    public partial class KozijnstaatSettingsWindow : Window
    {
        public KozijnstaatConfig Config { get; private set; }
        public bool DialogAccepted { get; private set; }
        private bool _isInitializing = true;
        private Autodesk.Revit.DB.Document _doc;
        
        // Static fields for persistence of location
        private static double _lastLeft = double.NaN;
        private static double _lastTop = double.NaN;

        public KozijnstaatSettingsWindow(KozijnstaatConfig config, Autodesk.Revit.DB.Document doc)
        {
            InitializeComponent();
            Config = config;
            _doc = doc;
            LoadSettings();

            // Restore location
            if (!double.IsNaN(_lastLeft))
            {
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.Left = _lastLeft;
                this.Top = _lastTop;
            }

            _isInitializing = false;
            UpdateCount();
            
            // Add closing handler to save state
            this.Closing += KozijnstaatSettingsWindow_Closing;
        }

        private void KozijnstaatSettingsWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Save location
            _lastLeft = this.Left;
            _lastTop = this.Top;

            // Save values to config object so they persist for next open
            SaveToConfig();
        }

        private void SaveToConfig()
        {
            if (_isInitializing) return;

            try
            {
                // Get element filter
                var elementFilterItem = cbElementFilter.SelectedItem as System.Windows.Controls.ComboBoxItem;
                string elementFilterStr = elementFilterItem?.Content.ToString() ?? "Kozijnen";
                switch (elementFilterStr)
                {
                    case "Deuren":
                        Config.ElementFilter = ElementTypeFilter.Deuren;
                        break;
                    case "Beide":
                        Config.ElementFilter = ElementTypeFilter.Beide;
                        break;
                    default:
                        Config.ElementFilter = ElementTypeFilter.Kozijnen;
                        break;
                }

                // Get selected format
                var formatItem = cbFormat.SelectedItem as System.Windows.Controls.ComboBoxItem;
                Config.Format = formatItem?.Content.ToString() ?? "A1";

                // Get custom dimensions if Custom format selected
                Config.CustomWidth_MM = ParseDouble(tbCustomWidth.Text, 841.0);
                Config.CustomHeight_MM = ParseDouble(tbCustomHeight.Text, 1189.0);

                // Get selected orientation
                var orientationItem = cbOrientation.SelectedItem as System.Windows.Controls.ComboBoxItem;
                Config.Orientation = orientationItem?.Content.ToString() ?? "Liggend";

                // Get selected view type
                var viewTypeItem = cbViewType.SelectedItem as System.Windows.Controls.ComboBoxItem;
                Config.ViewType = viewTypeItem?.Content.ToString() ?? "Back";

                // Get Assembly Code Filter
                Config.AssemblyCodeFilter = tbAssemblyCode.Text?.Trim() ?? "";

                // Parse numeric values
                Config.MarginBLR_MM = ParseDouble(tbMarginBLR.Text, 75.0);
                Config.MarginBottom_MM = ParseDouble(tbMarginBottom.Text, 75.0);
                Config.RowPitch_MM = ParseDouble(tbRowPitch.Text, 7200.0);
                Config.Gap_MM = ParseDouble(tbGap.Text, 2250.0);
                Config.LineOffset_MM = ParseDouble(tbLineOffset.Text, 250.0);
                Config.MinBackGap_MM = ParseDouble(tbMinBackGap.Text, 500.0);

                // Get boolean values
                Config.OffsetDims = chkOffsetDims.IsChecked ?? true;
                Config.UseHeights = chkUseHeights.IsChecked ?? true;
                Config.CreateReferencePlanes = chkRefPlanes.IsChecked ?? true;
                Config.RefPlaneLeft_MM = ParseDouble(tbRefLeft.Text, -1000.0);
                Config.RefPlaneTop_MM = ParseDouble(tbRefTop.Text, 4000.0);
            }
            catch { }
        }

        private void LoadSettings()
        {
            // Set Element Filter
            SelectComboBoxItem(cbElementFilter, Config.ElementFilter.ToString());
            
            // Set Format
            SelectComboBoxItem(cbFormat, Config.Format);
            
            // Set custom size fields
            if (Config.Format == "Custom")
            {
                tbCustomWidth.Text = Config.CustomWidth_MM.ToString();
                tbCustomHeight.Text = Config.CustomHeight_MM.ToString();
                gridCustomSize.Visibility = System.Windows.Visibility.Visible;
            }
            else
            {
                gridCustomSize.Visibility = System.Windows.Visibility.Collapsed;
            }

            // Set Orientation
            SelectComboBoxItem(cbOrientation, Config.Orientation);

            // Set View Type (Front/Back)
            SelectComboBoxItem(cbViewType, Config.ViewType);

            // Set Assembly Code Filter
            tbAssemblyCode.Text = Config.AssemblyCodeFilter;

            // Set text values
            tbMarginBLR.Text = Config.MarginBLR_MM.ToString();
            tbMarginBottom.Text = Config.MarginBottom_MM.ToString();
            tbRowPitch.Text = Config.RowPitch_MM.ToString();
            tbGap.Text = Config.Gap_MM.ToString();
            tbLineOffset.Text = Config.LineOffset_MM.ToString();
            tbMinBackGap.Text = Config.MinBackGap_MM.ToString();

            // Set checkboxes
            //chkPlaceDimensions.IsChecked = Config.PlaceDimensions;
            chkOffsetDims.IsChecked = Config.OffsetDims;
            chkUseHeights.IsChecked = Config.UseHeights;
            // chkAllTypes removed - always true effectively
            chkRefPlanes.IsChecked = Config.CreateReferencePlanes;
            tbRefLeft.Text = Config.RefPlaneLeft_MM.ToString();
            tbRefTop.Text = Config.RefPlaneTop_MM.ToString();
            gridRefPlaneOffsets.Visibility = (Config.CreateReferencePlanes ? Visibility.Visible : Visibility.Collapsed);
        }

        private void CbFormat_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            // Skip if we're still initializing
            if (_isInitializing || gridCustomSize == null)
                return;

            var formatItem = cbFormat.SelectedItem as System.Windows.Controls.ComboBoxItem;
            string format = formatItem?.Content.ToString() ?? "A1";
            
            if (format == "Custom")
            {
                gridCustomSize.Visibility = System.Windows.Visibility.Visible;
            }
            else
            {
                gridCustomSize.Visibility = System.Windows.Visibility.Collapsed;
            }
        }

        private void SelectComboBoxItem(System.Windows.Controls.ComboBox comboBox, string value)
        {
            for (int i = 0; i < comboBox.Items.Count; i++)
            {
                var item = comboBox.Items[i] as System.Windows.Controls.ComboBoxItem;
                if (item != null && item.Content.ToString() == value)
                {
                    comboBox.SelectedIndex = i;
                    return;
                }
            }
        }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveToConfig();
                
                DialogAccepted = true;
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fout bij het verwerken van de instellingen:\n{ex.Message}", 
                    "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TitleBar_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogAccepted = false;
            DialogResult = false;
            Close();
        }

        private void ChkRefPlanes_Checked(object sender, RoutedEventArgs e)
        {
            if (_isInitializing) return;
            gridRefPlaneOffsets.Visibility = Visibility.Visible;
        }

        private void ChkRefPlanes_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_isInitializing) return;
            gridRefPlaneOffsets.Visibility = Visibility.Collapsed;
        }

        private void UpdateCount()
        {
            if (_isInitializing || _doc == null) return;

            try
            {
                // Determine filters
                var elementFilterItem = cbElementFilter.SelectedItem as System.Windows.Controls.ComboBoxItem;
                string elementFilterStr = elementFilterItem?.Content.ToString() ?? "Kozijnen";
                
                bool includeWindows = elementFilterStr == "Kozijnen" || elementFilterStr == "Beide";
                bool includeDoors = elementFilterStr == "Deuren" || elementFilterStr == "Beide";

                // Get all symbols
                var allSyms = WindowHelpers.GetElementSymbols(_doc, includeWindows, includeDoors);
                
                // Filter by placed
                var placedIds = WindowHelpers.GetPlacedElementTypeIds(_doc, includeWindows, includeDoors);
                var placedSyms = allSyms.Where(s => placedIds.Contains(s.Id.IntegerValue)).ToList();

                // Filter by assembly code (if entered)
                string code = tbAssemblyCode.Text?.Trim() ?? "";
                if (!string.IsNullOrEmpty(code))
                {
                    placedSyms = placedSyms.Where(s => 
                    {
                        var ac = WindowHelpers.GetAssemblyCode(s);
                        return ac.StartsWith(code);
                    }).ToList();
                }

                // Filter by used
                var usedIds = WindowHelpers.GetUsedWindowTypeIds(_doc);
                var available = placedSyms.Where(s => !usedIds.Contains(s.Id.IntegerValue)).ToList();

                tbCountStatus.Text = $"Resterend: {available.Count} types (van {placedSyms.Count} geplaatst)";
            }
            catch (Exception ex)
            {
                tbCountStatus.Text = "Status: -";
            }
        }

        private void CbElementFilter_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!_isInitializing) UpdateCount();
        }

        private void TbAssemblyCode_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
             if (!_isInitializing) UpdateCount();
        }

        private double ParseDouble(string text, double defaultValue)
        {
            text = text?.Replace(",", ".").Trim();
            if (string.IsNullOrEmpty(text))
                return defaultValue;
            
            if (double.TryParse(text, System.Globalization.NumberStyles.Float, 
                System.Globalization.CultureInfo.InvariantCulture, out double result))
                return result;
            
            return defaultValue;
        }

        private int ParseInt(string text, int defaultValue)
        {
            text = text?.Trim();
            if (string.IsNullOrEmpty(text))
                return defaultValue;
            
            if (int.TryParse(text, out int result))
                return result;
            
            return defaultValue;
        }
    }
}