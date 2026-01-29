using System.Windows;
using System.Windows.Input;
using VH_Tools.Models;

namespace VH_Tools.Views
{
    public partial class DaglichtWindow : Window
    {
        // Persistence
        private static double _lastLeft = double.NaN;
        private static double _lastTop = double.NaN;

        private static bool _pActiveView = true;
        private static bool _pCurrentLevel = false;
        private static bool _pSelection = false;
        private static bool _pAllModel = false;

        private static bool _pDrawLines = true;
        private static bool _pFlipAlphaFan = false;
        private static bool _pDoAlpha = true;
        private static bool _pDoBeta = true;
        private static bool _pDoGlass = true;
        private static bool _pOnlyVG = false;
        private static bool _pDoExport = true;

        private static string _pSashWidth = "69";
        private static string _pFrameHeight = "67";
        private static string _pGlassOffset = "11";

        public DaglichtWindow()
        {
            InitializeComponent();
            RestoreSettings();
        }

        private void RestoreSettings()
        {
            // Position
            if (!double.IsNaN(_lastLeft))
            {
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.Left = _lastLeft;
                this.Top = _lastTop;
            }

            // RadioButtons
            rbActiveView.IsChecked = _pActiveView;
            rbCurrentLevel.IsChecked = _pCurrentLevel;
            rbSelection.IsChecked = _pSelection;
            rbAllModel.IsChecked = _pAllModel;

            // CheckBoxes
            cbDrawLines.IsChecked = _pDrawLines;
            cbFlipAlphaFan.IsChecked = _pFlipAlphaFan;
            cbDoAlpha.IsChecked = _pDoAlpha;
            cbDoBeta.IsChecked = _pDoBeta;
            cbDoGlass.IsChecked = _pDoGlass;
            cbOnlyVG.IsChecked = _pOnlyVG;
            cbDoExport.IsChecked = _pDoExport;

            // TextBoxes
            tbSashWidth.Text = _pSashWidth;
            tbFrameHeight.Text = _pFrameHeight;
            tbGlassOffset.Text = _pGlassOffset;
        }

        private void SaveSettings()
        {
            _pActiveView = rbActiveView.IsChecked == true;
            _pCurrentLevel = rbCurrentLevel.IsChecked == true;
            _pSelection = rbSelection.IsChecked == true;
            _pAllModel = rbAllModel.IsChecked == true;

            _pDrawLines = cbDrawLines.IsChecked == true;
            _pFlipAlphaFan = cbFlipAlphaFan.IsChecked == true;
            _pDoAlpha = cbDoAlpha.IsChecked == true;
            _pDoBeta = cbDoBeta.IsChecked == true;
            _pDoGlass = cbDoGlass.IsChecked == true;
            _pOnlyVG = cbOnlyVG.IsChecked == true;
            _pDoExport = cbDoExport.IsChecked == true;

            _pSashWidth = tbSashWidth.Text;
            _pFrameHeight = tbFrameHeight.Text;
            _pGlassOffset = tbGlassOffset.Text;
        }

        public DaglichtOptions GetOptions()
        {
            var opts = new DaglichtOptions
            {
                ActiveViewOnly = rbActiveView.IsChecked == true,
                CurrentLevelOnly = rbCurrentLevel.IsChecked == true,
                SelectionOnly = rbSelection.IsChecked == true,
                AllModelOnly = rbAllModel.IsChecked == true,
                
                DrawLines = cbDrawLines.IsChecked == true,
                FlipAlphaFan = cbFlipAlphaFan.IsChecked == true,
                DoAlpha = cbDoAlpha.IsChecked == true,
                DoBeta = cbDoBeta.IsChecked == true,
                DoGlass = cbDoGlass.IsChecked == true,
                OnlyVG = cbOnlyVG.IsChecked == true,
                DoExport = cbDoExport.IsChecked == true
            };

            opts.DefaultSashWidthMm = ParseDouble(tbSashWidth.Text, 69.0);
            opts.DefaultFrameHeightMm = ParseDouble(tbFrameHeight.Text, 67.0);
            opts.DefaultGlassOffsetMm = ParseDouble(tbGlassOffset.Text, 11.0);

            return opts;
        }

        private double ParseDouble(string text, double defaultValue)
        {
            if (string.IsNullOrWhiteSpace(text)) return defaultValue;
            string clean = text.Replace(",", ".");
            if (double.TryParse(clean, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val))
            {
                return val;
            }
            return defaultValue;
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            SaveSettings();
            DialogResult = true;
            Close();
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveSettings();
            _lastLeft = this.Left;
            _lastTop = this.Top;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
