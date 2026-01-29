using System.Windows;

namespace VH_DaglichtPlugin
{
    public partial class DaglichtWindow : Window
    {
        public DaglichtWindow()
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
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
                DoAlpha = cbDoAlpha.IsChecked == true,
                DoBeta = cbDoBeta.IsChecked == true,
                DoGlass = cbDoGlass.IsChecked == true,
                OnlyVG = cbOnlyVG.IsChecked == true,
                DoExport = cbDoExport.IsChecked == true
            };

            // Parse sash width
            if (double.TryParse(tbSashWidth.Text, out double val))
            {
                opts.DefaultSashWidthMm = val;
            }
            else
            {
                opts.DefaultSashWidthMm = 69.0; // Fallback
            }

            return opts;
        }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }

    public class DaglichtOptions
    {
        public bool ActiveViewOnly { get; set; }
        public bool CurrentLevelOnly { get; set; }
        public bool SelectionOnly { get; set; }
        public bool AllModelOnly { get; set; } // Kept for completeness
        
        public bool DrawLines { get; set; }
        public bool DoAlpha { get; set; }
        public bool DoBeta { get; set; }
        public bool DoGlass { get; set; }
        public bool OnlyVG { get; set; }
        public bool DoExport { get; set; }

        public double DefaultSashWidthMm { get; set; }
    }
}
