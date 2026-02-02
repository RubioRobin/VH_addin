using System;
using System.Windows;

namespace VH_Addin.Views
{
    public partial class KozijnGeneratorWindow : Window
    {
        public KozijnGeneratorWindow()
        {
            InitializeComponent();
        }

        private void TitleBar_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
