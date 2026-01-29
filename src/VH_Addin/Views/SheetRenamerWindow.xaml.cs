using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using Grid = System.Windows.Controls.Grid;
using ComboBox = System.Windows.Controls.ComboBox;
using TextBox = System.Windows.Controls.TextBox;
using Color = System.Windows.Media.Color;

namespace VH_Addin.Views
{
    public partial class SheetRenamerWindow : Window
    {
        private Document _doc;
        private List<SheetItem> _allSheets = new List<SheetItem>();
        private ObservableCollection<SheetItem> _filteredSheets = new ObservableCollection<SheetItem>();
        private TextBox _txtSearch;
        private ListBox _lbSheets;
        private TextBox _txtNewName;

        private static double _lastWindowLeft = double.NaN;
        private static double _lastWindowTop = double.NaN;

        public class SheetItem
        {
            public ViewSheet Sheet { get; set; }
            public string DisplayName => $"{Sheet.SheetNumber} - {Sheet.Name}";
            public bool IsSelected { get; set; }
        }

        public SheetRenamerWindow(Document doc)
        {
            _doc = doc;
            InitializeComponent();
            
            // Window style
            Width = 420;
            Height = 650;
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            if (!double.IsNaN(_lastWindowLeft) && !double.IsNaN(_lastWindowTop))
            {
                WindowStartupLocation = WindowStartupLocation.Manual;
                Left = _lastWindowLeft;
                Top = _lastWindowTop;
            }
            else
            {
                WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }

            this.MouseDown += (s, e) => { if (e.ChangedButton == MouseButton.Left) DragMove(); };
            this.Closing += (s, e) => { _lastWindowLeft = Left; _lastWindowTop = Top; };
            this.KeyDown += (s, e) => { if (e.Key == Key.Escape) Close(); };

            LoadSheets();
            BuildUI();
        }

        private void LoadSheets()
        {
            var sheets = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .OrderBy(s => s.SheetNumber)
                .ToList();

            foreach (var s in sheets)
            {
                _allSheets.Add(new SheetItem { Sheet = s, IsSelected = false });
            }
            
            UpdateFilter("");
        }

        private void BuildUI()
        {
            var mainGrid = new Grid { HorizontalAlignment = HorizontalAlignment.Center };
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var islandBorder = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(24, 20, 24, 20),
                Margin = new Thickness(10),
                MinWidth = 380,
                MaxWidth = 380,
                Effect = new DropShadowEffect { Color = Colors.Black, Direction = 270, ShadowDepth = 4, BlurRadius = 15, Opacity = 0.15 }
            };

            var contentStack = new StackPanel();

            // TITLE (no X button)
            contentStack.Children.Add(new TextBlock
            {
                Text = "Sheet Renamer",
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(30, 41, 59)),
                Margin = new Thickness(0, 0, 0, 16)
            });

            // SEARCH
            contentStack.Children.Add(CreateLabel("Zoek op nummer of naam:"));
            _txtSearch = CreateTextBox("Zoeken...");
            _txtSearch.TextChanged += (s, e) => UpdateFilter(_txtSearch.Text);
            contentStack.Children.Add(_txtSearch);

            // LIST
            contentStack.Children.Add(CreateLabel("Selecteer sheets om te hernoemen:"));
            _lbSheets = new ListBox
            {
                Height = 300,
                Margin = new Thickness(0, 0, 0, 16),
                ItemsSource = _filteredSheets,
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                BorderThickness = new Thickness(1)
            };

            // DataTemplate for ListBox items
            var template = new DataTemplate(typeof(SheetItem));
            var factory = new FrameworkElementFactory(typeof(StackPanel));
            factory.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);
            factory.SetValue(StackPanel.MarginProperty, new Thickness(5));

            var checkBox = new FrameworkElementFactory(typeof(CheckBox));
            checkBox.SetBinding(CheckBox.IsCheckedProperty, new System.Windows.Data.Binding("IsSelected") { Mode = BindingMode.TwoWay });
            checkBox.SetValue(CheckBox.VerticalAlignmentProperty, VerticalAlignment.Center);
            checkBox.SetValue(CheckBox.MarginProperty, new Thickness(0, 0, 8, 0));
            factory.AppendChild(checkBox);

            var textBlock = new FrameworkElementFactory(typeof(TextBlock));
            textBlock.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("DisplayName"));
            textBlock.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            factory.AppendChild(textBlock);

            template.VisualTree = factory;
            _lbSheets.ItemTemplate = template;
            contentStack.Children.Add(_lbSheets);

            // NEW NAME
            contentStack.Children.Add(CreateLabel("Nieuwe sheetnaam:"));
            _txtNewName = CreateTextBox("Nieuwe naam...");
            contentStack.Children.Add(_txtNewName);

            islandBorder.Child = contentStack;
            Grid.SetRow(islandBorder, 0);
            mainGrid.Children.Add(islandBorder);

            // BUTTONS
            var btnGrid = new Grid { Margin = new Thickness(10, 0, 10, 20), Width = 380 };
            btnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            btnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            btnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(8) });
            btnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var btnCancel = CreateModernButton("Annuleren", false);
            btnCancel.Click += (s, e) => Close();
            Grid.SetColumn(btnCancel, 1);
            btnGrid.Children.Add(btnCancel);

            var btnOk = CreateModernButton("Hernoem Selectie", true);
            btnOk.Click += OnRename;
            Grid.SetColumn(btnOk, 3);
            btnGrid.Children.Add(btnOk);

            Grid.SetRow(btnGrid, 1);
            mainGrid.Children.Add(btnGrid);

            this.Content = mainGrid;
        }

        private void UpdateFilter(string text)
        {
            _filteredSheets.Clear();
            string lower = text.ToLower();
            foreach (var item in _allSheets)
            {
                if (string.IsNullOrEmpty(lower) || item.DisplayName.ToLower().Contains(lower))
                {
                    _filteredSheets.Add(item);
                }
            }
        }

        private TextBlock CreateLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
                Margin = new Thickness(0, 12, 0, 6)
            };
        }

        private TextBox CreateTextBox(string placeholder)
        {
            var tb = new TextBox
            {
                Background = new SolidColorBrush(Color.FromRgb(249, 250, 251)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10, 8, 10, 8),
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                VerticalContentAlignment = VerticalAlignment.Center
            };

            // Simple placeholder logic
            tb.Text = placeholder;
            tb.Foreground = Brushes.Gray;
            tb.GotFocus += (s, e) => { if (tb.Text == placeholder) { tb.Text = ""; tb.Foreground = new SolidColorBrush(Color.FromRgb(51, 65, 85)); } };
            tb.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(tb.Text)) { tb.Text = placeholder; tb.Foreground = Brushes.Gray; } };

            return tb;
        }

        private Button CreateModernButton(string text, bool isPrimary)
        {
            var btn = new Button
            {
                Content = text,
                Padding = new Thickness(24, 12, 24, 12),
                MinWidth = 140,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Cursor = Cursors.Hand,
                Background = isPrimary ? new SolidColorBrush(Color.FromRgb(178, 153, 107)) : Brushes.White,
                Foreground = isPrimary ? Brushes.White : Brushes.Black,
                BorderBrush = Brushes.Black,
                BorderThickness = new Thickness(1),
                Template = GetButtonTemplate(isPrimary)
            };
            return btn;
        }

        private ControlTemplate GetButtonTemplate(bool isPrimary)
        {
            var template = new ControlTemplate(typeof(Button));
            var borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.Name = "border";
            borderFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(8));
            borderFactory.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
            borderFactory.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
            borderFactory.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Button.PaddingProperty));

            var contentFactory = new FrameworkElementFactory(typeof(ContentPresenter));
            contentFactory.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentFactory.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            borderFactory.AppendChild(contentFactory);

            var hoverTrigger = new Trigger { Property = Button.IsMouseOverProperty, Value = true };
            hoverTrigger.Setters.Add(new Setter(Button.BackgroundProperty, isPrimary ? new SolidColorBrush(Color.FromRgb(160, 138, 97)) : new SolidColorBrush(Color.FromRgb(248, 248, 248))));
            
            template.VisualTree = borderFactory;
            template.Triggers.Add(hoverTrigger);
            return template;
        }

        private void OnRename(object sender, RoutedEventArgs e)
        {
            string newName = _txtNewName.Text;
            if (string.IsNullOrWhiteSpace(newName) || newName == "Nieuwe naam...")
            {
                TaskDialog.Show("Fout", "Vul een geldige nieuwe naam in.");
                return;
            }

            var selectedItems = _allSheets.Where(x => x.IsSelected).ToList();
            if (selectedItems.Count == 0)
            {
                TaskDialog.Show("Fout", "Selecteer eerst sheets in de lijst.");
                return;
            }

            try
            {
                using (Transaction t = new Transaction(_doc, "VH Sheet Renamer"))
                {
                    t.Start();
                    int success = 0;
                    foreach (var item in selectedItems)
                    {
                        try
                        {
                            item.Sheet.Name = newName;
                            success++;
                        }
                        catch { }
                    }
                    t.Commit();
                    TaskDialog.Show("Succes", $"{success} sheets succesvol hernoemd.");
                    
                    // Close the window after completion as requested
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Fout", ex.Message);
            }
        }
    }
}
