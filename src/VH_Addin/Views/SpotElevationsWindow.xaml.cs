// Auteur: Robin Manintveld
// Datum: 22-01-2026
//
// Venster voor het aanpassen van spot elevation weergave
// ============================================================================

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Grid = System.Windows.Controls.Grid;
using ComboBox = System.Windows.Controls.ComboBox;
using Color = System.Windows.Media.Color;

namespace VH_Addin.Views
{
    public partial class SpotElevationsWindow : Window
    {
        private Document _doc;
        private ComboBox _cmb_linepattern;
        private ComboBox _cmb_color;
        private static Autodesk.Revit.DB.Color _lastSelectedColor = null;
        private static ElementId _lastSelectedPatternId = null;
        private static double _lastWindowLeft = double.NaN;
        private static double _lastWindowTop = double.NaN;

        public ElementId SelectedLinePatternId { get; private set; }
        public Autodesk.Revit.DB.Color SelectedColor { get; private set; }

        public SpotElevationsWindow(Document doc)
        {
            _doc = doc;
            InitializeComponent();
            Width = 400;
            Height = 320; // Increased height for Color row
            MinWidth = 400;
            MaxWidth = 400;
            SizeToContent = SizeToContent.Height;
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            
            // Restore last position or center
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
            BuildUI();
        }



        private void BuildUI()
        {
            var mainGrid = new Grid
            {
                Margin = new Thickness(0),
                MaxWidth = 400,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // MODERN STYLE: Wit eiland met subtiele schaduw
            var islandBorder = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(24, 20, 24, 20),
                Margin = new Thickness(8),
                MinWidth = 360,
                MaxWidth = 360,
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    Color = Colors.Black,
                    Direction = 270,
                    ShadowDepth = 2,
                    BlurRadius = 12,
                    Opacity = 0.1
                }
            };

            var contentStack = new StackPanel();

            // TITLE - Modern met donkere kleur
            var title = new TextBlock
            {
                Text = "Spot Elevations Override",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                Margin = new Thickness(0, 0, 0, 16)
            };
            contentStack.Children.Add(title);

            // FORM GRID
            var fieldGrid = new Grid();
            fieldGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            fieldGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            fieldGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            fieldGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(12) }); // Spacer
            fieldGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // ROW 1: LINE PATTERN
            var labelPattern = new TextBlock
            {
                Text = "Lijn patroon:",
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 16, 0)
            };
            Grid.SetRow(labelPattern, 0);
            Grid.SetColumn(labelPattern, 0);
            fieldGrid.Children.Add(labelPattern);

            _cmb_linepattern = new ComboBox
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 250, 252)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(203, 213, 225)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(12, 8, 12, 8),
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                MinWidth = 200
            };

            var linePatterns = GetLinePatterns();
            foreach (var lp in linePatterns)
            {
                var item = new ComboBoxItem { Content = lp.Name, Tag = lp.Id };
                _cmb_linepattern.Items.Add(item);
            }

            if (_lastSelectedPatternId != null)
            {
                for (int i = 0; i < _cmb_linepattern.Items.Count; i++)
                {
                    var item = (ComboBoxItem)_cmb_linepattern.Items[i];
                    if (((ElementId)item.Tag).Equals(_lastSelectedPatternId))
                    {
                        _cmb_linepattern.SelectedIndex = i;
                        break;
                    }
                }
                if (_cmb_linepattern.SelectedIndex == -1)
                    _cmb_linepattern.SelectedIndex = 0;
            }
            else if (_cmb_linepattern.Items.Count > 0)
                _cmb_linepattern.SelectedIndex = 0;

            Grid.SetRow(_cmb_linepattern, 0);
            Grid.SetColumn(_cmb_linepattern, 1);
            fieldGrid.Children.Add(_cmb_linepattern);

            // ROW 2: COLOR
            var labelColor = new TextBlock
            {
                Text = "Kleur:",
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 16, 0)
            };
            Grid.SetRow(labelColor, 2);
            Grid.SetColumn(labelColor, 0);
            fieldGrid.Children.Add(labelColor);

            _cmb_color = new ComboBox
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 250, 252)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(203, 213, 225)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(12, 8, 12, 8),
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                MinWidth = 200
            };

            var colors = GetVHEngineColors();
            foreach (var c in colors)
            {
                var colorBorder = new Border { Width = 16, Height = 16, Background = new SolidColorBrush(Color.FromRgb(c.Value.Red, c.Value.Green, c.Value.Blue)), Margin = new Thickness(0, 0, 8, 0), BorderBrush = Brushes.Gray, BorderThickness = new Thickness(1) };
                var text = new TextBlock { Text = c.Key, VerticalAlignment = VerticalAlignment.Center };
                var stack = new StackPanel { Orientation = Orientation.Horizontal };
                stack.Children.Add(colorBorder);
                stack.Children.Add(text);
                
                var item = new ComboBoxItem { Content = stack, Tag = c.Value };
                _cmb_color.Items.Add(item);
            }

            if (_lastSelectedColor != null)
            {
                for (int i = 0; i < _cmb_color.Items.Count; i++)
                {
                    var item = (ComboBoxItem)_cmb_color.Items[i];
                    var c = (Autodesk.Revit.DB.Color)item.Tag;
                    if (c.Red == _lastSelectedColor.Red && c.Green == _lastSelectedColor.Green && c.Blue == _lastSelectedColor.Blue)
                    {
                        _cmb_color.SelectedIndex = i;
                        break;
                    }
                }
            }
            if (_cmb_color.SelectedIndex == -1 && _cmb_color.Items.Count > 0)
                _cmb_color.SelectedIndex = 0;

            Grid.SetRow(_cmb_color, 2);
            Grid.SetColumn(_cmb_color, 1);
            fieldGrid.Children.Add(_cmb_color);

            contentStack.Children.Add(fieldGrid);

            islandBorder.Child = contentStack;
            Grid.SetRow(islandBorder, 0);
            mainGrid.Children.Add(islandBorder);

            // BUTTONS - Onder het eiland, rechts uitgelijnd
            var btnGrid = new Grid
            {
                Margin = new Thickness(0, 6, 8, 0),
                HorizontalAlignment = HorizontalAlignment.Right,
                Width = 360
            };
            btnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            btnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            btnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var btnCancel = CreateModernButton("Annuleren", false);
            btnCancel.IsCancel = true;
            btnCancel.Click += (s, e) => { DialogResult = false; Close(); };
            Grid.SetColumn(btnCancel, 1);
            btnGrid.Children.Add(btnCancel);

            var btnOk = CreateModernButton("Toepassen", true);
            btnOk.Margin = new Thickness(8, 0, 0, 0);
            btnOk.Click += OnOk;
            Grid.SetColumn(btnOk, 2);
            btnGrid.Children.Add(btnOk);

            Grid.SetRow(btnGrid, 1);
            mainGrid.Children.Add(btnGrid);

            this.Content = mainGrid;
        }

        private Dictionary<string, Autodesk.Revit.DB.Color> GetVHEngineColors()
        {
            return new Dictionary<string, Autodesk.Revit.DB.Color>
            {
                { "Zwart", new Autodesk.Revit.DB.Color(0, 0, 0) },
                { "Wit", new Autodesk.Revit.DB.Color(255, 255, 255) },
                { "Grijs", new Autodesk.Revit.DB.Color(128, 128, 128) },
                { "Rood", new Autodesk.Revit.DB.Color(255, 0, 0) },
                { "Groen", new Autodesk.Revit.DB.Color(0, 128, 0) },
                { "Blauw", new Autodesk.Revit.DB.Color(0, 0, 255) },
                { "Cyaan", new Autodesk.Revit.DB.Color(0, 255, 255) },
                { "Magenta", new Autodesk.Revit.DB.Color(255, 0, 255) },
                { "VH Goud", new Autodesk.Revit.DB.Color(178, 153, 107) },
                { "Oranje", new Autodesk.Revit.DB.Color(255, 165, 0) }
            };
        }

        private Button CreateModernButton(string text, bool isPrimary)
        {
            var btn = new Button
            {
                Content = text,
                Padding = new Thickness(20, 10, 20, 10),
                MinWidth = 110,
                FontSize = 13,
                FontWeight = FontWeights.Medium,
                Cursor = Cursors.Hand,
                BorderBrush = Brushes.Black,
                BorderThickness = new Thickness(1)
            };

            if (isPrimary)
            {
                // VH GOUD voor primaire knop
                btn.Background = new SolidColorBrush(Color.FromRgb(178, 153, 107)); // VH Goud
                btn.Foreground = Brushes.White;
            }
            else
            {
                // Witte secundaire knop met zwarte border
                btn.Background = Brushes.White;
                btn.Foreground = Brushes.Black;
            }

            // Rounded corners
            var template = new ControlTemplate(typeof(Button));
            var borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.Name = "border";
            borderFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            borderFactory.SetValue(Border.BorderBrushProperty, Brushes.Black);
            borderFactory.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(8));
            borderFactory.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Button.PaddingProperty));

            var contentFactory = new FrameworkElementFactory(typeof(ContentPresenter));
            contentFactory.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentFactory.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            borderFactory.AppendChild(contentFactory);

            // Hover effect
            var hoverTrigger = new Trigger { Property = Button.IsMouseOverProperty, Value = true };
            if (isPrimary)
                hoverTrigger.Setters.Add(new Setter(Button.BackgroundProperty, new SolidColorBrush(Color.FromRgb(160, 138, 97)))); // Donkerder goud
            else
                hoverTrigger.Setters.Add(new Setter(Button.BackgroundProperty, new SolidColorBrush(Color.FromRgb(248, 248, 248)))); // Licht grijs hover

            template.VisualTree = borderFactory;
            template.Triggers.Add(hoverTrigger);
            btn.Template = template;

            return btn;
        }

        private List<LinePatternElement> GetLinePatterns()
        {
            var patterns = new FilteredElementCollector(_doc)
                .OfClass(typeof(LinePatternElement))
                .Cast<LinePatternElement>()
                .OrderBy(p => p.Name)
                .ToList();
            return patterns;
        }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            if (_cmb_linepattern.SelectedItem != null && _cmb_color.SelectedItem != null)
            {
                SelectedLinePatternId = (ElementId)((ComboBoxItem)_cmb_linepattern.SelectedItem).Tag;
                _lastSelectedPatternId = SelectedLinePatternId;
                
                SelectedColor = (Autodesk.Revit.DB.Color)((ComboBoxItem)_cmb_color.SelectedItem).Tag;
                _lastSelectedColor = SelectedColor;

                DialogResult = true;
            }
            else
            {
                DialogResult = false;
            }
            Close();
        }
    }
}
