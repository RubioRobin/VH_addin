using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using Grid = System.Windows.Controls.Grid;
using ComboBox = System.Windows.Controls.ComboBox;
using TextBox = System.Windows.Controls.TextBox;
using Color = System.Windows.Media.Color;

namespace VH_Addin.Views
{
    public partial class SheetCreatorWindow : Window
    {
        private Document _doc;
        private ComboBox _cmb_titleblock;
        private ComboBox _cmb_phase;
        private TextBox _txt_phaseCustom;
        private ComboBox _cmb_type;
        private TextBox _txt_typeCustom;
        private TextBox _txt_sequence;
        private TextBox _txt_sheetName;
        private TextBox _txt_sheetCount;
        
        private Dictionary<string, ElementId> _titleblocks = new Dictionary<string, ElementId>();

        private static double _lastWindowLeft = double.NaN;
        private static double _lastWindowTop = double.NaN;
        private static ElementId _lastTitleblockId = null;
        private static string _lastPhase = "UO";
        private static string _lastPhaseCustom = "";
        private static string _lastType = "01 – situatie";
        private static string _lastTypeCustom = "";
        private static string _lastSequence = "01";
        private static string _lastSequenceCustom = "";
        private static string _lastSheetName = "Nieuwe Tekening";
        private static string _lastSheetCount = "1";

        public SheetCreatorWindow(Document doc)
        {
            _doc = doc;
            InitializeComponent();
            Width = 420;
            MinWidth = 420;
            MaxWidth = 420;
            SizeToContent = SizeToContent.Height;
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            Topmost = true;

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
            
            LoadTitleblocks();
            BuildUI();
        }

        private void LoadTitleblocks()
        {
            try
            {
                var collector = new FilteredElementCollector(_doc)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .WhereElementIsElementType();

                foreach (Element t in collector)
                {
                    string name = t.Name;
                    var param = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                    if (param != null && !string.IsNullOrEmpty(param.AsString()))
                        name = param.AsString();
                    
                    if (!_titleblocks.ContainsKey(name))
                        _titleblocks.Add(name, t.Id);
                }
            }
            catch { }
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
                Text = "Sheet Creator",
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(30, 41, 59)),
                Margin = new Thickness(0, 0, 0, 16)
            });

            // 1. TITLEBLOCK
            contentStack.Children.Add(CreateLabel("1. Titleblock:"));
            _cmb_titleblock = CreateComboBox();
            foreach (var name in _titleblocks.Keys.OrderBy(k => k))
            {
                var item = new ComboBoxItem { Content = name, Tag = _titleblocks[name] };
                _cmb_titleblock.Items.Add(item);
                if (_lastTitleblockId != null && _titleblocks[name] == _lastTitleblockId)
                    _cmb_titleblock.SelectedItem = item;
            }
            if (_cmb_titleblock.SelectedIndex == -1 && _cmb_titleblock.Items.Count > 0) 
                _cmb_titleblock.SelectedIndex = 0;
            contentStack.Children.Add(_cmb_titleblock);

            // 2. SHEET NUMBERING
            contentStack.Children.Add(CreateLabel("2. Nummering (Fase - Type - Volgnr):"));
            
            // PHASE
            var phaseGrid = CreateThreePartGrid();
            _cmb_phase = CreateComboBox();
            string[] phases = { "DO", "TO", "VK", "UO", "BB", "OV", "anders" };
            foreach (var p in phases) _cmb_phase.Items.Add(p);
            _cmb_phase.SelectedItem = _lastPhase;
            _cmb_phase.SelectionChanged += (s, e) => UpdateCustomVisibility();
            Grid.SetColumn(_cmb_phase, 0);
            phaseGrid.Children.Add(_cmb_phase);

            _txt_phaseCustom = CreateTextBox("Fase...");
            _txt_phaseCustom.Text = _lastPhaseCustom;
            Grid.SetColumn(_txt_phaseCustom, 2);
            phaseGrid.Children.Add(_txt_phaseCustom);
            contentStack.Children.Add(phaseGrid);

            // TYPE
            var typeGrid = CreateThreePartGrid();
            _cmb_type = CreateComboBox();
            string[] types = { "01 – situatie", "11 – plattegrond", "21 – doorsnede", "31 – gevel", "41 – details", "51 – kozijnstaat", "61 – vormtekening", "71 – fragmenttekening", "81 – opstallen", "91 – artist impression", "anders" };
            foreach (var t in types) _cmb_type.Items.Add(t);
            _cmb_type.SelectedItem = _lastType;
            _cmb_type.SelectionChanged += (s, e) => UpdateCustomVisibility();
            Grid.SetColumn(_cmb_type, 0);
            typeGrid.Children.Add(_cmb_type);

            _txt_typeCustom = CreateTextBox("Type...");
            _txt_typeCustom.Text = _lastTypeCustom;
            Grid.SetColumn(_txt_typeCustom, 2);
            typeGrid.Children.Add(_txt_typeCustom);
            contentStack.Children.Add(typeGrid);

            // SEQUENCE
            contentStack.Children.Add(CreateLabel("Volgnummer (bijv. 01):"));
            _txt_sequence = CreateTextBox("01");
            _txt_sequence.Text = _lastSequence;
            contentStack.Children.Add(_txt_sequence);

            // 3. SHEET NAME
            contentStack.Children.Add(CreateLabel("3. Naam (voor alle sheets):"));
            _txt_sheetName = CreateTextBox("Nieuwe Tekening");
            _txt_sheetName.Text = _lastSheetName;
            contentStack.Children.Add(_txt_sheetName);

            // 4. SHEET COUNT
            contentStack.Children.Add(CreateLabel("4. Aantal sheets:"));
            _txt_sheetCount = CreateTextBox("1");
            _txt_sheetCount.Text = _lastSheetCount;
            contentStack.Children.Add(_txt_sheetCount);

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

            var btnOk = CreateModernButton("Sheets Aanmaken", true);
            btnOk.Click += OnOk;
            Grid.SetColumn(btnOk, 3);
            btnGrid.Children.Add(btnOk);

            Grid.SetRow(btnGrid, 1);
            mainGrid.Children.Add(btnGrid);

            this.Content = mainGrid;
            UpdateCustomVisibility();
        }

        private Grid CreateThreePartGrid()
        {
            var grid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(8) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            return grid;
        }

        private void UpdateCustomVisibility()
        {
            _txt_phaseCustom.IsEnabled = _cmb_phase.SelectedItem?.ToString() == "anders";
            _txt_typeCustom.IsEnabled = _cmb_type.SelectedItem?.ToString() == "anders";
            
            _txt_phaseCustom.Opacity = _txt_phaseCustom.IsEnabled ? 1.0 : 0.4;
            _txt_typeCustom.Opacity = _txt_typeCustom.IsEnabled ? 1.0 : 0.4;
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

        private ComboBox CreateComboBox()
        {
            return new ComboBox
            {
                Background = new SolidColorBrush(Color.FromRgb(249, 250, 251)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10, 8, 10, 8),
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(51, 65, 85))
            };
        }

        private TextBox CreateTextBox(string placeholder)
        {
            return new TextBox
            {
                Text = placeholder,
                Background = new SolidColorBrush(Color.FromRgb(249, 250, 251)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10, 8, 10, 8),
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                VerticalContentAlignment = VerticalAlignment.Center
            };
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

        private void OnOk(object sender, RoutedEventArgs e)
        {
            if (_cmb_titleblock.SelectedItem == null) return;
            if (!int.TryParse(_txt_sheetCount.Text, out int count) || count <= 0)
            {
                TaskDialog.Show("Fout", "Vul een geldig aantal sheets in.");
                return;
            }

            string selectedPhase = _cmb_phase.SelectedItem.ToString();
            string selectedType = _cmb_type.SelectedItem.ToString();

            string pPart = selectedPhase;
            if (pPart == "anders") pPart = _txt_phaseCustom.Text;
            
            string tPart = selectedType;
            if (tPart == "anders") tPart = _txt_typeCustom.Text;
            else if (tPart.Contains(" – ")) tPart = tPart.Split(' ')[0];
            
            string baseSeq = _txt_sequence.Text;

            ElementId tbId = (ElementId)((ComboBoxItem)_cmb_titleblock.SelectedItem).Tag;
            string sheetName = _txt_sheetName.Text;

            // Mapping dictionaries
            var typeOnderdeelMap = new Dictionary<string, string>
            {
                { "01 – situatie", "01_situatie" },
                { "11 – plattegrond", "11_plattegrond" },
                { "21 – doorsnede", "21_doorsnede" },
                { "31 – gevel", "31_gevels" },
                { "41 – details", "41_details" },
                { "51 – kozijnstaat", "51_kozijnstaat" },
                { "61 – vormtekening", "61_vormtekening" },
                { "71 – fragmenttekening", "71_fragmenttekening" },
                { "81 – opstallen", "81_opstallen" },
                { "91 – artist impression", "91_artist impression" }
            };

            var phaseMap = new Dictionary<string, string>
            {
                { "DO", "DO_Definitief Ontwerp" },
                { "TO", "TO_Technisch Ontwerp" },
                { "VK", "VK_Verkooptekeningen" },
                { "UO", "UO_Uitvoerings Ontwerp" },
                { "BB", "BBL_Bouwbesluit" },
                { "OV", "OV_Omgevingsvergunning" }
            };

            var drawingPhaseMap = new Dictionary<string, string>
            {
                { "DO", "Definitief Ontwerp" },
                { "TO", "Technisch Ontwerp" },
                { "VK", "Verkooptekening" },
                { "UO", "Uitvoeringsgereed Ontwerp" },
                { "BB", "BBL_Bouwbesluit" },
                { "OV", "Omgevingsvergunning" }
            };

            // SMART NUMBERING LOGIC FOR SEQUENCE
            string prefix = baseSeq.TrimEnd("0123456789".ToCharArray());
            string numPartStr = baseSeq.Substring(prefix.Length);
            int startNum = 0;
            int pad = 2;
            if (!string.IsNullOrEmpty(numPartStr))
            {
                int.TryParse(numPartStr, out startNum);
                pad = numPartStr.Length;
            }

            try
            {
                int successCount = 0;
                string dateStr = DateTime.Now.ToString("dd-MM-yy");

                using (Transaction t = new Transaction(_doc, "VH Sheet Creator"))
                {
                    t.Start();
                    for (int i = 0; i < count; i++)
                    {
                        string incSeq = prefix + (startNum + i).ToString().PadLeft(pad, '0');
                        string fullNum = $"{pPart}-{tPart}-{incSeq}";

                        try
                        {
                            ViewSheet newSheet = ViewSheet.Create(_doc, tbId);
                            newSheet.SheetNumber = fullNum;
                            newSheet.Name = sheetName;
                            
                            // 1. Populate Sheet Issue Date
                            Parameter pDate = newSheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE);
                            if (pDate != null && !pDate.IsReadOnly)
                                pDate.Set(dateStr);

                            // 2. Populate VH Parameters
                            SetParameter(newSheet, "VH_categorie", "TEKENWERK");
                            SetParameter(newSheet, "VH_tekening status", "Concept");

                            if (typeOnderdeelMap.TryGetValue(selectedType, out string onderdeel))
                                SetParameter(newSheet, "VH_onderdeel", onderdeel);
                            else
                                SetParameter(newSheet, "VH_onderdeel", "");

                            if (phaseMap.TryGetValue(selectedPhase, out string phaseVal))
                                SetParameter(newSheet, "VH_fase", phaseVal);
                            else
                                SetParameter(newSheet, "VH_fase", "");

                            if (drawingPhaseMap.TryGetValue(selectedPhase, out string drPhaseVal))
                                SetParameter(newSheet, "VH_tekening fase", drPhaseVal);
                            else
                                SetParameter(newSheet, "VH_tekening fase", "");
                                
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            // Skip if already exists or other error
                        }
                    }
                    t.Commit();
                }

                // Persistence
                _lastTitleblockId = tbId;
                _lastPhase = selectedPhase;
                _lastPhaseCustom = _txt_phaseCustom.Text;
                _lastType = selectedType;
                _lastTypeCustom = _txt_typeCustom.Text;
                _lastSequence = baseSeq;
                _lastSheetName = sheetName;
                _lastSheetCount = _txt_sheetCount.Text;

                TaskDialog.Show("Succes", $"{successCount} sheets succesvol aangemaakt.");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Fout", ex.Message);
            }
        }

        private void SetParameter(ViewSheet sheet, string paramName, string value)
        {
            Parameter p = sheet.LookupParameter(paramName);
            if (p != null && !p.IsReadOnly)
            {
                p.Set(value);
            }
        }
    }
}
