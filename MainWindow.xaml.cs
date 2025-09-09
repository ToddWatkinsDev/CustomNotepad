using Microsoft.Win32;
using System;
using System.IO;
using System.Printing;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Xps.Packaging;
using WpfMedia = System.Windows.Media;
using Wpf = System.Windows;
using WpfDocs = System.Windows.Documents;
using Drawing = System.Drawing;
using Forms = System.Windows.Forms;
using Win32 = Microsoft.Win32;
using WpfControls = System.Windows.Controls;
using WpfInput = System.Windows.Input;

namespace CustomNotepad
{
    public partial class MainWindow : Window
    {
        private string currentFilePath = "";

        public MainWindow()
        {
            InitializeComponent();
            InitializeFontTypeComboBox();

            MainRichTextBox.TextChanged += MainRichTextBox_TextChanged;
            MainRichTextBox.SelectionChanged += MainRichTextBox_SelectionChanged;

            UpdateStatusBar();

            // Set default to Light Mode on startup
            ApplyLightMode();
        }

        #region Initialization
        private void InitializeFontTypeComboBox()
        {
            foreach (WpfMedia.FontFamily font in WpfMedia.Fonts.SystemFontFamilies)
            {
                FontTypeComboBox.Items.Add(font.Source);
            }
        }
        #endregion

        #region File Operations
        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Win32.OpenFileDialog
            {
                Filter = "Text Files (*.txt)|*.txt|Rich Text Format (*.rtf)|*.rtf|All Files (*.*)|*.*"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    var range = new TextRange(MainRichTextBox.Document.ContentStart, MainRichTextBox.Document.ContentEnd);
                    using (var fs = new FileStream(dlg.FileName, FileMode.Open, FileAccess.Read))
                    {
                        if (Path.GetExtension(dlg.FileName).ToLower() == ".rtf")
                        {
                            range.Load(fs, Wpf.DataFormats.Rtf);
                        }
                        else
                        {
                            range.Load(fs, Wpf.DataFormats.Text);
                        }
                    }
                    currentFilePath = dlg.FileName;
                    UpdateStatusBar();
                }
                catch (Exception ex)
                {
                    Wpf.MessageBox.Show($"Error opening file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void SaveFile_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(currentFilePath))
            {
                SaveToFile(currentFilePath);
            }
            else
            {
                SaveFileAs(sender, e);
            }
        }

        private void SaveFileAs(object sender, RoutedEventArgs e)
        {
            var dlg = new Win32.SaveFileDialog
            {
                Filter = "Text Files (*.txt)|*.txt|Rich Text Format (*.rtf)|*.rtf"
            };

            if (dlg.ShowDialog() == true)
            {
                SaveToFile(dlg.FileName);
                currentFilePath = dlg.FileName;
                UpdateStatusBar();
            }
        }

        private void SaveToFile(string path)
        {
            try
            {
                var range = new TextRange(MainRichTextBox.Document.ContentStart, MainRichTextBox.Document.ContentEnd);
                using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    if (Path.GetExtension(path).ToLower() == ".rtf")
                    {
                        range.Save(fs, Wpf.DataFormats.Rtf);
                    }
                    else
                    {
                        range.Save(fs, Wpf.DataFormats.Text);
                    }
                }
            }
            catch (Exception ex)
            {
                Wpf.MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        #endregion

        #region Edit Operations
        private void Undo_Click(object sender, RoutedEventArgs e)
        {
            if (MainRichTextBox.CanUndo)
                MainRichTextBox.Undo();
        }

        private void Redo_Click(object sender, RoutedEventArgs e)
        {
            if (MainRichTextBox.CanRedo)
                MainRichTextBox.Redo();
        }

        private void Cut_Click(object sender, RoutedEventArgs e)
        {
            MainRichTextBox.Cut();
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            MainRichTextBox.Copy();
        }

        private void Paste_Click(object sender, RoutedEventArgs e)
        {
            MainRichTextBox.Paste();
        }
        #endregion

        #region Formatting
        private void BoldButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleProperty(TextElement.FontWeightProperty, Wpf.FontWeights.Bold, Wpf.FontWeights.Normal);
        }

        private void ItalicButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleProperty(TextElement.FontStyleProperty, Wpf.FontStyles.Italic, Wpf.FontStyles.Normal);
        }

        private void UnderlineButton_Click(object sender, RoutedEventArgs e)
        {
            var selection = MainRichTextBox.Selection;
            if (!selection.IsEmpty)
            {
                var currentDecorations = selection.GetPropertyValue(Inline.TextDecorationsProperty) as Wpf.TextDecorationCollection;

                if (currentDecorations == null || !currentDecorations.Contains(Wpf.TextDecorations.Underline[0]))
                {
                    selection.ApplyPropertyValue(Inline.TextDecorationsProperty, Wpf.TextDecorations.Underline);
                }
                else
                {
                    var newDecorations = new Wpf.TextDecorationCollection(currentDecorations);
                    newDecorations.Remove(Wpf.TextDecorations.Underline[0]);
                    selection.ApplyPropertyValue(Inline.TextDecorationsProperty, newDecorations.Count > 0 ? newDecorations : null);
                }
            }
        }

        private void ToggleProperty(DependencyProperty property, object onValue, object offValue)
        {
            var selection = MainRichTextBox.Selection;
            if (!selection.IsEmpty)
            {
                var currentValue = selection.GetPropertyValue(property);
                if (currentValue == DependencyProperty.UnsetValue || !currentValue.Equals(onValue))
                {
                    selection.ApplyPropertyValue(property, onValue);
                }
                else
                {
                    selection.ApplyPropertyValue(property, offValue);
                }
            }
        }

        private void FontTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FontTypeComboBox.SelectedItem != null)
            {
                var fontName = FontTypeComboBox.SelectedItem.ToString();
                var selection = MainRichTextBox.Selection;

                if (!selection.IsEmpty)
                {
                    selection.ApplyPropertyValue(TextElement.FontFamilyProperty, new WpfMedia.FontFamily(fontName));
                }
                else
                {
                    MainRichTextBox.FontFamily = new WpfMedia.FontFamily(fontName);
                }
            }
        }

        private void FontSizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FontSizeComboBox.SelectedItem is ComboBoxItem item && item.Content != null)
            {
                var selectedSize = item.Content.ToString();

                if (double.TryParse(selectedSize, out double fontSize))
                {
                    var selection = MainRichTextBox.Selection;

                    if (!selection.IsEmpty)
                    {
                        selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize);
                    }
                    else
                    {
                        MainRichTextBox.FontSize = fontSize;
                    }
                }
            }
        }

        private void FontColorButton_Click(object sender, RoutedEventArgs e)
        {
            var colorDialog = new Forms.ColorDialog();
            if (colorDialog.ShowDialog() == Forms.DialogResult.OK)
            {
                var color = colorDialog.Color;
                var brush = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(color.A, color.R, color.G, color.B));
                var selection = MainRichTextBox.Selection;

                if (!selection.IsEmpty)
                {
                    selection.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
                }
                else
                {
                    MainRichTextBox.Foreground = brush;
                }
            }
        }

        private void AlignLeft_Click(object sender, RoutedEventArgs e)
        {
            ApplyParagraphAlignment(TextAlignment.Left);
        }

        private void AlignCenter_Click(object sender, RoutedEventArgs e)
        {
            ApplyParagraphAlignment(TextAlignment.Center);
        }

        private void AlignRight_Click(object sender, RoutedEventArgs e)
        {
            ApplyParagraphAlignment(TextAlignment.Right);
        }

        private void AlignJustify_Click(object sender, RoutedEventArgs e)
        {
            ApplyParagraphAlignment(TextAlignment.Justify);
        }

        private void ApplyParagraphAlignment(TextAlignment alignment)
        {
            var selection = MainRichTextBox.Selection;
            if (selection != null)
            {
                selection.ApplyPropertyValue(Paragraph.TextAlignmentProperty, alignment);
            }
        }

        private void IndentParagraph_Click(object sender, RoutedEventArgs e)
        {
            var selection = MainRichTextBox.Selection;
            if (selection != null)
            {
                var currentIndent = selection.GetPropertyValue(Paragraph.TextIndentProperty);
                var indentValue = 0.0;

                if (currentIndent != DependencyProperty.UnsetValue)
                    indentValue = (double)currentIndent;

                selection.ApplyPropertyValue(Paragraph.TextIndentProperty, indentValue + 20); // increase indent by 20
            }
        }

        private void LineSpacing_Click(object sender, RoutedEventArgs e)
        {
            var input = Microsoft.VisualBasic.Interaction.InputBox("Enter line spacing (e.g., 1.0, 1.5, 2.0):", "Line Spacing", "1.0");
            if (double.TryParse(input, out double spacing))
            {
                var selection = MainRichTextBox.Selection;
                if (selection != null)
                {
                    selection.ApplyPropertyValue(Paragraph.LineHeightProperty, spacing * MainRichTextBox.FontSize);
                }
            }
        }

        private void BulletList_Click(object sender, RoutedEventArgs e)
        {
            WpfDocs.EditingCommands.ToggleBullets.Execute(null, MainRichTextBox);
        }

        private void NumberedList_Click(object sender, RoutedEventArgs e)
        {
            WpfDocs.EditingCommands.ToggleNumbering.Execute(null, MainRichTextBox);
        }
        #endregion

        #region Export & Print
        private void ExportPdf_Click(object sender, RoutedEventArgs e)
        {
            var printDlg = new WpfControls.PrintDialog();
            printDlg.PrintTicket.PageOrientation = System.Printing.PageOrientation.Portrait;

            var localPrintServer = new PrintServer();
            PrintQueue? pdfPrinter = null;
            foreach (var queue in localPrintServer.GetPrintQueues())
            {
                if (queue.Name.Contains("Microsoft Print to PDF"))
                {
                    pdfPrinter = queue;
                    break;
                }
            }

            if (pdfPrinter != null)
            {
                printDlg.PrintQueue = pdfPrinter;
            }

            if (printDlg.ShowDialog() == true)
            {
                printDlg.PrintDocument(((IDocumentPaginatorSource)MainRichTextBox.Document).DocumentPaginator, "Export PDF Document");
            }
            else
            {
                Wpf.MessageBox.Show("Export to PDF cancelled or failed.", "PDF Export", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void PrintFile_Click(object sender, RoutedEventArgs e)
        {
            var printDlg = new WpfControls.PrintDialog();
            if (printDlg.ShowDialog() == true)
            {
                printDlg.PrintDocument(((IDocumentPaginatorSource)MainRichTextBox.Document).DocumentPaginator, "Print Document");
            }
        }
        #endregion

        #region Theme Switching

        private void ThemeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem clickedItem)
            {
                // Uncheck all theme menu items
                LightModeMenuItem.IsChecked = false;
                DarkModeMenuItem.IsChecked = false;
                GitHubModeMenuItem.IsChecked = false;
                // Add more unchecked lines here for additional modes

                clickedItem.IsChecked = true;

                switch (clickedItem.Name)
                {
                    case "LightModeMenuItem":
                        ApplyLightMode();
                        break;
                    case "DarkModeMenuItem":
                        ApplyDarkMode();
                        break;
                    case "GitHubModeMenuItem":
                        ApplyGitHubMode();
                        break;
                    // Add more cases here for additional modes
                }
            }
        }

        private void ApplyLightMode()
        {
            this.Background = WpfMedia.Brushes.White;
            MainRichTextBox.Background = WpfMedia.Brushes.White;
            MainRichTextBox.Foreground = WpfMedia.Brushes.Black;
            // Style other controls if needed for light mode
        }

        private void ApplyDarkMode()
        {
            this.Background = WpfMedia.Brushes.Black;
            MainRichTextBox.Background = WpfMedia.Brushes.Black;
            MainRichTextBox.Foreground = WpfMedia.Brushes.LightGray;
            // Style other controls if needed for dark mode
        }

        private void ApplyGitHubMode()
        {
            var bg = (WpfMedia.Brush)new WpfMedia.SolidColorBrush((WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString("#0d1117"));
            var fg = (WpfMedia.Brush)new WpfMedia.SolidColorBrush((WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString("#c9d1d9"));
            this.Background = bg;
            MainRichTextBox.Background = bg;
            MainRichTextBox.Foreground = fg;
            // Style other controls if needed for GitHub mode
        }

        #endregion

        #region Dark Mode Toggle (Deprecated)
        // You can optionally keep the old toggle dark mode button handler or remove it to avoid confusion
        private void ToggleDarkMode_Click(object sender, RoutedEventArgs e)
        {
            if (DarkModeMenuItem != null)
            {
                DarkModeMenuItem.IsChecked = true;
            ThemeMenuItem_Click(DarkModeMenuItem, new RoutedEventArgs());
            }
        }
        #endregion

        #region Status Bar and Auto-save
        private void MainRichTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateStatusBar();

            // TODO: Implement auto-save with debounce timer
        }

        private void MainRichTextBox_SelectionChanged(object sender, RoutedEventArgs e)
        {
            UpdateCursorPosition();
        }

        private void UpdateStatusBar()
        {
            var range = new TextRange(MainRichTextBox.Document.ContentStart, MainRichTextBox.Document.ContentEnd);
            var text = range.Text;

            var matches = Regex.Matches(text, @"\b\w+\b");
            StatusWordCount.Text = $"Words: {matches.Count}";

            if (string.IsNullOrEmpty(currentFilePath))
                StatusFileName.Text = "No file";
            else
                StatusFileName.Text = Path.GetFileName(currentFilePath);

            UpdateCursorPosition();
        }

        private void UpdateCursorPosition()
        {
            var caret = MainRichTextBox.CaretPosition;

            int line = GetLineNumber(caret);
            int col = GetColumnNumber(caret);

            StatusCursorPosition.Text = $"Line: {line}, Col: {col}";
        }

        private int GetLineNumber(TextPointer caret)
        {
            var start = MainRichTextBox.Document.ContentStart;
            int line = 1;
            var pointer = start;

            while (pointer != null && pointer.CompareTo(caret) < 0)
            {
                if (pointer.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.ElementStart)
                {
                    line++;
                }
                pointer = pointer.GetNextContextPosition(LogicalDirection.Forward);
            }
            return line;
        }

        private int GetColumnNumber(TextPointer caret)
        {
            var lineStart = caret.GetLineStartPosition(0);
            if (lineStart == null)
                return 1;
            return lineStart.GetOffsetToPosition(caret) + 1;
        }
        #endregion
    }
}
