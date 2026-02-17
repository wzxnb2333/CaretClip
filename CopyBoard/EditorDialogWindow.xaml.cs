using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace CopyBoard;

public partial class EditorDialogWindow : Window
{
    public string EditedText { get; private set; } = string.Empty;

    public EditorDialogWindow(string initialText, bool isDarkMode)
    {
        InitializeComponent();

        EditorTextBox.Text = initialText ?? string.Empty;
        ApplyTheme(isDarkMode);

        Loaded += (_, _) =>
        {
            EditorTextBox.Focus();
            EditorTextBox.CaretIndex = EditorTextBox.Text.Length;
        };
    }

    private void ApplyTheme(bool isDarkMode)
    {
        if (isDarkMode)
        {
            Background = BrushFromHex("#202020");
            Foreground = Brushes.White;
            EditorTextBox.Background = BrushFromHex("#2B2B2B");
            EditorTextBox.Foreground = Brushes.White;
            EditorTextBox.CaretBrush = Brushes.White;
        }
        else
        {
            Background = BrushFromHex("#F4F6FA");
            Foreground = BrushFromHex("#111827");
            EditorTextBox.Background = Brushes.White;
            EditorTextBox.Foreground = BrushFromHex("#111827");
            EditorTextBox.CaretBrush = BrushFromHex("#111827");
        }
    }

    private static Brush BrushFromHex(string hex)
    {
        return new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex)!);
    }

    protected override void OnPreviewKeyDown(KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            DialogResult = false;
            Close();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Enter && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
        {
            SaveAndClose();
            e.Handled = true;
            return;
        }

        base.OnPreviewKeyDown(e);
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        SaveAndClose();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void SaveAndClose()
    {
        EditedText = EditorTextBox.Text ?? string.Empty;
        DialogResult = true;
        Close();
    }
}
