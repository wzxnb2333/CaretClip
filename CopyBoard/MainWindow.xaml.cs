using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Windows.Automation;
using CopyBoard.Interop;
using CopyBoard.Models;
using CopyBoard.Services;
using Microsoft.Win32;
using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace CopyBoard;

public partial class MainWindow : Window
{
    private const int HotKeyId = 10001;
    private const uint VkV = 0x56;
    private const double MouseWheelScrollStep = 64;

    private readonly StorageService _storageService = new();
    private readonly StartupService _startupService = new();
    private readonly ObservableCollection<ClipboardEntry> _history = new();
    private readonly DispatcherTimer _cleanupTimer = new() { Interval = TimeSpan.FromMinutes(5) };
    private readonly ICollectionView _historyView;

    private AppSettings _settings = new();
    private HwndSource? _hwndSource;
    private IntPtr _windowHandle;
    private bool _hotKeyRegistered;
    private bool _clipboardListenerRegistered;
    private bool _ignoreClipboardUpdate;
    private bool _isExitRequested;
    private bool _isInitializingUiState;
    private bool _isHidingWithAnimation;
    private bool _isEditorDialogOpen;
    private string _searchQuery = string.Empty;
    private double _historyScrollTarget;
    private ClipboardEntry? _editingEntry;
    private ScrollViewer? _historyScrollViewer;
    private ScrollViewer? _editTextScrollViewer;
    private Forms.NotifyIcon? _notifyIcon;
    private Forms.ToolStripMenuItem? _trayStartupMenuItem;

    private static readonly DependencyProperty HistoryAnimatedVerticalOffsetProperty =
        DependencyProperty.Register(
            "HistoryAnimatedVerticalOffset",
            typeof(double),
            typeof(MainWindow),
            new PropertyMetadata(0d, OnHistoryAnimatedVerticalOffsetChanged));

    public MainWindow()
    {
        InitializeComponent();

        _historyView = CollectionViewSource.GetDefaultView(_history);
        _historyView.Filter = FilterHistory;
        _historyView.SortDescriptions.Add(new SortDescription(nameof(ClipboardEntry.IsPinned), ListSortDirection.Descending));
        _historyView.SortDescriptions.Add(new SortDescription(nameof(ClipboardEntry.CreatedAtUtc), ListSortDirection.Descending));
        HistoryListBox.ItemsSource = _historyView;

        SourceInitialized += MainWindow_SourceInitialized;
        Loaded += MainWindow_Loaded;
        Closing += MainWindow_Closing;
        Deactivated += MainWindow_Deactivated;
        PreviewKeyDown += MainWindow_PreviewKeyDown;
        _cleanupTimer.Tick += CleanupTimer_Tick;
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        _settings = _storageService.LoadSettings();
        NormalizeSettings();

        foreach (var item in _storageService.LoadHistory())
        {
            _history.Add(item);
        }

        ApplyTheme(_settings.ThemeMode);
        SetStartupEnabled(_settings.StartWithWindows, persistSettings: false);
        ApplySettingsToControls();
        ApplyRetentionAndLimit(saveImmediately: false);
        RefreshHistoryView();
        PopulateAboutInfo();
        InitializeTrayIcon();
        PositionWindowForQuickShow();

        _cleanupTimer.Start();
        UpdateStatus("已启动，正在监听剪贴板。可通过 Alt+Shift+V 随时唤出。");
        AnimateWindowShow();
    }

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        _windowHandle = new WindowInteropHelper(this).Handle;
        _hwndSource = HwndSource.FromHwnd(_windowHandle);
        _hwndSource?.AddHook(WndProc);

        _clipboardListenerRegistered = NativeMethods.AddClipboardFormatListener(_windowHandle);
        _hotKeyRegistered = NativeMethods.RegisterHotKey(
            _windowHandle,
            HotKeyId,
            NativeMethods.MOD_ALT | NativeMethods.MOD_SHIFT,
            VkV);

        if (!_hotKeyRegistered)
        {
            UpdateStatus("热键 Alt+Shift+V 注册失败，可能被其他程序占用。");
        }

        ApplyTheme(_settings.ThemeMode);
    }

    private void MainWindow_Closing(object? sender, CancelEventArgs e)
    {
        if (!_isExitRequested)
        {
            e.Cancel = true;
            HideWithAnimation();
            UpdateStatus("程序已最小化到系统托盘。");
            return;
        }

        _cleanupTimer.Stop();
        CleanupNativeHooks();
        DisposeTrayIcon();
    }

    private void MainWindow_Deactivated(object? sender, EventArgs e)
    {
        if (!IsVisible || _isExitRequested || _isHidingWithAnimation || _isEditorDialogOpen)
        {
            return;
        }

        HideWithAnimation();
    }

    private void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape && SettingsOverlay.Visibility == Visibility.Visible)
        {
            SettingsOverlay.Visibility = Visibility.Collapsed;
            e.Handled = true;
        }
    }

    private void CleanupTimer_Tick(object? sender, EventArgs e)
    {
        ApplyRetentionAndLimit();
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        switch (msg)
        {
            case NativeMethods.WM_CLIPBOARDUPDATE:
                if (!_ignoreClipboardUpdate)
                {
                    TryCaptureClipboard();
                }

                handled = true;
                break;

            case NativeMethods.WM_HOTKEY:
                if (wParam.ToInt32() == HotKeyId)
                {
                    ToggleWindow();
                    handled = true;
                }

                break;
        }

        return IntPtr.Zero;
    }

    private void ToggleWindow()
    {
        if (!IsVisible || WindowState == WindowState.Minimized)
        {
            PositionWindowForQuickShow();
            ShowAndActivateWindow();
            UpdateStatus("已通过 Alt+Shift+V 唤出窗口。");
            return;
        }

        if (IsActive)
        {
            HideWithAnimation();
            return;
        }

        PositionWindowForQuickShow();
        ShowAndActivateWindow();
    }

    private void ShowAndActivateWindow()
    {
        if (!IsVisible)
        {
            Show();
        }

        WindowState = WindowState.Normal;
        Topmost = true;
        Activate();
        Focus();
        AnimateWindowShow();
    }

    private void AnimateWindowShow()
    {
        WindowContentRoot.BeginAnimation(UIElement.OpacityProperty, null);
        RootScaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, null);
        RootScaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, null);
        RootTranslateTransform.BeginAnimation(TranslateTransform.YProperty, null);

        WindowContentRoot.Opacity = 0;
        RootScaleTransform.ScaleX = 0.985;
        RootScaleTransform.ScaleY = 0.985;
        RootTranslateTransform.Y = 8;

        var fadeAnimation = new DoubleAnimation(1, TimeSpan.FromMilliseconds(140))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };

        var scaleAnimation = new DoubleAnimation(1, TimeSpan.FromMilliseconds(170))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };

        var moveAnimation = new DoubleAnimation(0, TimeSpan.FromMilliseconds(170))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };

        WindowContentRoot.BeginAnimation(UIElement.OpacityProperty, fadeAnimation);
        RootScaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnimation);
        RootScaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnimation);
        RootTranslateTransform.BeginAnimation(TranslateTransform.YProperty, moveAnimation);
    }

    private void HideWithAnimation()
    {
        if (!IsVisible || _isHidingWithAnimation)
        {
            return;
        }

        _isHidingWithAnimation = true;
        WindowContentRoot.BeginAnimation(UIElement.OpacityProperty, null);
        RootScaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, null);
        RootScaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, null);
        RootTranslateTransform.BeginAnimation(TranslateTransform.YProperty, null);

        var fadeAnimation = new DoubleAnimation(0, TimeSpan.FromMilliseconds(150))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
        };

        var scaleAnimation = new DoubleAnimation(0.982, TimeSpan.FromMilliseconds(165))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
        };

        var moveAnimation = new DoubleAnimation(10, TimeSpan.FromMilliseconds(165))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
        };

        fadeAnimation.Completed += (_, _) =>
        {
            _isHidingWithAnimation = false;
            Hide();
            WindowContentRoot.Opacity = 1;
            RootScaleTransform.ScaleX = 1;
            RootScaleTransform.ScaleY = 1;
            RootTranslateTransform.Y = 0;
            SettingsOverlay.Visibility = Visibility.Collapsed;
        };

        WindowContentRoot.BeginAnimation(UIElement.OpacityProperty, fadeAnimation);
        RootScaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnimation);
        RootScaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnimation);
        RootTranslateTransform.BeginAnimation(TranslateTransform.YProperty, moveAnimation);
    }

    private void TryCaptureClipboard()
    {
        var entry = ReadClipboardEntry();
        if (entry is null)
        {
            return;
        }

        var latest = _history.OrderByDescending(x => x.CreatedAtUtc).FirstOrDefault();
        if (latest is not null && latest.ContentEquals(entry))
        {
            return;
        }

        _history.Add(entry);
        ApplyRetentionAndLimit(saveImmediately: false);
        RefreshHistoryView();
        PersistHistory();
        UpdateStatus($"已记录 {entry.DisplayType} 内容，共 {_history.Count} 条。");
    }

    private ClipboardEntry? ReadClipboardEntry()
    {
        for (var attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                if (Clipboard.ContainsText())
                {
                    var text = Clipboard.GetText();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return new ClipboardEntry
                        {
                            ContentType = ClipboardContentType.Text,
                            TextContent = text,
                            CreatedAtUtc = DateTime.UtcNow
                        };
                    }
                }

                if (Clipboard.ContainsImage())
                {
                    var image = Clipboard.GetImage();
                    if (image is not null)
                    {
                        return new ClipboardEntry
                        {
                            ContentType = ClipboardContentType.Image,
                            ImageBase64 = EncodeImage(image),
                            CreatedAtUtc = DateTime.UtcNow
                        };
                    }
                }

                if (Clipboard.ContainsFileDropList())
                {
                    StringCollection? files = Clipboard.GetFileDropList();
                    if (files.Count > 0)
                    {
                        return new ClipboardEntry
                        {
                            ContentType = ClipboardContentType.FileList,
                            FilePaths = files.Cast<string>().ToList(),
                            CreatedAtUtc = DateTime.UtcNow
                        };
                    }
                }

                return null;
            }
            catch (COMException)
            {
                Dispatcher.Invoke(() => { }, DispatcherPriority.Background);
            }
        }

        return null;
    }

    private void PositionWindowForQuickShow()
    {
        var windowWidth = ResolveWindowWidth();
        var windowHeight = ResolveWindowHeight();

        if (TryGetCaretScreenPoint(out var caretPoint))
        {
            var area = Forms.Screen.FromPoint(caretPoint).WorkingArea;
            var targetX = caretPoint.X + 10.0;
            var targetY = caretPoint.Y + 16.0;

            if (targetY + windowHeight > area.Bottom)
            {
                targetY = caretPoint.Y - windowHeight - 16.0;
            }

            targetX = Math.Clamp(targetX, area.Left + 8.0, area.Right - windowWidth - 8.0);
            targetY = Math.Clamp(targetY, area.Top + 8.0, area.Bottom - windowHeight - 8.0);

            Left = targetX;
            Top = targetY;
            return;
        }

        var fallbackArea = GetPreferredWorkArea();
        Left = fallbackArea.Right - windowWidth - 12.0;
        Top = fallbackArea.Bottom - windowHeight - 12.0;
    }

    private double ResolveWindowWidth()
    {
        if (Width > 0)
        {
            return Width;
        }

        if (ActualWidth > 0)
        {
            return ActualWidth;
        }

        return 440;
    }

    private double ResolveWindowHeight()
    {
        if (Height > 0)
        {
            return Height;
        }

        if (ActualHeight > 0)
        {
            return ActualHeight;
        }

        return 470;
    }

    private static Drawing.Rectangle GetPreferredWorkArea()
    {
        var foregroundWindow = NativeMethods.GetForegroundWindow();
        if (foregroundWindow != IntPtr.Zero)
        {
            return Forms.Screen.FromHandle(foregroundWindow).WorkingArea;
        }

        return Forms.Screen.PrimaryScreen?.WorkingArea
               ?? new Drawing.Rectangle(
                   0,
                   0,
                   (int)SystemParameters.WorkArea.Width,
                   (int)SystemParameters.WorkArea.Height);
    }

    private static bool TryGetCaretScreenPoint(out Drawing.Point caretPoint)
    {
        caretPoint = default;

        if (TryGetCaretPointViaAttachedInput(out caretPoint) && IsLikelyValidCaretPoint(caretPoint))
        {
            return true;
        }

        if (TryGetCaretPointViaGuiThreadInfo(out caretPoint) && IsLikelyValidCaretPoint(caretPoint))
        {
            return true;
        }

        if (TryGetCaretPointViaUiAutomation(out caretPoint) && IsLikelyValidCaretPoint(caretPoint))
        {
            return true;
        }

        caretPoint = default;
        return false;
    }

    private static bool TryGetCaretPointViaAttachedInput(out Drawing.Point caretPoint)
    {
        caretPoint = default;

        var foregroundWindow = NativeMethods.GetForegroundWindow();
        if (foregroundWindow == IntPtr.Zero)
        {
            return false;
        }

        var threadId = NativeMethods.GetWindowThreadProcessId(foregroundWindow, out _);
        if (threadId == 0)
        {
            return false;
        }

        var currentThreadId = NativeMethods.GetCurrentThreadId();
        var attached = false;
        try
        {
            if (threadId != currentThreadId)
            {
                attached = NativeMethods.AttachThreadInput(currentThreadId, threadId, true);
            }

            var focusHandle = NativeMethods.GetFocus();
            if (focusHandle == IntPtr.Zero)
            {
                return false;
            }

            if (!NativeMethods.GetCaretPos(out var localCaretPoint))
            {
                return false;
            }

            if (!NativeMethods.ClientToScreen(focusHandle, ref localCaretPoint))
            {
                return false;
            }

            caretPoint = new Drawing.Point(localCaretPoint.X, localCaretPoint.Y);
            return true;
        }
        finally
        {
            if (attached)
            {
                NativeMethods.AttachThreadInput(currentThreadId, threadId, false);
            }
        }
    }

    private static bool TryGetCaretPointViaGuiThreadInfo(out Drawing.Point caretPoint)
    {
        caretPoint = default;

        var guiThreadInfo = new NativeMethods.GUITHREADINFO
        {
            cbSize = Marshal.SizeOf<NativeMethods.GUITHREADINFO>()
        };

        if (!NativeMethods.GetGUIThreadInfo(0, ref guiThreadInfo) || guiThreadInfo.hwndCaret == IntPtr.Zero)
        {
            return false;
        }

        if (guiThreadInfo.hwndFocus == IntPtr.Zero)
        {
            return false;
        }

        if (guiThreadInfo.rcCaret.Right <= guiThreadInfo.rcCaret.Left &&
            guiThreadInfo.rcCaret.Bottom <= guiThreadInfo.rcCaret.Top)
        {
            return false;
        }

        var point = new NativeMethods.POINT
        {
            X = guiThreadInfo.rcCaret.Left,
            Y = guiThreadInfo.rcCaret.Bottom
        };

        if (!NativeMethods.ClientToScreen(guiThreadInfo.hwndCaret, ref point))
        {
            return false;
        }

        caretPoint = new Drawing.Point(point.X, point.Y);
        return true;
    }

    private static bool TryGetCaretPointViaUiAutomation(out Drawing.Point caretPoint)
    {
        caretPoint = default;

        try
        {
            var focusedElement = AutomationElement.FocusedElement;
            if (focusedElement is null)
            {
                return false;
            }

            if (!focusedElement.TryGetCurrentPattern(TextPattern.Pattern, out var patternObject) || patternObject is not TextPattern textPattern)
            {
                return false;
            }

            var selection = textPattern.GetSelection();
            if (selection is null || selection.Length == 0 || selection[0] is null)
            {
                return false;
            }

            var rectangles = selection[0].GetBoundingRectangles();
            if (rectangles is null || rectangles.Length == 0)
            {
                return false;
            }

            var rect = rectangles[0];
            if (rect.IsEmpty)
            {
                return false;
            }

            caretPoint = new Drawing.Point((int)Math.Round(rect.Left), (int)Math.Round(rect.Bottom));
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsLikelyValidCaretPoint(Drawing.Point point)
    {
        if (point.X == 0 && point.Y == 0)
        {
            return false;
        }

        var virtualLeft = (int)SystemParameters.VirtualScreenLeft;
        var virtualTop = (int)SystemParameters.VirtualScreenTop;
        var virtualRight = virtualLeft + (int)SystemParameters.VirtualScreenWidth;
        var virtualBottom = virtualTop + (int)SystemParameters.VirtualScreenHeight;

        return point.X >= virtualLeft &&
               point.X <= virtualRight &&
               point.Y >= virtualTop &&
               point.Y <= virtualBottom;
    }

    private static string EncodeImage(BitmapSource image)
    {
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(image));
        using var ms = new MemoryStream();
        encoder.Save(ms);
        return Convert.ToBase64String(ms.ToArray());
    }

    private static BitmapImage? DecodeImage(string? base64)
    {
        if (string.IsNullOrWhiteSpace(base64))
        {
            return null;
        }

        var bytes = Convert.FromBase64String(base64);
        using var ms = new MemoryStream(bytes);
        var image = new BitmapImage();
        image.BeginInit();
        image.CacheOption = BitmapCacheOption.OnLoad;
        image.StreamSource = ms;
        image.EndInit();
        image.Freeze();
        return image;
    }

    private void NormalizeSettings()
    {
        if (_settings.RetentionDays is not (1 or 7 or 30 or -1))
        {
            _settings.RetentionDays = 30;
        }

        if (_settings.MaxHistoryItems < 10)
        {
            _settings.MaxHistoryItems = 1000;
        }

        if (_settings.ThemeMode is not ("System" or "Light" or "Dark"))
        {
            _settings.ThemeMode = "System";
        }
    }

    private void ApplySettingsToControls()
    {
        _isInitializingUiState = true;

        Retention1DayRadio.IsChecked = _settings.RetentionDays == 1;
        Retention7DayRadio.IsChecked = _settings.RetentionDays == 7;
        Retention30DayRadio.IsChecked = _settings.RetentionDays == 30;
        RetentionForeverRadio.IsChecked = _settings.RetentionDays == -1;

        ThemeSystemRadio.IsChecked = string.Equals(_settings.ThemeMode, "System", StringComparison.OrdinalIgnoreCase);
        ThemeLightRadio.IsChecked = string.Equals(_settings.ThemeMode, "Light", StringComparison.OrdinalIgnoreCase);
        ThemeDarkRadio.IsChecked = string.Equals(_settings.ThemeMode, "Dark", StringComparison.OrdinalIgnoreCase);

        SettingsStartWithWindowsCheckBox.IsChecked = _settings.StartWithWindows;
        if (_trayStartupMenuItem is not null)
        {
            _trayStartupMenuItem.Checked = _settings.StartWithWindows;
        }

        _isInitializingUiState = false;
    }

    private void ApplyRetentionAndLimit(bool saveImmediately = true)
    {
        var changed = false;

        if (_settings.RetentionDays > 0)
        {
            var cutoffUtc = DateTime.UtcNow.AddDays(-_settings.RetentionDays);
            for (var i = _history.Count - 1; i >= 0; i--)
            {
                if (_history[i].CreatedAtUtc < cutoffUtc && !_history[i].IsPinned)
                {
                    _history.RemoveAt(i);
                    changed = true;
                }
            }
        }

        var max = Math.Max(10, _settings.MaxHistoryItems);
        while (_history.Count > max)
        {
            var removable = _history
                .Where(x => !x.IsPinned)
                .OrderBy(x => x.CreatedAtUtc)
                .FirstOrDefault();

            if (removable is null)
            {
                break;
            }

            _history.Remove(removable);
            changed = true;
        }

        if (changed)
        {
            RefreshHistoryView();
            if (saveImmediately)
            {
                PersistHistory();
            }
        }
    }

    private void RefreshHistoryView()
    {
        _historyView.Refresh();
        UpdateSelectionState();
    }

    private static void OnHistoryAnimatedVerticalOffsetChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
    {
        if (dependencyObject is MainWindow window && window._historyScrollViewer is not null)
        {
            window._historyScrollViewer.ScrollToVerticalOffset((double)e.NewValue);
        }
    }

    private bool FilterHistory(object item)
    {
        if (item is not ClipboardEntry entry)
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(_searchQuery))
        {
            return true;
        }

        var query = _searchQuery.Trim();
        if (query.Length == 0)
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(entry.TextContent) &&
            entry.TextContent.Contains(query, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(entry.PreviewText) &&
            entry.PreviewText.Contains(query, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (entry.FilePaths is not null && entry.FilePaths.Any(path => path.Contains(query, StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        return false;
    }

    private void PopulateAboutInfo()
    {
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        AboutVersionTextBlock.Text = $"版本: {(version is null ? "开发版" : version.ToString(3))}";
        AboutDataPathTextBlock.Text = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "CaretClip");
    }

    private void PersistHistory()
    {
        _storageService.SaveHistory(_history.ToList());
    }

    private void PersistSettings()
    {
        _storageService.SaveSettings(_settings);
    }

    private void UpdateStatus(string text)
    {
        _ = text;
    }

    private void UpdateSelectionState()
    {
        var selected = HistoryListBox.SelectedItem as ClipboardEntry;
        var hasSelection = selected is not null;

        CopyButton.IsEnabled = hasSelection;
        DeleteButton.IsEnabled = hasSelection;
        PinButton.IsEnabled = hasSelection;
        EditButton.IsEnabled = selected?.IsEditable == true;
        PinButton.Content = selected?.IsPinned == true ? "\uE77A" : "\uE718";
        PinButton.ToolTip = selected?.IsPinned == true ? "取消置顶" : "置顶条目";
    }

    private void HistoryListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateSelectionState();
    }

    private void HistoryListBox_Loaded(object sender, RoutedEventArgs e)
    {
        _historyScrollViewer ??= FindVisualChild<ScrollViewer>(HistoryListBox);
        if (_historyScrollViewer is null)
        {
            return;
        }

        _historyScrollTarget = _historyScrollViewer.VerticalOffset;
        _historyScrollViewer.ScrollChanged -= HistoryScrollViewer_ScrollChanged;
        _historyScrollViewer.ScrollChanged += HistoryScrollViewer_ScrollChanged;
    }

    private void HistoryScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        _historyScrollTarget = e.VerticalOffset;
    }

    private void HistoryListBox_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        _historyScrollViewer ??= FindVisualChild<ScrollViewer>(HistoryListBox);
        if (_historyScrollViewer is null || _historyScrollViewer.ScrollableHeight <= 0)
        {
            return;
        }

        e.Handled = true;
        _historyScrollTarget -= (e.Delta / 120.0) * MouseWheelScrollStep;
        _historyScrollTarget = Math.Clamp(_historyScrollTarget, 0, _historyScrollViewer.ScrollableHeight);

        var animation = new DoubleAnimation(_historyScrollViewer.VerticalOffset, _historyScrollTarget, TimeSpan.FromMilliseconds(180))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };

        BeginAnimation(HistoryAnimatedVerticalOffsetProperty, animation, HandoffBehavior.SnapshotAndReplace);
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        var childCount = VisualTreeHelper.GetChildrenCount(parent);
        for (var i = 0; i < childCount; i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T typedChild)
            {
                return typedChild;
            }

            var nestedChild = FindVisualChild<T>(child);
            if (nestedChild is not null)
            {
                return nestedChild;
            }
        }

        return null;
    }

    private void HistoryListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        CopySelectedToClipboard();
    }

    private void CopySelected_Click(object sender, RoutedEventArgs e)
    {
        CopySelectedToClipboard();
    }

    private void CopySelectedToClipboard()
    {
        if (HistoryListBox.SelectedItem is not ClipboardEntry selected)
        {
            return;
        }

        _ignoreClipboardUpdate = true;
        try
        {
            switch (selected.ContentType)
            {
                case ClipboardContentType.Text:
                    Clipboard.SetText(selected.TextContent ?? string.Empty);
                    break;
                case ClipboardContentType.Image:
                {
                    var image = DecodeImage(selected.ImageBase64);
                    if (image is not null)
                    {
                        Clipboard.SetImage(image);
                    }

                    break;
                }
                case ClipboardContentType.FileList:
                    if (selected.FilePaths is { Count: > 0 })
                    {
                        var files = new StringCollection();
                        foreach (var file in selected.FilePaths)
                        {
                            files.Add(file);
                        }

                        Clipboard.SetFileDropList(files);
                    }

                    break;
            }

            UpdateStatus("已写回系统剪贴板。");
            HideWithAnimation();
        }
        finally
        {
            _ignoreClipboardUpdate = false;
        }
    }

    private void BeginEdit_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryListBox.SelectedItem is not ClipboardEntry selected || !selected.IsEditable)
        {
            return;
        }

        _isEditorDialogOpen = true;
        try
        {
            var editorDialog = new EditorDialogWindow(selected.TextContent ?? string.Empty, ResolveDarkMode(_settings.ThemeMode))
            {
                Owner = this
            };

            var result = editorDialog.ShowDialog();
            if (result == true)
            {
                selected.TextContent = editorDialog.EditedText;
                selected.CreatedAtUtc = DateTime.UtcNow;
                PersistHistory();
                RefreshHistoryView();
                HistoryListBox.SelectedItem = selected;
                UpdateStatus("文本已更新。");
            }
        }
        finally
        {
            _isEditorDialogOpen = false;
        }
    }

    private void SaveEdit_Click(object sender, RoutedEventArgs e)
    {
        if (_editingEntry is null)
        {
            return;
        }

        _editingEntry.TextContent = EditTextBox.Text ?? string.Empty;
        _editingEntry.CreatedAtUtc = DateTime.UtcNow;
        EditorCard.Visibility = Visibility.Collapsed;
        PersistHistory();
        RefreshHistoryView();
        HistoryListBox.SelectedItem = _editingEntry;
        _editingEntry = null;
        UpdateStatus("文本已更新。");
    }

    private void CancelEdit_Click(object sender, RoutedEventArgs e)
    {
        EditorCard.Visibility = Visibility.Collapsed;
        _editingEntry = null;
    }

    private void TogglePin_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryListBox.SelectedItem is not ClipboardEntry selected)
        {
            return;
        }

        selected.IsPinned = !selected.IsPinned;
        PersistHistory();
        RefreshHistoryView();
        HistoryListBox.SelectedItem = selected;
        UpdateStatus(selected.IsPinned ? "条目已置顶。" : "已取消置顶。");
    }

    private void DeleteSelected_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryListBox.SelectedItem is not ClipboardEntry selected)
        {
            return;
        }

        _history.Remove(selected);
        PersistHistory();
        RefreshHistoryView();
        UpdateStatus($"已删除 1 条，剩余 {_history.Count} 条。");
    }

    private void ClearHistory_Click(object sender, RoutedEventArgs e)
    {
        if (_history.Count == 0)
        {
            return;
        }

        var result = MessageBox.Show(
            "确认清空所有历史记录？",
            "清空确认",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        _history.Clear();
        PersistHistory();
        RefreshHistoryView();
        UpdateStatus("历史记录已清空。");
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _searchQuery = SearchTextBox.Text ?? string.Empty;
        RefreshHistoryView();
        UpdateStatus($"筛选后显示 {_historyView.Cast<object>().Count()} 条。");
    }

    private void EditTextBox_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        _editTextScrollViewer ??= FindVisualChild<ScrollViewer>(EditTextBox);
        if (_editTextScrollViewer is null || _editTextScrollViewer.ScrollableHeight <= 0)
        {
            return;
        }

        var targetOffset = _editTextScrollViewer.VerticalOffset - (e.Delta / 120.0) * 36;
        targetOffset = Math.Clamp(targetOffset, 0, _editTextScrollViewer.ScrollableHeight);
        _editTextScrollViewer.ScrollToVerticalOffset(targetOffset);
        e.Handled = true;
    }

    private void OpenSettings_Click(object sender, RoutedEventArgs e)
    {
        ApplySettingsToControls();
        SettingsOverlay.Visibility = Visibility.Visible;
    }

    private void CloseSettings_Click(object sender, RoutedEventArgs e)
    {
        SettingsOverlay.Visibility = Visibility.Collapsed;
    }

    private void RetentionRadio_Checked(object sender, RoutedEventArgs e)
    {
        if (_isInitializingUiState || sender is not RadioButton rb || rb.Tag is not string tag || !int.TryParse(tag, out var days))
        {
            return;
        }

        _settings.RetentionDays = days;
        PersistSettings();
        ApplyRetentionAndLimit();
        UpdateStatus($"保存时长已设置为 {rb.Content}。");
    }

    private void ThemeRadio_Checked(object sender, RoutedEventArgs e)
    {
        if (_isInitializingUiState || sender is not RadioButton rb || rb.Tag is not string themeKey)
        {
            return;
        }

        _settings.ThemeMode = themeKey;
        ApplyTheme(themeKey);
        PersistSettings();
        UpdateStatus($"主题已切换为 {rb.Content}。");
    }

    private void SettingsStartWithWindowsCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        if (_isInitializingUiState)
        {
            return;
        }

        SetStartupEnabled(SettingsStartWithWindowsCheckBox.IsChecked == true, persistSettings: true);
    }

    private void SetStartupEnabled(bool enabled, bool persistSettings)
    {
        _startupService.SetStartupEnabled(enabled);
        var actualEnabled = _startupService.IsStartupEnabled();
        _settings.StartWithWindows = actualEnabled;

        _isInitializingUiState = true;
        SettingsStartWithWindowsCheckBox.IsChecked = actualEnabled;
        if (_trayStartupMenuItem is not null)
        {
            _trayStartupMenuItem.Checked = actualEnabled;
        }

        _isInitializingUiState = false;

        if (persistSettings)
        {
            PersistSettings();
            UpdateStatus(actualEnabled ? "已开启开机启动。" : "已关闭开机启动。");
        }
    }

    private void ApplyTheme(string themeMode)
    {
        var isDark = ResolveDarkMode(themeMode);
        if (isDark)
        {
            ApplyBrushPalette(
                windowBackground: "#CC202020",
                card: "#D1262626",
                cardBorder: "#B34A4A4A",
                primaryText: "#F5F5F5",
                secondaryText: "#C7C7C7",
                accent: "#60CDFF",
                controlBackground: "#C0323232",
                controlBorder: "#AA505050",
                itemBackground: "#CC2F2F2F",
                dangerBackground: "#4A2A2A",
                dangerBorder: "#694040",
                itemHover: "#343434",
                itemSelected: "#1F3B5C",
                titleBar: "#2A2A2A",
                chip: "#16324D",
                chipBorder: "#24527A",
                overlay: "#A0000000",
                tabStrip: "#262626",
                tabActive: "#303030",
                tabBorder: "#454545",
                scrollThumb: "#7A7F86",
                scrollThumbHover: "#9AA0A8");
        }
        else
        {
            ApplyBrushPalette(
                windowBackground: "#D9F3F4F6",
                card: "#D9FAFCFF",
                cardBorder: "#99D8DFEA",
                primaryText: "#111827",
                secondaryText: "#667085",
                accent: "#0A5FC2",
                controlBackground: "#DDEFF4FC",
                controlBorder: "#B3C8D7EA",
                itemBackground: "#E6FFFFFF",
                dangerBackground: "#FCECEC",
                dangerBorder: "#F3BBBB",
                itemHover: "#F2F7FF",
                itemSelected: "#E8F1FF",
                titleBar: "#FFFFFF",
                chip: "#ECF4FF",
                chipBorder: "#BED6F5",
                overlay: "#8C101215",
                tabStrip: "#EDF3FC",
                tabActive: "#FFFFFF",
                tabBorder: "#C8D7EA",
                scrollThumb: "#A0A8B1",
                scrollThumbHover: "#7F8893");
        }

        ApplyWindowEffects(isDark);
    }

    private void ApplyBrushPalette(
        string windowBackground,
        string card,
        string cardBorder,
        string primaryText,
        string secondaryText,
        string accent,
        string controlBackground,
        string controlBorder,
        string itemBackground,
        string dangerBackground,
        string dangerBorder,
        string itemHover,
        string itemSelected,
        string titleBar,
        string chip,
        string chipBorder,
        string overlay,
        string tabStrip,
        string tabActive,
        string tabBorder,
        string scrollThumb,
        string scrollThumbHover)
    {
        Resources["WindowBackgroundBrush"] = CreateBrush(windowBackground);
        Resources["CardBrush"] = CreateBrush(card);
        Resources["CardBorderBrush"] = CreateBrush(cardBorder);
        Resources["PrimaryTextBrush"] = CreateBrush(primaryText);
        Resources["SecondaryTextBrush"] = CreateBrush(secondaryText);
        Resources["AccentBrush"] = CreateBrush(accent);
        Resources["ControlBackgroundBrush"] = CreateBrush(controlBackground);
        Resources["ControlBorderBrush"] = CreateBrush(controlBorder);
        Resources["ItemBackgroundBrush"] = CreateBrush(itemBackground);
        Resources["DangerBackgroundBrush"] = CreateBrush(dangerBackground);
        Resources["DangerBorderBrush"] = CreateBrush(dangerBorder);
        Resources["ItemHoverBrush"] = CreateBrush(itemHover);
        Resources["ItemSelectedBrush"] = CreateBrush(itemSelected);
        Resources["TitleBarBrush"] = CreateBrush(titleBar);
        Resources["ChipBrush"] = CreateBrush(chip);
        Resources["ChipBorderBrush"] = CreateBrush(chipBorder);
        Resources["OverlayBrush"] = CreateBrush(overlay);
        Resources["TabHeaderStripBrush"] = CreateBrush(tabStrip);
        Resources["TabHeaderActiveBrush"] = CreateBrush(tabActive);
        Resources["TabHeaderBorderBrush"] = CreateBrush(tabBorder);
        Resources["ScrollThumbBrush"] = CreateBrush(scrollThumb);
        Resources["ScrollThumbHoverBrush"] = CreateBrush(scrollThumbHover);
    }

    private static SolidColorBrush CreateBrush(string hex)
    {
        var color = (Color)ColorConverter.ConvertFromString(hex);
        var brush = new SolidColorBrush(color);
        brush.Freeze();
        return brush;
    }

    private static bool ResolveDarkMode(string themeMode)
    {
        if (string.Equals(themeMode, "Dark", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (string.Equals(themeMode, "Light", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return IsSystemDarkMode();
    }

    private static bool IsSystemDarkMode()
    {
        using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", writable: false);
        var value = key?.GetValue("AppsUseLightTheme");
        return value is int intValue && intValue == 0;
    }

    private void ApplyWindowEffects(bool useDarkMode)
    {
        if (_windowHandle == IntPtr.Zero)
        {
            return;
        }

        var dark = useDarkMode ? 1 : 0;
        var backdrop = NativeMethods.DWM_SYSTEMBACKDROP_TRANSIENTWINDOW;
        var corner = NativeMethods.DWM_WINDOW_CORNER_ROUND;

        NativeMethods.DwmSetWindowAttribute(_windowHandle, NativeMethods.DWMWA_USE_IMMERSIVE_DARK_MODE, ref dark, Marshal.SizeOf<int>());
        var result = NativeMethods.DwmSetWindowAttribute(_windowHandle, NativeMethods.DWMWA_SYSTEMBACKDROP_TYPE, ref backdrop, Marshal.SizeOf<int>());
        if (result != 0)
        {
            backdrop = NativeMethods.DWM_SYSTEMBACKDROP_MAINWINDOW;
            NativeMethods.DwmSetWindowAttribute(_windowHandle, NativeMethods.DWMWA_SYSTEMBACKDROP_TYPE, ref backdrop, Marshal.SizeOf<int>());
        }

        NativeMethods.DwmSetWindowAttribute(_windowHandle, NativeMethods.DWMWA_WINDOW_CORNER_PREFERENCE, ref corner, Marshal.SizeOf<int>());
    }

    private void InitializeTrayIcon()
    {
        if (_notifyIcon is not null)
        {
            return;
        }

        _notifyIcon = new Forms.NotifyIcon
        {
            Icon = Drawing.SystemIcons.Application,
            Text = "CaretClip - Alt+Shift+V",
            Visible = true
        };

        _notifyIcon.DoubleClick += (_, _) => Dispatcher.Invoke(ToggleWindow);

        var menu = new Forms.ContextMenuStrip();
        var toggleItem = new Forms.ToolStripMenuItem("显示/隐藏窗口");
        toggleItem.Click += (_, _) => Dispatcher.Invoke(ToggleWindow);

        var openSettingsItem = new Forms.ToolStripMenuItem("打开设置");
        openSettingsItem.Click += (_, _) => Dispatcher.Invoke(() =>
        {
            ShowAndActivateWindow();
            OpenSettings_Click(this, new RoutedEventArgs());
        });

        _trayStartupMenuItem = new Forms.ToolStripMenuItem("开机启动")
        {
            CheckOnClick = true,
            Checked = _settings.StartWithWindows
        };
        _trayStartupMenuItem.Click += (_, _) =>
            Dispatcher.Invoke(() => SetStartupEnabled(_trayStartupMenuItem.Checked, persistSettings: true));

        var exitItem = new Forms.ToolStripMenuItem("退出");
        exitItem.Click += (_, _) => Dispatcher.Invoke(ExitApplication);

        menu.Items.Add(toggleItem);
        menu.Items.Add(openSettingsItem);
        menu.Items.Add(_trayStartupMenuItem);
        menu.Items.Add(new Forms.ToolStripSeparator());
        menu.Items.Add(exitItem);
        _notifyIcon.ContextMenuStrip = menu;
    }

    private void DisposeTrayIcon()
    {
        if (_notifyIcon is null)
        {
            return;
        }

        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _notifyIcon = null;
        _trayStartupMenuItem = null;
    }

    private void CleanupNativeHooks()
    {
        if (_windowHandle != IntPtr.Zero)
        {
            if (_hotKeyRegistered)
            {
                NativeMethods.UnregisterHotKey(_windowHandle, HotKeyId);
                _hotKeyRegistered = false;
            }

            if (_clipboardListenerRegistered)
            {
                NativeMethods.RemoveClipboardFormatListener(_windowHandle);
                _clipboardListenerRegistered = false;
            }
        }

        _hwndSource?.RemoveHook(WndProc);
    }

    private void HideWindow_Click(object sender, RoutedEventArgs e)
    {
        HideWithAnimation();
    }

    private void ExitApp_Click(object sender, RoutedEventArgs e)
    {
        ExitApplication();
    }

    private void ExitApplication()
    {
        _isExitRequested = true;
        Close();
    }
}
