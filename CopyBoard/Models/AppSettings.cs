namespace CopyBoard.Models;

public sealed class AppSettings
{
    public int RetentionDays { get; set; } = 30;

    public int MaxHistoryItems { get; set; } = 1000;

    public string ThemeMode { get; set; } = "System";

    public bool StartWithWindows { get; set; }
}
