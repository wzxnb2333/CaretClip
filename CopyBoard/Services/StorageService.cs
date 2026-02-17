using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using CopyBoard.Models;

namespace CopyBoard.Services;

public sealed class StorageService
{
    private readonly string _appFolder;
    private readonly string _historyFilePath;
    private readonly string _settingsFilePath;
    private readonly JsonSerializerOptions _jsonOptions = new() { WriteIndented = true };

    public StorageService()
    {
        _appFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "CaretClip");
        _historyFilePath = Path.Combine(_appFolder, "history.json");
        _settingsFilePath = Path.Combine(_appFolder, "settings.json");
        Directory.CreateDirectory(_appFolder);
    }

    public List<ClipboardEntry> LoadHistory()
    {
        if (!File.Exists(_historyFilePath))
        {
            return new List<ClipboardEntry>();
        }

        try
        {
            var json = File.ReadAllText(_historyFilePath);
            return JsonSerializer.Deserialize<List<ClipboardEntry>>(json, _jsonOptions) ?? new List<ClipboardEntry>();
        }
        catch
        {
            return new List<ClipboardEntry>();
        }
    }

    public void SaveHistory(IReadOnlyCollection<ClipboardEntry> items)
    {
        var json = JsonSerializer.Serialize(items, _jsonOptions);
        File.WriteAllText(_historyFilePath, json);
    }

    public AppSettings LoadSettings()
    {
        if (!File.Exists(_settingsFilePath))
        {
            return new AppSettings();
        }

        try
        {
            var json = File.ReadAllText(_settingsFilePath);
            return JsonSerializer.Deserialize<AppSettings>(json, _jsonOptions) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public void SaveSettings(AppSettings settings)
    {
        var json = JsonSerializer.Serialize(settings, _jsonOptions);
        File.WriteAllText(_settingsFilePath, json);
    }
}
