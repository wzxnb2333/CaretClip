using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace CopyBoard.Models;

public enum ClipboardContentType
{
    Text = 0,
    Image = 1,
    FileList = 2
}

public sealed class ClipboardEntry
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public ClipboardContentType ContentType { get; set; }

    public string? TextContent { get; set; }

    public string? ImageBase64 { get; set; }

    public List<string>? FilePaths { get; set; }

    public DateTime CreatedAtUtc { get; set; } = DateTime.UtcNow;

    public bool IsPinned { get; set; }

    [JsonIgnore]
    public bool IsEditable => ContentType == ClipboardContentType.Text;

    [JsonIgnore]
    public string DisplayType => ContentType switch
    {
        ClipboardContentType.Text => "文本",
        ClipboardContentType.Image => "图片",
        ClipboardContentType.FileList => "文件",
        _ => "未知"
    };

    [JsonIgnore]
    public string PreviewText
    {
        get
        {
            if (ContentType == ClipboardContentType.Text)
            {
                var text = TextContent ?? string.Empty;
                text = text.Replace("\r\n", " ").Replace('\n', ' ').Trim();
                return text.Length > 100 ? text[..100] + "..." : text;
            }

            if (ContentType == ClipboardContentType.Image)
            {
                return "图片内容";
            }

            if (ContentType == ClipboardContentType.FileList)
            {
                if (FilePaths is null || FilePaths.Count == 0)
                {
                    return "文件列表为空";
                }

                var first = FilePaths[0];
                return FilePaths.Count == 1 ? first : $"{first} 等 {FilePaths.Count} 项";
            }

            return string.Empty;
        }
    }

    [JsonIgnore]
    public string CreatedAtLocalText => CreatedAtUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");

    public bool ContentEquals(ClipboardEntry other)
    {
        if (other is null || other.ContentType != ContentType)
        {
            return false;
        }

        return ContentType switch
        {
            ClipboardContentType.Text => string.Equals(TextContent, other.TextContent, StringComparison.Ordinal),
            ClipboardContentType.Image => string.Equals(ImageBase64, other.ImageBase64, StringComparison.Ordinal),
            ClipboardContentType.FileList => (FilePaths ?? new List<string>()).SequenceEqual(other.FilePaths ?? new List<string>(), StringComparer.OrdinalIgnoreCase),
            _ => false
        };
    }
}
