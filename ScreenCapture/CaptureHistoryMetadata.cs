using System.Text.Json;

namespace ScreenCapture;

public sealed class CaptureHistoryMetadata
{
    public DateTime CapturedAt { get; set; } = DateTime.Now;
    public string ProcessName { get; set; } = "";
    public string WindowTitle { get; set; } = "";
    public string? Url { get; set; }
    public string? FilePath { get; set; }
    public string? FolderPath { get; set; }
    public string? ExePath { get; set; }
    public string? ClipboardText { get; set; }
    public string? OcrText { get; set; }
    public Rectangle CapturedRegion { get; set; }
    public string? SourceKind { get; set; }
    public string? SourceAnchor { get; set; }
    public List<CaptureSourceMetadata> Sources { get; set; } = new();

    public static string GetPath(string imagePath) => imagePath + ".json";

    public static CaptureHistoryMetadata FromCaptureInfo(CaptureInfo info) =>
        new()
        {
            CapturedAt = DateTime.Now,
            ProcessName = info.ProcessName,
            WindowTitle = info.WindowTitle,
            Url = info.Url,
            FilePath = info.FilePath,
            FolderPath = info.FolderPath,
            ExePath = info.ExePath,
            ClipboardText = info.ClipboardText,
            OcrText = info.OcrText,
            CapturedRegion = info.CapturedRegion,
            SourceKind = info.SourceKind,
            SourceAnchor = info.SourceAnchor,
            Sources = info.Sources.Select(CaptureSourceMetadata.FromCaptureSource).ToList()
        };

    public static void Save(string imagePath, CaptureInfo info)
    {
        var metadata = FromCaptureInfo(info);
        var json = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(GetPath(imagePath), json);
    }

    public static CaptureInfo LoadInfo(string imagePath)
    {
        var metadataPath = GetPath(imagePath);
        if (!File.Exists(metadataPath))
            return CreateFallbackInfo(imagePath);

        try
        {
            var json = File.ReadAllText(metadataPath);
            var metadata = JsonSerializer.Deserialize<CaptureHistoryMetadata>(json);
            if (metadata == null)
                return CreateFallbackInfo(imagePath);

            return new CaptureInfo
            {
                ProcessName = metadata.ProcessName,
                WindowTitle = string.IsNullOrWhiteSpace(metadata.WindowTitle)
                    ? Path.GetFileName(imagePath)
                    : metadata.WindowTitle,
                Url = metadata.Url,
                FilePath = metadata.FilePath,
                FolderPath = metadata.FolderPath,
                ExePath = metadata.ExePath,
                ClipboardText = metadata.ClipboardText,
                OcrText = metadata.OcrText,
                CapturedImagePath = imagePath,
                CapturedRegion = metadata.CapturedRegion,
                SourceKind = metadata.SourceKind,
                SourceAnchor = metadata.SourceAnchor,
                Sources = metadata.Sources.Select(s => s.ToCaptureSource()).ToList()
            };
        }
        catch
        {
            return CreateFallbackInfo(imagePath);
        }
    }

    private static CaptureInfo CreateFallbackInfo(string imagePath) =>
        new()
        {
            ProcessName = "보관함",
            WindowTitle = Path.GetFileName(imagePath),
            CapturedImagePath = imagePath,
            FilePath = imagePath,
            SourceKind = "File",
            SourceAnchor = Path.GetFileName(imagePath)
        };
}

public sealed class CaptureSourceMetadata
{
    public string ProcessName { get; set; } = "";
    public string WindowTitle { get; set; } = "";
    public string? Url { get; set; }
    public string? FilePath { get; set; }
    public string? FolderPath { get; set; }
    public string? ExePath { get; set; }
    public Rectangle WindowBounds { get; set; }
    public Rectangle CaptureBounds { get; set; }
    public string? SourceKind { get; set; }
    public string? SourceAnchor { get; set; }

    public static CaptureSourceMetadata FromCaptureSource(CaptureSource source) =>
        new()
        {
            ProcessName = source.ProcessName,
            WindowTitle = source.WindowTitle,
            Url = source.Url,
            FilePath = source.FilePath,
            FolderPath = source.FolderPath,
            ExePath = source.ExePath,
            WindowBounds = source.WindowBounds,
            CaptureBounds = source.CaptureBounds,
            SourceKind = source.SourceKind,
            SourceAnchor = source.SourceAnchor
        };

    public CaptureSource ToCaptureSource() =>
        new()
        {
            ProcessName = ProcessName,
            WindowTitle = WindowTitle,
            Url = Url,
            FilePath = FilePath,
            FolderPath = FolderPath,
            ExePath = ExePath,
            WindowBounds = WindowBounds,
            CaptureBounds = CaptureBounds,
            SourceKind = SourceKind,
            SourceAnchor = SourceAnchor
        };
}
