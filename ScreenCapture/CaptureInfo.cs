namespace ScreenCapture;

public class CaptureInfo
{
    public string ProcessName { get; set; } = "";
    public string WindowTitle { get; set; } = "";
    public string? Url { get; set; }
    public string? FilePath { get; set; }
    public string? FolderPath { get; set; }
    public string? ExePath { get; set; }
    public string? ClipboardText { get; set; }
    public string? OcrText { get; set; }
    public string? CapturedImagePath { get; set; }
    public Rectangle CapturedRegion { get; set; }
    public IntPtr SourceHwnd { get; set; }
    public string? SourceKind { get; set; }
    public string? SourceAnchor { get; set; }
    public List<CaptureSource> Sources { get; set; } = new();

    public static List<CaptureSource> CloneSources(IEnumerable<CaptureSource> sources) =>
        sources.Select(s => new CaptureSource
        {
            Hwnd = s.Hwnd,
            ProcessName = s.ProcessName,
            WindowTitle = s.WindowTitle,
            Url = s.Url,
            FilePath = s.FilePath,
            FolderPath = s.FolderPath,
            ExePath = s.ExePath,
            WindowBounds = s.WindowBounds,
            CaptureBounds = s.CaptureBounds,
            SourceKind = s.SourceKind,
            SourceAnchor = s.SourceAnchor
        }).ToList();
}

public class CaptureSource
{
    public IntPtr Hwnd { get; set; }
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
}
