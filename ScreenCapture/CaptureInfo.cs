namespace ScreenCapture;

public class CaptureInfo
{
    public string ProcessName { get; set; } = "";
    public string WindowTitle { get; set; } = "";
    public string? Url { get; set; }
    public string? FilePath { get; set; }
    public string? ClipboardText { get; set; }
    public string? OcrText { get; set; }
    public string? CapturedImagePath { get; set; }
    public Rectangle CapturedRegion { get; set; }
    public IntPtr SourceHwnd { get; set; }
    public string? SourceKind { get; set; }
    public string? SourceAnchor { get; set; }
}
