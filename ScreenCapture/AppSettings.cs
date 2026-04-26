using System.Text.Json;

namespace ScreenCapture;

public enum CaptureAction
{
    ClipboardOnly = 0,      // 클립보드만 (이미지)
    FileAndClipboard = 1,   // 파일 저장 + 경로 클립보드
    FileOnly = 2            // 파일만 저장
}

public class HotkeyConfig
{
    public bool Ctrl { get; set; } = true;
    public bool Shift { get; set; } = true;
    public bool Alt { get; set; }
    public string Key { get; set; } = "A";

    public int GetModifiers()
    {
        int m = 0;
        if (Ctrl) m |= HotkeyWindow.MOD_CTRL;
        if (Shift) m |= HotkeyWindow.MOD_SHIFT;
        if (Alt) m |= HotkeyWindow.MOD_ALT;
        return m;
    }

    public Keys GetKey()
    {
        if (Enum.TryParse<Keys>(Key, true, out var k)) return k;
        return Keys.None;
    }

    public string ToDisplayString()
    {
        var parts = new List<string>();
        if (Ctrl) parts.Add("Ctrl");
        if (Shift) parts.Add("Shift");
        if (Alt) parts.Add("Alt");
        parts.Add(Key);
        return string.Join("+", parts);
    }
}

public class AppSettings
{
    // --- General ---
    public bool StartWithWindows { get; set; } = true;
    public bool StartMinimized { get; set; } = true;

    // --- Capture ---
    public bool IncludeCursor { get; set; }
    public int CaptureDelay { get; set; }  // ms, 0 = no delay
    public bool PlaySound { get; set; } = true;
    public bool PinToDesktop { get; set; }
    public bool EnableOcr { get; set; } = true;

    // --- Save ---
    public string SaveFolder { get; set; } =
        Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
    public bool AutoSaveCaptures { get; set; } = true;
    public string ImageFormat { get; set; } = "JPG";   // JPG, PNG, BMP
    public int JpegQuality { get; set; } = 95;         // 1-100
    public string FileNamePattern { get; set; } = "yyyyMMddHHmmss";  // DateTime format

    // --- Hotkeys ---
    public HotkeyConfig HotkeyA { get; set; } = new() { Ctrl = true, Shift = true, Alt = false, Key = "A" };
    public CaptureAction ActionA { get; set; } = CaptureAction.ClipboardOnly;
    public HotkeyConfig HotkeyB { get; set; } = new() { Ctrl = true, Shift = true, Alt = false, Key = "B" };
    public CaptureAction ActionB { get; set; } = CaptureAction.FileAndClipboard;
    public HotkeyConfig HotkeyC { get; set; } = new() { Ctrl = true, Shift = true, Alt = false, Key = "S" };
    public CaptureAction ActionC { get; set; } = CaptureAction.FileAndClipboard;

    // --- Persistence ---

    private static string ConfigPath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ScreenCapture", "settings.json");

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath);
                return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            }
        }
        catch { }
        return new AppSettings();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ConfigPath, json);
        }
        catch { }
    }

    public AppSettings Clone()
    {
        var json = JsonSerializer.Serialize(this);
        return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
    }

    public string GetFileExtension() => ImageFormat.ToUpper() switch
    {
        "PNG" => ".png",
        "BMP" => ".bmp",
        _ => ".jpg"
    };

    public string GetHistoryRoot() =>
        Path.Combine(SaveFolder, "ScreenCapture History");
}
