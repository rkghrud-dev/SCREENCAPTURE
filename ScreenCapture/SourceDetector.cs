using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;

namespace ScreenCapture;

public static class SourceDetector
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int maxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    private static readonly string[] Browsers =
        { "chrome", "msedge", "firefox", "brave", "opera", "whale", "vivaldi" };

    private static readonly string[] OfficeProcesses =
        { "excel", "winword", "powerpnt", "hwp", "hwpviewer" };

    public record WindowSnapshot(IntPtr Hwnd, string ProcessName, string Title, string? ExePath);

    public static WindowSnapshot SnapshotForeground()
    {
        var hwnd = GetForegroundWindow();

        var sb = new StringBuilder(512);
        GetWindowText(hwnd, sb, 512);
        var title = sb.ToString();

        string processName = "";
        string? exePath = null;
        GetWindowThreadProcessId(hwnd, out uint pid);
        try
        {
            var proc = Process.GetProcessById((int)pid);
            processName = proc.ProcessName;
            try { exePath = proc.MainModule?.FileName; } catch { }
        }
        catch { }

        return new WindowSnapshot(hwnd, processName, title, exePath);
    }

    public static CaptureInfo Analyze(WindowSnapshot snapshot)
    {
        var info = new CaptureInfo
        {
            ProcessName = snapshot.ProcessName,
            WindowTitle = snapshot.Title,
            SourceHwnd = snapshot.Hwnd
        };

        var isBrowser = Array.Exists(Browsers,
            b => snapshot.ProcessName.Equals(b, StringComparison.OrdinalIgnoreCase));

        if (isBrowser)
        {
            info.Url = GetBrowserUrlSafe(snapshot.Hwnd);
        }
        else
        {
            info.FilePath = FindOfficeFilePath(snapshot) ?? FindFilePath(snapshot);
            if (info.FilePath != null)
            {
                info.SourceKind = IsOfficeLike(snapshot.ProcessName) ? "Document" : "File";
                info.SourceAnchor = Path.GetFileName(info.FilePath);
            }
        }

        try
        {
            if (Clipboard.ContainsText())
                info.ClipboardText = Clipboard.GetText();
        }
        catch { }

        return info;
    }

    private static string? GetBrowserUrlSafe(IntPtr hwnd)
    {
        try
        {
            var task = Task.Run(() => GetBrowserUrl(hwnd));
            return task.Wait(TimeSpan.FromSeconds(3)) ? task.Result : null;
        }
        catch { return null; }
    }

    private static string? GetBrowserUrl(IntPtr hwnd)
    {
        try
        {
            var element = AutomationElement.FromHandle(hwnd);

            var edits = element.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));

            foreach (AutomationElement edit in edits)
            {
                try
                {
                    if (!edit.TryGetCurrentPattern(ValuePattern.Pattern, out var obj)) continue;
                    var val = ((ValuePattern)obj).Current.Value;
                    if (string.IsNullOrWhiteSpace(val) || val.Contains(' ') || val.Length > 2048) continue;

                    if (val.StartsWith("http://") || val.StartsWith("https://") ||
                        val.StartsWith("file://") || LooksLikeDomain(val))
                    {
                        return val.StartsWith("http") || val.StartsWith("file")
                            ? val : "https://" + val;
                    }
                }
                catch { }
            }
        }
        catch { }
        return null;
    }

    private static bool LooksLikeDomain(string s) =>
        s.Length < 256 && s.Contains('.') && !s.Contains(' ') &&
        !s.StartsWith(".") && !s.EndsWith(".");

    private static string? FindFilePath(WindowSnapshot snapshot)
    {
        var separators = new[] { " - ", " \u2014 ", " \u2013 ", " | " };
        var parts = snapshot.Title.Split(separators,
            StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

        foreach (var part in parts)
        {
            var cleaned = CleanTitlePart(part);
            if (cleaned.Length < 2 || cleaned.Length > 260) continue;

            if (cleaned.Length > 3 && cleaned[1] == ':' && File.Exists(cleaned))
                return cleaned;

            foreach (var fileName in ExpandCandidateFileNames(cleaned, snapshot.ProcessName))
            {
                var found = SearchFile(fileName);
                if (found != null) return found;
            }
        }
        return null;
    }

    private static string? FindOfficeFilePath(WindowSnapshot snapshot)
    {
        var process = snapshot.ProcessName.ToLowerInvariant();

        return process switch
        {
            "excel" => GetOfficeComPath("Excel.Application", "ActiveWorkbook", "FullName"),
            "winword" => GetOfficeComPath("Word.Application", "ActiveDocument", "FullName"),
            "powerpnt" => GetOfficeComPath("PowerPoint.Application", "ActivePresentation", "FullName"),
            _ => null
        };
    }

    private static string? GetOfficeComPath(string progId, string activeObjectProperty, string pathProperty)
    {
        object? app = null;
        object? activeObject = null;

        try
        {
            CLSIDFromProgID(progId, out var clsid);
            GetActiveObject(ref clsid, IntPtr.Zero, out app);
            if (app == null) return null;

            activeObject = app.GetType().InvokeMember(activeObjectProperty,
                System.Reflection.BindingFlags.GetProperty, null, app, null);
            if (activeObject == null) return null;

            var path = activeObject.GetType().InvokeMember(pathProperty,
                System.Reflection.BindingFlags.GetProperty, null, activeObject, null) as string;
            return !string.IsNullOrWhiteSpace(path) && File.Exists(path) ? path : null;
        }
        catch
        {
            return null;
        }
        finally
        {
            ReleaseComObject(activeObject);
            ReleaseComObject(app);
        }
    }

    private static void ReleaseComObject(object? obj)
    {
        try
        {
            if (obj != null && Marshal.IsComObject(obj))
                Marshal.FinalReleaseComObject(obj);
        }
        catch { }
    }

    private static string CleanTitlePart(string part)
    {
        var cleaned = part.Trim().Trim('"');
        var bracket = cleaned.IndexOf(" [", StringComparison.Ordinal);
        if (bracket > 0)
            cleaned = cleaned[..bracket].Trim();
        return cleaned;
    }

    private static IEnumerable<string> ExpandCandidateFileNames(string candidate, string processName)
    {
        if (candidate.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            yield break;

        if (Path.HasExtension(candidate))
        {
            yield return candidate;
            yield break;
        }

        foreach (var ext in GetLikelyExtensions(processName))
            yield return candidate + ext;
    }

    private static IEnumerable<string> GetLikelyExtensions(string processName)
    {
        var process = processName.ToLowerInvariant();
        if (process == "excel")
            return new[] { ".xlsx", ".xlsm", ".xls", ".csv" };
        if (process == "winword")
            return new[] { ".docx", ".doc", ".rtf", ".txt" };
        if (process == "powerpnt")
            return new[] { ".pptx", ".pptm", ".ppt" };
        if (process is "hwp" or "hwpviewer")
            return new[] { ".hwp", ".hwpx" };

        return Array.Empty<string>();
    }

    private static string? SearchFile(string fileName)
    {
        var searchDirs = GetSearchRoots()
            .Where(Directory.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        foreach (var dir in searchDirs)
        {
            var path = Path.Combine(dir, fileName);
            if (File.Exists(path)) return path;
        }

        foreach (var dir in searchDirs)
        {
            var found = SearchFileRecursive(dir, fileName, depth: 3);
            if (found != null) return found;
        }
        return null;
    }

    private static IEnumerable<string> GetSearchRoots()
    {
        var user = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        yield return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        yield return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        yield return Path.Combine(user, "Downloads");
        yield return Path.Combine(user, "OneDrive");
        yield return Path.Combine(user, "OneDrive - Personal");

        var oneDrive = Environment.GetEnvironmentVariable("OneDrive");
        if (!string.IsNullOrWhiteSpace(oneDrive)) yield return oneDrive;

        var commercial = Environment.GetEnvironmentVariable("OneDriveCommercial");
        if (!string.IsNullOrWhiteSpace(commercial)) yield return commercial;
    }

    private static string? SearchFileRecursive(string dir, string fileName, int depth)
    {
        if (depth < 0) return null;

        try
        {
            foreach (var file in Directory.EnumerateFiles(dir, fileName))
                return file;

            foreach (var subDir in Directory.EnumerateDirectories(dir))
            {
                var found = SearchFileRecursive(subDir, fileName, depth - 1);
                if (found != null) return found;
            }
        }
        catch { }

        return null;
    }

    private static bool IsOfficeLike(string processName) =>
        OfficeProcesses.Contains(processName.ToLowerInvariant());
}
