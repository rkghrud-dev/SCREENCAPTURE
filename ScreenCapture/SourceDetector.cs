using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;

namespace ScreenCapture;

public static class SourceDetector
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern IntPtr GetShellWindow();

    [DllImport("dwmapi.dll")]
    private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out int pvAttribute, int cbAttribute);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int maxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    private static readonly string[] Browsers =
        { "chrome", "msedge", "firefox", "brave", "opera", "whale", "vivaldi" };

    private static readonly string[] OfficeProcesses =
        { "excel", "winword", "powerpnt", "hwp", "hwpviewer" };

    private const int ForegroundBrowserUrlTimeoutMs = 3000;
    private const int BackgroundBrowserUrlTimeoutMs = 900;
    private const int MaxBackgroundBrowserUrlLookups = 4;

    public record WindowSnapshot(
        IntPtr Hwnd,
        string ProcessName,
        string Title,
        string? ExePath,
        Rectangle WindowBounds);

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

        return new WindowSnapshot(hwnd, processName, title, exePath, GetWindowBounds(hwnd));
    }

    public static List<WindowSnapshot> SnapshotVisibleWindows()
    {
        var result = new List<WindowSnapshot>();
        var shellWindow = GetShellWindow();
        var currentPid = Environment.ProcessId;

        EnumWindows((hwnd, _) =>
        {
            if (hwnd == IntPtr.Zero || hwnd == shellWindow || !IsWindowVisible(hwnd))
                return true;

            if (IsCloaked(hwnd))
                return true;

            var bounds = GetWindowBounds(hwnd);
            if (bounds.Width < 30 || bounds.Height < 30)
                return true;

            var sb = new StringBuilder(512);
            GetWindowText(hwnd, sb, 512);
            var title = sb.ToString();

            string processName = "";
            string? exePath = null;
            GetWindowThreadProcessId(hwnd, out uint pid);
            if (pid == currentPid)
                return true;

            try
            {
                var proc = Process.GetProcessById((int)pid);
                processName = proc.ProcessName;
                try { exePath = proc.MainModule?.FileName; } catch { }
            }
            catch { }

            if (string.IsNullOrWhiteSpace(title) && string.IsNullOrWhiteSpace(processName))
                return true;

            result.Add(new WindowSnapshot(hwnd, processName, title, exePath, bounds));
            return true;
        }, IntPtr.Zero);

        return result;
    }

    public static CaptureInfo Analyze(WindowSnapshot snapshot)
    {
        var source = AnalyzeSource(snapshot, Rectangle.Empty, includeDeepMetadata: true);
        var info = ToCaptureInfo(source);

        try
        {
            if (Clipboard.ContainsText())
                info.ClipboardText = Clipboard.GetText();
        }
        catch { }

        return info;
    }

    public static CaptureInfo Analyze(
        WindowSnapshot foreground,
        IEnumerable<WindowSnapshot> visibleWindows,
        Rectangle capturedRegion)
    {
        var info = Analyze(foreground);
        info.Sources = AnalyzeSources(visibleWindows, capturedRegion, foreground.Hwnd);

        var primarySource = info.Sources.FirstOrDefault(s => s.Hwnd == foreground.Hwnd)
                            ?? info.Sources.FirstOrDefault();
        if (primarySource != null)
            ApplyPrimarySource(info, primarySource);

        return info;
    }

    public static List<CaptureSource> AnalyzeSources(
        IEnumerable<WindowSnapshot> visibleWindows,
        Rectangle capturedRegion,
        IntPtr foregroundHwnd)
    {
        if (capturedRegion.Width <= 0 || capturedRegion.Height <= 0)
            return new List<CaptureSource>();

        var sources = new List<CaptureSource>();
        var backgroundBrowserUrlLookups = 0;
        foreach (var snapshot in visibleWindows)
        {
            var intersection = Rectangle.Intersect(snapshot.WindowBounds, capturedRegion);
            if (intersection.Width < 6 || intersection.Height < 6)
                continue;

            var isForeground = snapshot.Hwnd == foregroundHwnd;
            var isBrowser = IsBrowserLike(snapshot.ProcessName);
            var shouldReadMetadata = isForeground || IsExplorerLike(snapshot.ProcessName);
            var browserUrlTimeoutMs = ForegroundBrowserUrlTimeoutMs;

            if (isBrowser && !isForeground && backgroundBrowserUrlLookups < MaxBackgroundBrowserUrlLookups)
            {
                shouldReadMetadata = true;
                browserUrlTimeoutMs = BackgroundBrowserUrlTimeoutMs;
                backgroundBrowserUrlLookups++;
            }

            var source = AnalyzeSource(
                snapshot,
                capturedRegion,
                shouldReadMetadata,
                browserUrlTimeoutMs);
            source.CaptureBounds = new Rectangle(
                intersection.X - capturedRegion.X,
                intersection.Y - capturedRegion.Y,
                intersection.Width,
                intersection.Height);
            sources.Add(source);
        }

        return sources;
    }

    private static CaptureSource AnalyzeSource(
        WindowSnapshot snapshot,
        Rectangle capturedRegion,
        bool includeDeepMetadata,
        int browserUrlTimeoutMs = ForegroundBrowserUrlTimeoutMs)
    {
        var source = new CaptureSource
        {
            Hwnd = snapshot.Hwnd,
            ProcessName = snapshot.ProcessName,
            WindowTitle = snapshot.Title,
            ExePath = snapshot.ExePath,
            WindowBounds = snapshot.WindowBounds
        };

        if (capturedRegion.Width > 0 && capturedRegion.Height > 0)
        {
            var intersection = Rectangle.Intersect(snapshot.WindowBounds, capturedRegion);
            source.CaptureBounds = new Rectangle(
                intersection.X - capturedRegion.X,
                intersection.Y - capturedRegion.Y,
                intersection.Width,
                intersection.Height);
        }

        var isBrowser = IsBrowserLike(snapshot.ProcessName);

        if (isBrowser && includeDeepMetadata)
        {
            source.Url = GetBrowserUrlSafe(snapshot.Hwnd, browserUrlTimeoutMs);
            if (!string.IsNullOrEmpty(source.Url))
            {
                source.SourceKind = "Url";
                source.SourceAnchor = source.Url;
            }
        }
        else if (IsExplorerLike(snapshot.ProcessName))
        {
            source.FolderPath = FindExplorerFolderPath(snapshot);
            if (!string.IsNullOrWhiteSpace(source.FolderPath))
            {
                source.SourceKind = "Folder";
                source.SourceAnchor = source.FolderPath;
            }
        }
        else
        {
            source.FilePath = FindOfficeFilePath(snapshot) ?? FindFilePath(snapshot);
            if (source.FilePath != null)
            {
                source.SourceKind = IsOfficeLike(snapshot.ProcessName) ? "Document" : "File";
                source.SourceAnchor = Path.GetFileName(source.FilePath);
            }
        }

        return source;
    }

    private static CaptureInfo ToCaptureInfo(CaptureSource source) =>
        new()
        {
            ProcessName = source.ProcessName,
            WindowTitle = source.WindowTitle,
            Url = source.Url,
            FilePath = source.FilePath,
            FolderPath = source.FolderPath,
            ExePath = source.ExePath,
            SourceHwnd = source.Hwnd,
            SourceKind = source.SourceKind,
            SourceAnchor = source.SourceAnchor
        };

    private static void ApplyPrimarySource(CaptureInfo info, CaptureSource source)
    {
        info.ProcessName = source.ProcessName;
        info.WindowTitle = source.WindowTitle;
        info.Url = source.Url;
        info.FilePath = source.FilePath;
        info.FolderPath = source.FolderPath;
        info.ExePath = source.ExePath;
        info.SourceHwnd = source.Hwnd;
        info.SourceKind = source.SourceKind;
        info.SourceAnchor = source.SourceAnchor;
    }

    private static Rectangle GetWindowBounds(IntPtr hwnd)
    {
        if (!GetWindowRect(hwnd, out var rect))
            return Rectangle.Empty;

        return Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom);
    }

    private static bool IsCloaked(IntPtr hwnd)
    {
        try
        {
            return DwmGetWindowAttribute(hwnd, 14, out var cloaked, Marshal.SizeOf<int>()) == 0 &&
                   cloaked != 0;
        }
        catch
        {
            return false;
        }
    }

    private static string? GetBrowserUrlSafe(IntPtr hwnd, int timeoutMs = ForegroundBrowserUrlTimeoutMs)
    {
        try
        {
            var task = Task.Run(() => GetBrowserUrl(hwnd));
            return task.Wait(TimeSpan.FromMilliseconds(timeoutMs)) ? task.Result : null;
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

    private static string? FindExplorerFolderPath(WindowSnapshot snapshot)
    {
        if (!IsExplorerLike(snapshot.ProcessName))
            return null;

        object? shell = null;
        object? windows = null;

        try
        {
            var shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType == null) return null;

            shell = Activator.CreateInstance(shellType);
            if (shell == null) return null;

            windows = shell.GetType().InvokeMember("Windows",
                System.Reflection.BindingFlags.InvokeMethod, null, shell, null);
            if (windows == null) return null;

            var countObj = windows.GetType().InvokeMember("Count",
                System.Reflection.BindingFlags.GetProperty, null, windows, null);
            var count = Convert.ToInt32(countObj);

            for (var i = 0; i < count; i++)
            {
                object? window = null;
                try
                {
                    window = windows.GetType().InvokeMember("Item",
                        System.Reflection.BindingFlags.InvokeMethod, null, windows, new object[] { i });
                    if (window == null) continue;

                    var hwndObj = window.GetType().InvokeMember("HWND",
                        System.Reflection.BindingFlags.GetProperty, null, window, null);
                    if (!HwndMatches(hwndObj, snapshot.Hwnd))
                        continue;

                    var locationUrl = window.GetType().InvokeMember("LocationURL",
                        System.Reflection.BindingFlags.GetProperty, null, window, null) as string;
                    var path = FileUrlToPath(locationUrl);
                    if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                        return path;

                    path = TryGetShellFolderPath(window);
                    if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                        return path;
                }
                catch { }
                finally
                {
                    ReleaseComObject(window);
                }
            }
        }
        catch
        {
            return null;
        }
        finally
        {
            ReleaseComObject(windows);
            ReleaseComObject(shell);
        }

        return null;
    }

    private static bool HwndMatches(object? candidate, IntPtr expected)
    {
        if (candidate == null)
            return false;

        try
        {
            var candidateValue = Convert.ToInt64(candidate);
            var expectedValue = expected.ToInt64();
            if (candidateValue == expectedValue)
                return true;

            return unchecked((uint)candidateValue) == unchecked((uint)expectedValue);
        }
        catch
        {
            return false;
        }
    }

    private static string? FileUrlToPath(string? url)
    {
        if (string.IsNullOrWhiteSpace(url) ||
            !url.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
            return null;

        try
        {
            return new Uri(url).LocalPath;
        }
        catch
        {
            return null;
        }
    }

    private static string? TryGetShellFolderPath(object window)
    {
        object? document = null;
        object? folder = null;
        object? self = null;

        try
        {
            document = window.GetType().InvokeMember("Document",
                System.Reflection.BindingFlags.GetProperty, null, window, null);
            if (document == null) return null;

            folder = document.GetType().InvokeMember("Folder",
                System.Reflection.BindingFlags.GetProperty, null, document, null);
            if (folder == null) return null;

            self = folder.GetType().InvokeMember("Self",
                System.Reflection.BindingFlags.GetProperty, null, folder, null);
            if (self == null) return null;

            return self.GetType().InvokeMember("Path",
                System.Reflection.BindingFlags.GetProperty, null, self, null) as string;
        }
        catch
        {
            return null;
        }
        finally
        {
            ReleaseComObject(self);
            ReleaseComObject(folder);
            ReleaseComObject(document);
        }
    }

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
        if (IsImageViewerLike(process))
            return new[] { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff", ".heic" };

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

    private static bool IsExplorerLike(string processName) =>
        processName.Equals("explorer", StringComparison.OrdinalIgnoreCase);

    private static bool IsBrowserLike(string processName) =>
        Array.Exists(Browsers, b => processName.Equals(b, StringComparison.OrdinalIgnoreCase));

    private static bool IsImageViewerLike(string processName)
    {
        var process = processName.ToLowerInvariant();
        return process is "photos" or "photoviewer" ||
               process.Contains("photo") || process.Contains("image");
    }
}
