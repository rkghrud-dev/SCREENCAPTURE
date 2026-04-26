using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace ScreenCapture;

public class StickyWindow : Form
{
    private const int Border = 3;
    private const int GripSize = 8;

    private Bitmap _originalImage;
    private readonly CaptureInfo _info;
    private readonly string _saveFolder;

    public event Action<Bitmap, CaptureInfo>? NewStickyRequested;

    private bool _dragging;
    private bool _scaling;
    private Point _dragOffset;
    private Point _scaleStartScreen;
    private Size _scaleStartSize;
    private Point _scaleStartLocation;
    private bool _showOpacity;
    private bool _showZoom;
    private string _zoomText = "";
    private readonly System.Windows.Forms.Timer _opacityTimer;
    private readonly System.Windows.Forms.Timer _zoomTimer;
    private ToolStripMenuItem? _topMostItem;

    private const int WM_NCHITTEST = 0x0084;
    private const int HTLEFT = 10;
    private const int HTRIGHT = 11;
    private const int HTTOP = 12;
    private const int HTTOPLEFT = 13;
    private const int HTTOPRIGHT = 14;
    private const int HTBOTTOM = 15;
    private const int HTBOTTOMLEFT = 16;
    private const int HTBOTTOMRIGHT = 17;
    private const int SW_RESTORE = 9;
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

    public StickyWindow(Bitmap image, CaptureInfo info, string? saveFolder = null)
    {
        _info = info;
        _saveFolder = saveFolder ?? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        _originalImage = new Bitmap(image);

        var scale = Math.Min(500.0 / image.Width, 400.0 / image.Height);
        if (scale > 1) scale = 1;
        var w = (int)(image.Width * scale);
        var h = (int)(image.Height * scale);

        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.CenterScreen;
        TopMost = true;
        ShowInTaskbar = false;
        DoubleBuffered = true;
        KeyPreview = true;
        MinimumSize = new Size(60, 60);
        Size = new Size(w + Border * 2, h + Border * 2);
        Opacity = 0.9;

        ContextMenuStrip = BuildMenu();

        _opacityTimer = new System.Windows.Forms.Timer { Interval = 1000 };
        _opacityTimer.Tick += (_, _) =>
        {
            _showOpacity = false;
            _opacityTimer.Stop();
            Invalidate();
        };

        _zoomTimer = new System.Windows.Forms.Timer { Interval = 1000 };
        _zoomTimer.Tick += (_, _) =>
        {
            _showZoom = false;
            _zoomTimer.Stop();
            Invalidate();
        };
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ClassStyle |= 0x00020000;
            return cp;
        }
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_NCHITTEST)
        {
            base.WndProc(ref m);
            var lp = m.LParam.ToInt64();
            int sx = (short)(lp & 0xFFFF);
            int sy = (short)((lp >> 16) & 0xFFFF);
            var pt = PointToClient(new Point(sx, sy));

            if (pt.X <= GripSize && pt.Y <= GripSize) { m.Result = (IntPtr)HTTOPLEFT; return; }
            if (pt.X >= Width - GripSize && pt.Y <= GripSize) { m.Result = (IntPtr)HTTOPRIGHT; return; }
            if (pt.X <= GripSize && pt.Y >= Height - GripSize) { m.Result = (IntPtr)HTBOTTOMLEFT; return; }
            if (pt.X >= Width - GripSize && pt.Y >= Height - GripSize) { m.Result = (IntPtr)HTBOTTOMRIGHT; return; }
            if (pt.X <= GripSize) { m.Result = (IntPtr)HTLEFT; return; }
            if (pt.X >= Width - GripSize) { m.Result = (IntPtr)HTRIGHT; return; }
            if (pt.Y <= GripSize) { m.Result = (IntPtr)HTTOP; return; }
            if (pt.Y >= Height - GripSize) { m.Result = (IntPtr)HTBOTTOM; return; }
            return;
        }
        base.WndProc(ref m);
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;

        using (var borderBrush = new SolidBrush(Color.FromArgb(60, 60, 60)))
            g.FillRectangle(borderBrush, ClientRectangle);

        var destRect = new Rectangle(Border, Border,
            Width - Border * 2, Height - Border * 2);
        g.DrawImage(_originalImage, destRect);

        var tag = _info.Url != null ? "[URL]" :
            _info.FilePath != null ? "[File]" :
            _info.FolderPath != null ? "[Folder]" : "";
        var sourceCount = _info.Sources.Count > 1 ? $" +{_info.Sources.Count - 1}" : "";
        var tagText = string.IsNullOrEmpty(tag)
            ? $"{_info.ProcessName}{sourceCount}"
            : $"{tag} {_info.ProcessName}{sourceCount}";
        if (!string.IsNullOrEmpty(tagText))
        {
            using var tagFont = new Font("Segoe UI", 8);
            using var tagBg = new SolidBrush(Color.FromArgb(180, 0, 0, 0));
            using var tagFg = new SolidBrush(Color.White);
            var tagSize = g.MeasureString(tagText, tagFont);
            g.FillRectangle(tagBg, Border, Border, tagSize.Width + 6, tagSize.Height + 2);
            g.DrawString(tagText, tagFont, tagFg, Border + 3, Border + 1);
        }

        if (_showOpacity)
        {
            var opText = $"{(int)(Opacity * 100)}%";
            DrawCenterBadge(g, opText);
        }
        else if (_showZoom)
        {
            DrawCenterBadge(g, _zoomText);
        }
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        Activate();
        Focus();

        if (e.Clicks > 1)
        {
            _dragging = false;
            return;
        }

        if (e.Button == MouseButtons.Left)
        {
            if ((ModifierKeys & Keys.Shift) == Keys.Shift)
            {
                BeginScaleDrag(e.Location);
                return;
            }

            _dragging = true;
            _dragOffset = e.Location;
        }
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        if (_scaling)
        {
            UpdateScaleDrag(e.Location);
            return;
        }

        if (_dragging)
            Location = new Point(
                Location.X + e.X - _dragOffset.X,
                Location.Y + e.Y - _dragOffset.Y);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left)
            return;

        if (_scaling)
        {
            _scaling = false;
            Capture = false;
            Cursor = Cursors.Default;
            _zoomTimer.Stop();
            _zoomTimer.Start();
            return;
        }

        _dragging = false;
    }

    protected override void OnMouseDoubleClick(MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left && TryOpenSourceAt(e.Location))
            return;

        base.OnMouseDoubleClick(e);
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        if ((ModifierKeys & Keys.Shift) == Keys.Shift)
        {
            var factor = e.Delta > 0 ? 1.15 : 1.0 / 1.15;
            var newW = Math.Max(MinimumSize.Width, (int)(Width * factor));
            var newH = Math.Max(MinimumSize.Height, (int)(Height * factor));
            var cx = Location.X + Width / 2;
            var cy = Location.Y + Height / 2;
            Size = new Size(newW, newH);
            Location = new Point(cx - newW / 2, cy - newH / 2);

            var imageScale = (Width - Border * 2) * 100.0 / _originalImage.Width;
            _zoomText = $"{Math.Max(1, (int)Math.Round(imageScale))}%";
            _showZoom = true;
            _showOpacity = false;
            _zoomTimer.Stop();
            _zoomTimer.Start();
            Invalidate();
            return;
        }

        Opacity = Math.Clamp(Opacity + e.Delta / 1200.0, 0.1, 1.0);
        _showOpacity = true;
        _opacityTimer.Stop();
        _opacityTimer.Start();
        Invalidate();
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode is Keys.Escape or Keys.Delete)
        {
            Close();
            return;
        }

        base.OnKeyDown(e);
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        var keyCode = keyData & Keys.KeyCode;
        if (keyCode is Keys.Escape or Keys.Delete)
        {
            Close();
            return true;
        }

        return base.ProcessCmdKey(ref msg, keyData);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _originalImage.Dispose();
            _opacityTimer.Dispose();
            _zoomTimer.Dispose();
        }
        base.Dispose(disposing);
    }

    private void BeginScaleDrag(Point clientPoint)
    {
        _dragging = false;
        _scaling = true;
        _scaleStartScreen = PointToScreen(clientPoint);
        _scaleStartSize = Size;
        _scaleStartLocation = Location;
        Capture = true;
        Cursor = Cursors.SizeNWSE;
    }

    private void UpdateScaleDrag(Point clientPoint)
    {
        var current = PointToScreen(clientPoint);
        var delta = (current.X - _scaleStartScreen.X) + (current.Y - _scaleStartScreen.Y);
        var scale = Math.Clamp(1.0 + delta / 280.0, 0.2, 5.0);

        var newWidth = Math.Max(MinimumSize.Width, (int)Math.Round(_scaleStartSize.Width * scale));
        var newHeight = Math.Max(MinimumSize.Height, (int)Math.Round(_scaleStartSize.Height * scale));

        var centerX = _scaleStartLocation.X + _scaleStartSize.Width / 2;
        var centerY = _scaleStartLocation.Y + _scaleStartSize.Height / 2;

        Size = new Size(newWidth, newHeight);
        Location = new Point(centerX - newWidth / 2, centerY - newHeight / 2);

        var imageScale = (Width - Border * 2) * 100.0 / _originalImage.Width;
        _zoomText = $"{Math.Max(1, (int)Math.Round(imageScale))}%";
        _showZoom = true;
        _showOpacity = false;
        Invalidate();
    }

    private bool TryOpenSourceAt(Point clientPoint)
    {
        if (TryGetSourcePoint(clientPoint, out var screenPoint, out var imagePoint))
        {
            var source = FindSourceAt(imagePoint, screenPoint);
            if (source != null)
            {
                if (!string.IsNullOrEmpty(source.Url) && TryOpenSource(source))
                    return true;

                if (TryForwardClickToWindow(source.Hwnd, screenPoint))
                    return true;

                if (TryOpenSource(source))
                    return true;
            }

            if (TryForwardClickToWindow(_info.SourceHwnd, screenPoint))
                return true;
        }

        return TryOpenInfoSource();
    }

    private bool TryOpenInfoSource()
    {
        if (!string.IsNullOrEmpty(_info.Url))
        {
            return OpenUrlInCapturedBrowser(_info.Url, _info.ExePath, _info.ProcessName);
        }

        if (!string.IsNullOrEmpty(_info.FilePath) && File.Exists(_info.FilePath))
        {
            Process.Start(new ProcessStartInfo(_info.FilePath) { UseShellExecute = true });
            return true;
        }

        if (!string.IsNullOrEmpty(_info.FolderPath) && Directory.Exists(_info.FolderPath))
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{_info.FolderPath}\"") { UseShellExecute = true });
            return true;
        }

        return false;
    }

    private bool TryGetSourcePoint(Point clientPoint, out Point screenPoint, out Point imagePoint)
    {
        screenPoint = Point.Empty;
        imagePoint = Point.Empty;

        if (_info.CapturedRegion.Width <= 0 || _info.CapturedRegion.Height <= 0)
            return false;

        var imageRect = GetImageClientRect();
        if (!imageRect.Contains(clientPoint))
            return false;

        imagePoint = new Point(
            (int)Math.Round((clientPoint.X - imageRect.X) * (double)_originalImage.Width / imageRect.Width),
            (int)Math.Round((clientPoint.Y - imageRect.Y) * (double)_originalImage.Height / imageRect.Height));
        screenPoint = new Point(
            _info.CapturedRegion.X + Math.Clamp(imagePoint.X, 0, _originalImage.Width - 1),
            _info.CapturedRegion.Y + Math.Clamp(imagePoint.Y, 0, _originalImage.Height - 1));
        return true;
    }

    private CaptureSource? FindSourceAt(Point imagePoint, Point screenPoint)
    {
        foreach (var source in _info.Sources)
        {
            if (source.CaptureBounds.Width > 0 && source.CaptureBounds.Height > 0)
            {
                if (source.CaptureBounds.Contains(imagePoint))
                    return source;
            }
            else if (source.WindowBounds.Contains(screenPoint))
            {
                return source;
            }
        }

        return null;
    }

    private bool TryForwardClickToWindow(IntPtr hwnd, Point screenPoint)
    {
        if (hwnd == IntPtr.Zero || !IsWindow(hwnd))
            return false;

        _ = ForwardDoubleClickAsync(hwnd, screenPoint);
        return true;
    }

    private bool TryOpenSource(CaptureSource source)
    {
        if (!string.IsNullOrEmpty(source.Url))
        {
            return OpenUrlInCapturedBrowser(source.Url, source.ExePath, source.ProcessName);
        }

        if (!string.IsNullOrEmpty(source.FilePath) && File.Exists(source.FilePath))
        {
            Process.Start(new ProcessStartInfo(source.FilePath) { UseShellExecute = true });
            return true;
        }

        if (!string.IsNullOrEmpty(source.FolderPath) && Directory.Exists(source.FolderPath))
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{source.FolderPath}\"") { UseShellExecute = true });
            return true;
        }

        if (IsLaunchableExeFallback(source) &&
            !string.IsNullOrEmpty(source.ExePath) &&
            File.Exists(source.ExePath))
        {
            Process.Start(new ProcessStartInfo(source.ExePath) { UseShellExecute = true });
            return true;
        }

        return false;
    }

    private static bool OpenUrlInCapturedBrowser(string url, string? exePath, string processName)
    {
        try
        {
            var browserPath = ResolveBrowserPath(exePath, processName);
            if (!string.IsNullOrEmpty(browserPath))
            {
                var browserName = GetBrowserName(processName, browserPath);
                var psi = new ProcessStartInfo(browserPath)
                {
                    UseShellExecute = false
                };

                if (browserName == "firefox")
                    psi.ArgumentList.Add("-new-window");
                else
                    psi.ArgumentList.Add("--new-window");

                psi.ArgumentList.Add(url);
                Process.Start(psi);
                return true;
            }

            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string? ResolveBrowserPath(string? exePath, string processName)
    {
        if (!string.IsNullOrWhiteSpace(exePath) &&
            File.Exists(exePath) &&
            IsKnownBrowser(GetBrowserName(processName, exePath)))
        {
            return exePath;
        }

        var browserName = GetBrowserName(processName, exePath);
        foreach (var candidate in GetBrowserPathCandidates(browserName))
        {
            if (File.Exists(candidate))
                return candidate;
        }

        return null;
    }

    private static string GetBrowserName(string? processName, string? exePath)
    {
        var name = !string.IsNullOrWhiteSpace(processName)
            ? processName
            : Path.GetFileNameWithoutExtension(exePath ?? "");

        return name.ToLowerInvariant();
    }

    private static bool IsKnownBrowser(string browserName) =>
        browserName is "chrome" or "msedge" or "firefox" or "brave" or "opera" or "whale" or "vivaldi";

    private static IEnumerable<string> GetBrowserPathCandidates(string browserName)
    {
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        return browserName switch
        {
            "brave" => new[]
            {
                Path.Combine(programFiles, "BraveSoftware", "Brave-Browser", "Application", "brave.exe"),
                Path.Combine(programFilesX86, "BraveSoftware", "Brave-Browser", "Application", "brave.exe"),
                Path.Combine(localAppData, "BraveSoftware", "Brave-Browser", "Application", "brave.exe")
            },
            "chrome" => new[]
            {
                Path.Combine(programFiles, "Google", "Chrome", "Application", "chrome.exe"),
                Path.Combine(programFilesX86, "Google", "Chrome", "Application", "chrome.exe"),
                Path.Combine(localAppData, "Google", "Chrome", "Application", "chrome.exe")
            },
            "msedge" => new[]
            {
                Path.Combine(programFilesX86, "Microsoft", "Edge", "Application", "msedge.exe"),
                Path.Combine(programFiles, "Microsoft", "Edge", "Application", "msedge.exe")
            },
            "firefox" => new[]
            {
                Path.Combine(programFiles, "Mozilla Firefox", "firefox.exe"),
                Path.Combine(programFilesX86, "Mozilla Firefox", "firefox.exe")
            },
            "opera" => new[]
            {
                Path.Combine(localAppData, "Programs", "Opera", "opera.exe")
            },
            "whale" => new[]
            {
                Path.Combine(programFiles, "Naver", "Naver Whale", "Application", "whale.exe"),
                Path.Combine(programFilesX86, "Naver", "Naver Whale", "Application", "whale.exe"),
                Path.Combine(localAppData, "Naver", "Naver Whale", "Application", "whale.exe")
            },
            "vivaldi" => new[]
            {
                Path.Combine(programFiles, "Vivaldi", "Application", "vivaldi.exe"),
                Path.Combine(programFilesX86, "Vivaldi", "Application", "vivaldi.exe"),
                Path.Combine(localAppData, "Vivaldi", "Application", "vivaldi.exe")
            },
            _ => Array.Empty<string>()
        };
    }

    private static bool IsLaunchableExeFallback(CaptureSource source)
    {
        var process = source.ProcessName.ToLowerInvariant();
        return process is not ("applicationframehost" or "dllhost");
    }

    private async Task ForwardDoubleClickAsync(IntPtr hwnd, Point screenPoint)
    {
        _dragging = false;
        TopMost = false;
        if (_topMostItem != null)
            _topMostItem.Checked = false;

        ShowWindow(hwnd, SW_RESTORE);
        SetForegroundWindow(hwnd);
        await Task.Delay(120);

        SetCursorPos(screenPoint.X, screenPoint.Y);
        ClickMouse();
        await Task.Delay(60);
        ClickMouse();
    }

    private static void ClickMouse()
    {
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
    }

    private Rectangle GetImageClientRect() =>
        new(Border, Border, Width - Border * 2, Height - Border * 2);

    // --- Image transforms ---

    private void FlipHorizontal()
    {
        _originalImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
        Invalidate();
    }

    private void FlipVertical()
    {
        _originalImage.RotateFlip(RotateFlipType.RotateNoneFlipY);
        Invalidate();
    }

    private void RotateCW()
    {
        _originalImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
        Size = new Size(Height, Width);
        Invalidate();
    }

    private void RotateCCW()
    {
        _originalImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
        Size = new Size(Height, Width);
        Invalidate();
    }

    private async void RemoveBackground()
    {
        Cursor = Cursors.WaitCursor;
        Enabled = false;
        try
        {
            var tmpIn = Path.Combine(Path.GetTempPath(), $"sc_bg_in_{DateTime.Now:yyyyMMddHHmmss}.png");
            var tmpOut = Path.Combine(Path.GetTempPath(), $"sc_bg_out_{DateTime.Now:yyyyMMddHHmmss}.png");

            _originalImage.Save(tmpIn, ImageFormat.Png);

            var script = $"from rembg import remove;from PIL import Image;" +
                         $"remove(Image.open(r'{tmpIn}')).save(r'{tmpOut}')";

            var psi = new ProcessStartInfo
            {
                FileName = "python",
                Arguments = $"-c \"{script}\"",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardError = true
            };

            using var proc = Process.Start(psi);
            if (proc == null) throw new FileNotFoundException("python");

            await proc.WaitForExitAsync();

            if (proc.ExitCode == 0 && File.Exists(tmpOut))
            {
                using var fs = new FileStream(tmpOut, FileMode.Open, FileAccess.Read);
                var newImg = new Bitmap(fs);
                fs.Close();

                _originalImage.Dispose();
                _originalImage = new Bitmap(newImg);
                Invalidate();

                // Auto-save with _nobg suffix
                var ts = DateTime.Now.ToString("yyyyMMddHHmmss");
                var savePath = Path.Combine(_saveFolder, $"{ts}_nobg.png");
                Directory.CreateDirectory(_saveFolder);
                newImg.Save(savePath, ImageFormat.Png);

                // Spawn new sticky with the result
                var nobgInfo = new CaptureInfo
                {
                    ProcessName = _info.ProcessName,
                    WindowTitle = "배경 제거 결과",
                    Url = _info.Url,
                    FilePath = _info.FilePath,
                    FolderPath = _info.FolderPath,
                    ExePath = _info.ExePath,
                    ClipboardText = _info.ClipboardText,
                    OcrText = _info.OcrText,
                    CapturedImagePath = savePath,
                    CapturedRegion = _info.CapturedRegion,
                    SourceHwnd = _info.SourceHwnd,
                    SourceKind = _info.SourceKind,
                    SourceAnchor = _info.SourceAnchor,
                    Sources = CaptureInfo.CloneSources(_info.Sources)
                };
                NewStickyRequested?.Invoke(new Bitmap(newImg), nobgInfo);

                newImg.Dispose();
                try { File.Delete(tmpOut); } catch { }
            }
            else
            {
                var err = await proc.StandardError.ReadToEndAsync();
                MessageBox.Show(
                    $"배경 제거 실패:\n{err}\n\npip install rembg Pillow 로 설치해주세요.",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            try { File.Delete(tmpIn); } catch { }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Python을 찾을 수 없습니다.\n\n{ex.Message}",
                "AI 배경 제거", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        finally
        {
            Enabled = true;
            Cursor = Cursors.Default;
        }
    }

    private void SaveImageAs()
    {
        using var dlg = new SaveFileDialog
        {
            Filter = "PNG|*.png|JPG|*.jpg|BMP|*.bmp",
            FileName = $"sticky_{DateTime.Now:yyyyMMddHHmmss}"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        var ext = Path.GetExtension(dlg.FileName).ToLower();
        var fmt = ext switch
        {
            ".jpg" or ".jpeg" => ImageFormat.Jpeg,
            ".bmp" => ImageFormat.Bmp,
            _ => ImageFormat.Png
        };
        _originalImage.Save(dlg.FileName, fmt);
    }

    private void OpenAiTools()
    {
        var form = new AiToolsForm(_originalImage, _info, _saveFolder)
        {
            TopMost = TopMost
        };
        form.NewStickyRequested += (newImg, newInfo) => NewStickyRequested?.Invoke(newImg, newInfo);
        form.Show(this);
    }

    private async void OpenInChatGpt()
    {
        try
        {
            using var clipboardImage = new Bitmap(_originalImage);
            Clipboard.SetDataObject(clipboardImage, true);

            TopMost = false;
            if (_topMostItem != null)
                _topMostItem.Checked = false;

            Process.Start(new ProcessStartInfo("https://chatgpt.com/") { UseShellExecute = true });
            await Task.Delay(3500);
            SendKeys.SendWait("^v");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"ChatGPT로 이미지를 보내지 못했습니다.\n이미지는 클립보드에 남아 있을 수 있으니 ChatGPT에서 Ctrl+V를 눌러보세요.\n\n{ex.Message}",
                "ChatGPT 연결", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private async void TranslateImageText()
    {
        var text = _info.OcrText ?? _info.ClipboardText;

        if (string.IsNullOrWhiteSpace(text))
        {
            Cursor = Cursors.WaitCursor;
            Enabled = false;
            try
            {
                text = await OcrHelper.ExtractTextAsync(_originalImage);
                _info.OcrText = text;
            }
            finally
            {
                Enabled = true;
                Cursor = Cursors.Default;
            }
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            MessageBox.Show(
                "번역할 텍스트를 찾지 못했습니다.",
                "번역", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        text = text.Trim();
        Clipboard.SetText(text);

        var urlText = text.Length > 4500 ? text[..4500] : text;
        var url = "https://translate.google.com/?sl=auto&tl=ko&text=" +
                  Uri.EscapeDataString(urlText) +
                  "&op=translate";
        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
    }

    // --- Context menu ---

    private ContextMenuStrip BuildMenu()
    {
        var menu = new ContextMenuStrip { Font = new Font("Segoe UI", 9) };

        // Source actions
        if (!string.IsNullOrEmpty(_info.Url))
        {
            var item = menu.Items.Add("링크 열기", null, (_, _) =>
                OpenUrlInCapturedBrowser(_info.Url, _info.ExePath, _info.ProcessName));
            item.Font = new Font(menu.Font, FontStyle.Bold);
        }

        if (!string.IsNullOrEmpty(_info.FilePath) && File.Exists(_info.FilePath))
        {
            var fileItem = menu.Items.Add("파일 열기", null, (_, _) =>
                Process.Start(new ProcessStartInfo(_info.FilePath) { UseShellExecute = true }));
            if (string.IsNullOrEmpty(_info.Url))
                fileItem.Font = new Font(menu.Font, FontStyle.Bold);
            menu.Items.Add("위치 열기", null, (_, _) =>
                Process.Start("explorer.exe", $"/select,\"{_info.FilePath}\""));
        }

        if (!string.IsNullOrEmpty(_info.FolderPath) && Directory.Exists(_info.FolderPath))
        {
            var folderItem = menu.Items.Add("폴더 열기", null, (_, _) =>
                Process.Start(new ProcessStartInfo("explorer.exe", $"\"{_info.FolderPath}\"") { UseShellExecute = true }));
            if (string.IsNullOrEmpty(_info.Url) && string.IsNullOrEmpty(_info.FilePath))
                folderItem.Font = new Font(menu.Font, FontStyle.Bold);
        }

        // Captured image path
        if (!string.IsNullOrEmpty(_info.CapturedImagePath))
        {
            if (menu.Items.Count > 0) menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("캡쳐 경로 복사", null, (_, _) =>
                Clipboard.SetText(_info.CapturedImagePath));
            menu.Items.Add("캡쳐 위치 열기", null, (_, _) =>
                Process.Start("explorer.exe", $"/select,\"{_info.CapturedImagePath}\""));
        }

        // Text
        var text = _info.OcrText ?? _info.ClipboardText;
        if (!string.IsNullOrEmpty(text))
        {
            if (menu.Items.Count > 0) menu.Items.Add(new ToolStripSeparator());
            var display = text.Length > 30 ? text[..30] + "..." : text;
            menu.Items.Add($"텍스트 복사 ({display})", null, (_, _) =>
                Clipboard.SetText(text));
            menu.Items.Add("메모장에 열기", null, (_, _) => OpenInNotepad(text));
        }

        // Add-ons
        if (menu.Items.Count > 0) menu.Items.Add(new ToolStripSeparator());

        var addOnMenu = new ToolStripMenuItem("부가기능");
        addOnMenu.DropDownItems.Add("ChatGPT로 보내기", null, (_, _) => OpenInChatGpt());
        addOnMenu.DropDownItems.Add("배경 제거", null, (_, _) => RemoveBackground());
        addOnMenu.DropDownItems.Add(new ToolStripSeparator());
        addOnMenu.DropDownItems.Add(new ToolStripMenuItem("AI 이미지 편집 (미구현)")
        {
            Enabled = false
        });
        menu.Items.Add(addOnMenu);

        var editMenu = new ToolStripMenuItem("이미지 편집");
        editMenu.DropDownItems.Add("좌우 반전", null, (_, _) => FlipHorizontal());
        editMenu.DropDownItems.Add("상하 반전", null, (_, _) => FlipVertical());
        editMenu.DropDownItems.Add("시계방향 회전", null, (_, _) => RotateCW());
        editMenu.DropDownItems.Add("반시계방향 회전", null, (_, _) => RotateCCW());
        editMenu.DropDownItems.Add(new ToolStripSeparator());
        editMenu.DropDownItems.Add("다른 이름으로 저장...", null, (_, _) => SaveImageAs());
        menu.Items.Add(editMenu);

        // Window controls
        menu.Items.Add(new ToolStripSeparator());

        _topMostItem = new ToolStripMenuItem("항상 위(T)")
        {
            Checked = true,
            CheckOnClick = true
        };
        _topMostItem.CheckedChanged += (_, _) => TopMost = _topMostItem.Checked;
        menu.Items.Add(_topMostItem);

        menu.Items.Add("닫기", null, (_, _) => Close());

        return menu;
    }

    private static void OpenInNotepad(string text)
    {
        var path = Path.Combine(Path.GetTempPath(), $"capture_{DateTime.Now:yyyyMMddHHmmss}.txt");
        File.WriteAllText(path, text);
        Process.Start(new ProcessStartInfo("notepad.exe", $"\"{path}\"") { UseShellExecute = true });
    }

    private static void DrawCenterBadge(Graphics g, string text)
    {
        using var font = new Font("Segoe UI", 14, FontStyle.Bold);
        var size = g.MeasureString(text, font);
        var x = (g.VisibleClipBounds.Width - size.Width) / 2;
        var y = (g.VisibleClipBounds.Height - size.Height) / 2;
        using var bg = new SolidBrush(Color.FromArgb(160, 0, 0, 0));
        FillRoundedRect(g, bg, x - 8, y - 4, size.Width + 16, size.Height + 8, 6);
        g.DrawString(text, font, Brushes.White, x, y);
    }

    private static void FillRoundedRect(Graphics g, Brush brush, float x, float y, float w, float h, float r)
    {
        using var path = new GraphicsPath();
        path.AddArc(x, y, r * 2, r * 2, 180, 90);
        path.AddArc(x + w - r * 2, y, r * 2, r * 2, 270, 90);
        path.AddArc(x + w - r * 2, y + h - r * 2, r * 2, r * 2, 0, 90);
        path.AddArc(x, y + h - r * 2, r * 2, r * 2, 90, 90);
        path.CloseFigure();
        g.FillPath(brush, path);
    }
}
