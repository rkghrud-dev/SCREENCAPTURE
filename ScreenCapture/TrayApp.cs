using System.Drawing.Imaging;
using System.Media;
using System.Runtime.InteropServices;

namespace ScreenCapture;

public class TrayApp : ApplicationContext
{
    private const int HOTKEY_A = 1;
    private const int HOTKEY_B = 2;
    private const int HOTKEY_C = 3;

    private readonly NotifyIcon _tray;
    private readonly HotkeyWindow _hotkey;
    private AppSettings _settings;
    private readonly List<StickyWindow> _stickyWindows = new();

    [StructLayout(LayoutKind.Sequential)]
    private struct CURSORINFO { public int cbSize; public int flags; public IntPtr hCursor; public POINT ptScreenPos; }
    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int x, y; }
    [StructLayout(LayoutKind.Sequential)]
    private struct ICONINFO { public bool fIcon; public int xHotspot, yHotspot; public IntPtr hbmMask, hbmColor; }

    [DllImport("user32.dll")] private static extern bool GetCursorInfo(ref CURSORINFO pci);
    [DllImport("user32.dll")] private static extern bool GetIconInfo(IntPtr hIcon, out ICONINFO piconinfo);
    [DllImport("user32.dll")] private static extern bool DrawIcon(IntPtr hdc, int x, int y, IntPtr hIcon);
    [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr hObject);

    public TrayApp()
    {
        _settings = AppSettings.Load();

        _tray = new NotifyIcon
        {
            Icon = CreateIcon(),
            Text = "Screen Capture",
            Visible = true,
            ContextMenuStrip = BuildMenu()
        };
        _tray.DoubleClick += (_, _) => ShowOptions();

        _hotkey = new HotkeyWindow();
        _hotkey.HotkeyPressed += OnHotkey;
        RegisterHotkeys();

        _tray.ShowBalloonTip(2000, "Screen Capture",
            $"{_settings.HotkeyA.ToDisplayString()} : 캡쳐 → 클립보드\n" +
            $"{_settings.HotkeyB.ToDisplayString()} : 캡쳐 → 파일 + 경로\n" +
            $"{_settings.HotkeyC.ToDisplayString()} : 스티키 캡쳐\n" +
            $"저장: {_settings.SaveFolder}",
            ToolTipIcon.Info);
    }

    private void RegisterHotkeys()
    {
        _hotkey.UnregisterAll();

        var okA = _hotkey.Register(HOTKEY_A,
            _settings.HotkeyA.GetModifiers(), _settings.HotkeyA.GetKey());
        var okB = _hotkey.Register(HOTKEY_B,
            _settings.HotkeyB.GetModifiers(), _settings.HotkeyB.GetKey());
        var okC = _hotkey.Register(HOTKEY_C,
            _settings.HotkeyC.GetModifiers(), _settings.HotkeyC.GetKey());

        if (!okA || !okB || !okC)
        {
            _tray.ShowBalloonTip(3000, "경고",
                "일부 단축키를 등록할 수 없습니다.\n다른 프로그램이 사용 중일 수 있습니다.",
                ToolTipIcon.Warning);
        }
    }

    private ToolStripMenuItem? _pinToggleItem;

    private ContextMenuStrip BuildMenu()
    {
        var menu = new ContextMenuStrip();
        menu.Font = new Font("Segoe UI", 9);

        var captureMenu = new ToolStripMenuItem("화면 캡쳐 도구(C)");
        captureMenu.DropDownItems.Add(
            $"단축키 A  ({_settings.HotkeyA.ToDisplayString()})",
            null, (_, _) => DoCapture(_settings.ActionA));
        captureMenu.DropDownItems.Add(
            $"단축키 B  ({_settings.HotkeyB.ToDisplayString()})",
            null, (_, _) => DoCapture(_settings.ActionB));
        captureMenu.DropDownItems.Add(
            $"스티키 C  ({_settings.HotkeyC.ToDisplayString()})",
            null, (_, _) => DoCapture(_settings.ActionC, forcePin: true));
        captureMenu.DropDownItems.Add(new ToolStripSeparator());
        captureMenu.DropDownItems.Add("마지막 캡쳐 영역 반복(L)", null, (_, _) => RepeatLastCapture());
        menu.Items.Add(captureMenu);

        menu.Items.Add(new ToolStripSeparator());

        _pinToggleItem = new ToolStripMenuItem("스티키 모드(S)")
        {
            Checked = _settings.PinToDesktop,
            CheckOnClick = true
        };
        _pinToggleItem.CheckedChanged += (_, _) =>
        {
            _settings.PinToDesktop = _pinToggleItem.Checked;
            _settings.Save();
        };
        menu.Items.Add(_pinToggleItem);

        menu.Items.Add("모든 스티키 닫기(A)", null, (_, _) => CloseAllStickies());

        menu.Items.Add(new ToolStripSeparator());

        menu.Items.Add("옵션(O)...", null, (_, _) => ShowOptions());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("종료(X)", null, (_, _) => Exit());

        return menu;
    }

    private void OnHotkey(int id)
    {
        if (id == HOTKEY_A) DoCapture(_settings.ActionA);
        else if (id == HOTKEY_B) DoCapture(_settings.ActionB);
        else if (id == HOTKEY_C) DoCapture(_settings.ActionC, forcePin: true);
    }

    private Rectangle _lastRegion;
    private CaptureAction _lastAction = CaptureAction.ClipboardOnly;

    private async void DoCapture(CaptureAction action, bool forcePin = false)
    {
        if (_settings.CaptureDelay > 0)
            await Task.Delay(_settings.CaptureDelay);

        var snapshot = SourceDetector.SnapshotForeground();

        using var overlay = new SelectionOverlay();
        if (overlay.ShowDialog() != DialogResult.OK) return;

        using var img = overlay.GetSelectedImage();
        if (img == null) return;

        _lastRegion = overlay.SelectedRegion;
        _lastAction = action;

        if (_settings.IncludeCursor)
            DrawCursorOnImage(img, _lastRegion);

        if (_settings.PlaySound)
            SystemSounds.Exclamation.Play();

        var info = SourceDetector.Analyze(snapshot);
        info.CapturedRegion = _lastRegion;
        var savedPath = ProcessCapture(img, action);
        info.CapturedImagePath = savedPath;

        if (_settings.EnableOcr)
            info.OcrText = await OcrHelper.ExtractTextAsync(img);

        if (_settings.PinToDesktop || forcePin)
            ShowSticky(img, info);
    }

    private async void RepeatLastCapture()
    {
        if (_lastRegion.Width < 5 || _lastRegion.Height < 5)
        {
            _tray.ShowBalloonTip(1500, "알림", "이전 캡쳐 영역이 없습니다.", ToolTipIcon.Info);
            return;
        }

        var snapshot = SourceDetector.SnapshotForeground();

        using var bmp = new Bitmap(_lastRegion.Width, _lastRegion.Height);
        using (var g = Graphics.FromImage(bmp))
            g.CopyFromScreen(_lastRegion.Location, Point.Empty, _lastRegion.Size);

        if (_settings.IncludeCursor)
            DrawCursorOnImage(bmp, _lastRegion);

        if (_settings.PlaySound)
            SystemSounds.Exclamation.Play();

        var info = SourceDetector.Analyze(snapshot);
        info.CapturedRegion = _lastRegion;
        var savedPath = ProcessCapture(bmp, _lastAction);
        info.CapturedImagePath = savedPath;

        if (_settings.EnableOcr)
            info.OcrText = await OcrHelper.ExtractTextAsync(bmp);

        if (_settings.PinToDesktop)
            ShowSticky(bmp, info);
    }

    private string? ProcessCapture(Bitmap img, CaptureAction action)
    {
        string? savedPath = null;

        if (action is CaptureAction.FileAndClipboard or CaptureAction.FileOnly)
        {
            var ts = DateTime.Now.ToString(_settings.FileNamePattern);
            var ext = _settings.GetFileExtension();
            savedPath = Path.Combine(_settings.SaveFolder, $"{ts}{ext}");
            Directory.CreateDirectory(_settings.SaveFolder);
            SaveImage(img, savedPath);
        }

        switch (action)
        {
            case CaptureAction.ClipboardOnly:
                Clipboard.SetImage(img);
                _tray.ShowBalloonTip(1500, "캡쳐 완료",
                    $"{img.Width}x{img.Height} → 클립보드 (이미지)",
                    ToolTipIcon.Info);
                break;

            case CaptureAction.FileAndClipboard:
                Clipboard.SetText(savedPath!);
                _tray.ShowBalloonTip(1500, "저장 완료",
                    $"{savedPath}\n(경로가 클립보드에 복사됨)",
                    ToolTipIcon.Info);
                break;

            case CaptureAction.FileOnly:
                _tray.ShowBalloonTip(1500, "저장 완료",
                    savedPath!,
                    ToolTipIcon.Info);
                break;
        }

        return savedPath;
    }

    private void ShowSticky(Bitmap img, CaptureInfo info)
    {
        var sticky = new StickyWindow(img, info, _settings.SaveFolder);
        sticky.FormClosed += (_, _) => _stickyWindows.Remove(sticky);
        sticky.NewStickyRequested += (newImg, newInfo) => ShowSticky(newImg, newInfo);
        _stickyWindows.Add(sticky);
        sticky.Show();
    }

    private void CloseAllStickies()
    {
        foreach (var s in _stickyWindows.ToList())
            s.Close();
        _stickyWindows.Clear();
    }

    private void SaveImage(Bitmap img, string path)
    {
        var format = _settings.ImageFormat.ToUpper();

        if (format == "PNG")
        {
            img.Save(path, ImageFormat.Png);
        }
        else if (format == "BMP")
        {
            img.Save(path, ImageFormat.Bmp);
        }
        else
        {
            var encoder = ImageCodecInfo.GetImageEncoders()
                .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
            var encParams = new EncoderParameters(1);
            encParams.Param[0] = new EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality, (long)_settings.JpegQuality);
            img.Save(path, encoder, encParams);
        }
    }

    private static void DrawCursorOnImage(Bitmap bmp, Rectangle captureRegion)
    {
        var ci = new CURSORINFO { cbSize = Marshal.SizeOf<CURSORINFO>() };
        if (!GetCursorInfo(ref ci) || ci.flags == 0) return;

        var cursorX = ci.ptScreenPos.x - captureRegion.X;
        var cursorY = ci.ptScreenPos.y - captureRegion.Y;

        if (cursorX < 0 || cursorY < 0 || cursorX > bmp.Width || cursorY > bmp.Height)
            return;

        using var g = Graphics.FromImage(bmp);
        DrawIcon(g.GetHdc(), cursorX, cursorY, ci.hCursor);
        g.ReleaseHdc();
    }

    private void ShowOptions()
    {
        using var form = new SettingsForm(_settings);
        if (form.ShowDialog() == DialogResult.OK)
        {
            _settings = form.Result;
            _settings.Save();
            RegisterHotkeys();

            _tray.ContextMenuStrip?.Dispose();
            _tray.ContextMenuStrip = BuildMenu();

            _tray.ShowBalloonTip(1500, "설정 저장됨",
                $"{_settings.HotkeyA.ToDisplayString()} : 클립보드\n" +
                $"{_settings.HotkeyB.ToDisplayString()} : 파일 + 경로\n" +
                $"{_settings.HotkeyC.ToDisplayString()} : 스티키\n" +
                $"폴더: {_settings.SaveFolder}",
                ToolTipIcon.Info);
        }
    }

    private void Exit()
    {
        CloseAllStickies();
        _hotkey.Dispose();
        _tray.Visible = false;
        _tray.Dispose();
        Application.Exit();
    }

    private static Icon CreateIcon()
    {
        var bmp = new Bitmap(32, 32);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        using var bgBrush = new SolidBrush(Color.FromArgb(50, 120, 220));
        g.FillEllipse(bgBrush, 1, 1, 30, 30);

        using var pen = new Pen(Color.White, 2.5f);
        g.DrawRectangle(pen, 7, 11, 18, 13);
        g.DrawEllipse(pen, 12, 13, 8, 8);
        g.FillRectangle(Brushes.White, 10, 8, 6, 4);

        return Icon.FromHandle(bmp.GetHicon());
    }
}
