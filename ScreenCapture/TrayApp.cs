using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Media;
using System.Reflection;
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
    private CaptureLibraryForm? _libraryForm;

    private static readonly Color MenuBackground = Color.FromArgb(250, 251, 253);
    private static readonly Color MenuHoverBackground = Color.FromArgb(232, 240, 254);
    private static readonly Color MenuBorder = Color.FromArgb(218, 224, 235);
    private static readonly Color MenuText = Color.FromArgb(31, 41, 55);
    private static readonly Color MenuMutedText = Color.FromArgb(107, 114, 128);
    private static readonly Color MenuAccent = Color.FromArgb(37, 99, 235);
    private static readonly Color MenuSuccess = Color.FromArgb(22, 163, 74);
    private static readonly Color MenuDanger = Color.FromArgb(220, 38, 38);

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
        var menu = new ContextMenuStrip
        {
            Font = new Font("Segoe UI", 9.5f),
            BackColor = MenuBackground,
            ForeColor = MenuText,
            ImageScalingSize = new Size(18, 18),
            Padding = new Padding(6),
            Renderer = new ModernMenuRenderer(),
            ShowImageMargin = true,
            ShowItemToolTips = true
        };

        menu.Items.Add(MakeMenuHeader("Screen Capture"));
        menu.Items.Add(MakeMenuSubheader("빠른 캡쳐와 스티키 작업"));
        menu.Items.Add(new ToolStripSeparator());

        var captureMenu = MakeMenuItem(
            "캡쳐 실행",
            "설정된 단축키 작업을 바로 실행합니다.",
            CreateMenuIcon(MenuAccent, MenuIconKind.Capture));
        StyleDropDown(captureMenu.DropDown);
        captureMenu.DropDownItems.Add(MakeMenuItem(
            $"A · {DescribeAction(_settings.ActionA)}",
            "단축키 A 작업을 실행합니다.",
            CreateMenuIcon(MenuAccent, MenuIconKind.Capture),
            (_, _) => DoCapture(_settings.ActionA),
            _settings.HotkeyA.ToDisplayString()));
        captureMenu.DropDownItems.Add(MakeMenuItem(
            $"B · {DescribeAction(_settings.ActionB)}",
            "단축키 B 작업을 실행합니다.",
            CreateMenuIcon(MenuAccent, MenuIconKind.Capture),
            (_, _) => DoCapture(_settings.ActionB),
            _settings.HotkeyB.ToDisplayString()));
        captureMenu.DropDownItems.Add(MakeMenuItem(
            $"C · 스티키 캡쳐",
            "캡쳐 후 스티키 창으로 고정합니다.",
            CreateMenuIcon(MenuSuccess, MenuIconKind.Pin),
            (_, _) => DoCapture(_settings.ActionC, forcePin: true),
            _settings.HotkeyC.ToDisplayString()));
        captureMenu.DropDownItems.Add(new ToolStripSeparator());
        captureMenu.DropDownItems.Add(MakeMenuItem(
            "마지막 영역 다시 캡쳐",
            "이전에 선택한 영역을 같은 방식으로 다시 캡쳐합니다.",
            CreateMenuIcon(MenuAccent, MenuIconKind.Repeat),
            (_, _) => RepeatLastCapture(),
            "L"));
        menu.Items.Add(captureMenu);

        menu.Items.Add(MakeMenuItem(
            "캡쳐 보관함",
            "날짜별로 자동 저장된 캡쳐 이미지를 봅니다.",
            CreateMenuIcon(MenuAccent, MenuIconKind.Library),
            (_, _) => ShowCaptureLibrary()));

        menu.Items.Add(new ToolStripSeparator());

        _pinToggleItem = MakeMenuItem(
            "스티키 모드",
            "캡쳐 결과를 스티키 창으로 유지합니다.",
            CreateMenuIcon(MenuSuccess, MenuIconKind.Pin),
            shortcut: "S");
        _pinToggleItem.Checked = _settings.PinToDesktop;
        _pinToggleItem.CheckOnClick = true;
        _pinToggleItem.CheckedChanged += (_, _) =>
        {
            _settings.PinToDesktop = _pinToggleItem.Checked;
            _settings.Save();
        };
        menu.Items.Add(_pinToggleItem);

        menu.Items.Add(MakeMenuItem(
            "모든 스티키 닫기",
            "열려 있는 스티키 캡쳐 창을 모두 닫습니다.",
            CreateMenuIcon(MenuDanger, MenuIconKind.Close),
            (_, _) => CloseAllStickies(),
            "A"));

        menu.Items.Add(new ToolStripSeparator());

        menu.Items.Add(MakeMenuItem(
            "설정...",
            "저장, 캡쳐, 단축키 설정을 엽니다.",
            CreateMenuIcon(MenuAccent, MenuIconKind.Settings),
            (_, _) => ShowOptions(),
            "O"));

        var helpMenu = MakeMenuItem(
            "도움말 및 정보",
            "사용법과 버전 정보를 확인합니다.",
            CreateMenuIcon(MenuAccent, MenuIconKind.Help));
        StyleDropDown(helpMenu.DropDown);
        helpMenu.DropDownItems.Add(MakeMenuItem(
            "사용법 도움말",
            "주요 단축키와 스티키 창 조작법을 봅니다.",
            CreateMenuIcon(MenuAccent, MenuIconKind.Help),
            (_, _) => ShowHelp(),
            "H"));
        helpMenu.DropDownItems.Add(MakeMenuItem(
            "버전 정보",
            "앱 버전과 제품 정보를 확인합니다.",
            CreateMenuIcon(MenuAccent, MenuIconKind.Info),
            (_, _) => ShowAbout()));
        menu.Items.Add(helpMenu);

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(MakeMenuItem(
            "종료",
            "Screen Capture를 종료합니다.",
            CreateMenuIcon(MenuDanger, MenuIconKind.Exit),
            (_, _) => Exit(),
            "X"));

        return menu;
    }

    private static ToolStripLabel MakeMenuHeader(string text) =>
        new()
        {
            Text = text,
            AutoSize = false,
            Size = new Size(260, 24),
            Padding = new Padding(8, 2, 8, 0),
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
            ForeColor = MenuText
        };

    private static ToolStripLabel MakeMenuSubheader(string text) =>
        new()
        {
            Text = text,
            AutoSize = false,
            Size = new Size(260, 22),
            Padding = new Padding(8, 0, 8, 4),
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI", 8.5f),
            ForeColor = MenuMutedText
        };

    private static ToolStripMenuItem MakeMenuItem(
        string text,
        string? tooltip,
        Image? image,
        EventHandler? onClick = null,
        string? shortcut = null)
    {
        var item = new ToolStripMenuItem(text, image, onClick)
        {
            BackColor = MenuBackground,
            ForeColor = MenuText,
            Padding = new Padding(8, 5, 10, 5),
            Margin = new Padding(0, 1, 0, 1),
            ToolTipText = tooltip ?? string.Empty
        };

        if (!string.IsNullOrWhiteSpace(shortcut))
            item.ShortcutKeyDisplayString = shortcut;

        return item;
    }

    private static void StyleDropDown(ToolStripDropDown dropDown)
    {
        dropDown.BackColor = MenuBackground;
        dropDown.ForeColor = MenuText;
        dropDown.Padding = new Padding(6);
        dropDown.Renderer = new ModernMenuRenderer();

        if (dropDown is ToolStripDropDownMenu menu)
        {
            menu.ImageScalingSize = new Size(18, 18);
            menu.ShowImageMargin = true;
            menu.ShowItemToolTips = true;
        }
    }

    private enum MenuIconKind
    {
        Capture,
        Repeat,
        Pin,
        Close,
        Library,
        Help,
        Info,
        Settings,
        Exit
    }

    private static Image CreateMenuIcon(Color color, MenuIconKind kind)
    {
        var bmp = new Bitmap(18, 18);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;

        using var pen = new Pen(color, 1.7f)
        {
            StartCap = LineCap.Round,
            EndCap = LineCap.Round,
            LineJoin = LineJoin.Round
        };
        using var fill = new SolidBrush(color);

        switch (kind)
        {
            case MenuIconKind.Capture:
                g.DrawRectangle(pen, 3, 5, 12, 9);
                g.DrawEllipse(pen, 7, 7, 4, 4);
                g.FillRectangle(fill, 5, 3, 4, 2);
                break;

            case MenuIconKind.Repeat:
                g.DrawArc(pen, 4, 4, 10, 10, 40, 290);
                g.DrawLines(pen, new[] { new PointF(13, 3.5f), new PointF(15, 6.5f), new PointF(11.5f, 6.8f) });
                break;

            case MenuIconKind.Pin:
                g.DrawLine(pen, 6, 4, 12, 10);
                g.DrawLine(pen, 11, 3, 15, 7);
                g.DrawLine(pen, 4, 10, 8, 14);
                g.DrawLine(pen, 7, 9, 11, 13);
                break;

            case MenuIconKind.Close:
                g.DrawLine(pen, 5, 5, 13, 13);
                g.DrawLine(pen, 13, 5, 5, 13);
                break;

            case MenuIconKind.Library:
                g.DrawRectangle(pen, 3, 4, 12, 10);
                g.DrawLine(pen, 6, 7, 12, 7);
                g.DrawLine(pen, 6, 10, 11, 10);
                g.FillRectangle(fill, 5, 3, 5, 2);
                break;

            case MenuIconKind.Help:
                DrawCenteredGlyph(g, "?", color);
                break;

            case MenuIconKind.Info:
                DrawCenteredGlyph(g, "i", color);
                break;

            case MenuIconKind.Settings:
                g.DrawLine(pen, 4, 5, 14, 5);
                g.FillEllipse(fill, 7, 3, 4, 4);
                g.DrawLine(pen, 4, 9, 14, 9);
                g.FillEllipse(fill, 11, 7, 4, 4);
                g.DrawLine(pen, 4, 13, 14, 13);
                g.FillEllipse(fill, 5, 11, 4, 4);
                break;

            case MenuIconKind.Exit:
                g.DrawRectangle(pen, 4, 4, 7, 10);
                g.DrawLine(pen, 9, 9, 15, 9);
                g.DrawLines(pen, new[] { new PointF(12, 6), new PointF(15, 9), new PointF(12, 12) });
                break;
        }

        return bmp;
    }

    private static void DrawCenteredGlyph(Graphics g, string glyph, Color color)
    {
        using var font = new Font("Segoe UI", 11, FontStyle.Bold);
        using var brush = new SolidBrush(color);
        using var sf = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center
        };
        g.DrawString(glyph, font, brush, new RectangleF(0, 0, 18, 18), sf);
    }

    private sealed class ModernMenuRenderer : ToolStripProfessionalRenderer
    {
        public ModernMenuRenderer() : base(new ModernMenuColorTable())
        {
        }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            using var pen = new Pen(MenuBorder);
            var bounds = new Rectangle(Point.Empty, e.ToolStrip.Size);
            bounds.Width -= 1;
            bounds.Height -= 1;
            e.Graphics.DrawRectangle(pen, bounds);
        }

        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            var bounds = new Rectangle(Point.Empty, e.Item.Size);
            var selected = e.Item.Selected && e.Item.Enabled;

            using var bgBrush = new SolidBrush(selected ? MenuHoverBackground : MenuBackground);
            e.Graphics.FillRectangle(bgBrush, bounds);

            if (selected)
            {
                using var accentBrush = new SolidBrush(MenuAccent);
                e.Graphics.FillRectangle(accentBrush, new Rectangle(0, 5, 3, Math.Max(0, bounds.Height - 10)));
            }
        }

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            if (e.ToolStrip == null) return;

            var y = e.Item.Height / 2;
            using var pen = new Pen(MenuBorder);
            e.Graphics.DrawLine(pen, e.ToolStrip.DisplayRectangle.Left + 30, y, e.ToolStrip.DisplayRectangle.Right - 8, y);
        }
    }

    private sealed class ModernMenuColorTable : ProfessionalColorTable
    {
        public override Color ToolStripDropDownBackground => MenuBackground;
        public override Color ImageMarginGradientBegin => MenuBackground;
        public override Color ImageMarginGradientMiddle => MenuBackground;
        public override Color ImageMarginGradientEnd => MenuBackground;
        public override Color MenuItemSelected => MenuHoverBackground;
        public override Color MenuItemBorder => MenuHoverBackground;
        public override Color SeparatorDark => MenuBorder;
        public override Color SeparatorLight => MenuBorder;
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
        var visibleWindows = SourceDetector.SnapshotVisibleWindows();

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

        var info = SourceDetector.Analyze(snapshot, visibleWindows, _lastRegion);
        info.CapturedRegion = _lastRegion;
        var savedPath = ProcessCapture(img, action);
        var historyPath = AutoSaveHistoryCapture(img);
        info.CapturedImagePath = savedPath ?? historyPath;
        RefreshLibraryIfOpen();

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
        var visibleWindows = SourceDetector.SnapshotVisibleWindows();

        using var bmp = new Bitmap(_lastRegion.Width, _lastRegion.Height);
        using (var g = Graphics.FromImage(bmp))
            g.CopyFromScreen(_lastRegion.Location, Point.Empty, _lastRegion.Size);

        if (_settings.IncludeCursor)
            DrawCursorOnImage(bmp, _lastRegion);

        if (_settings.PlaySound)
            SystemSounds.Exclamation.Play();

        var info = SourceDetector.Analyze(snapshot, visibleWindows, _lastRegion);
        info.CapturedRegion = _lastRegion;
        var savedPath = ProcessCapture(bmp, _lastAction);
        var historyPath = AutoSaveHistoryCapture(bmp);
        info.CapturedImagePath = savedPath ?? historyPath;
        RefreshLibraryIfOpen();

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

    private string? AutoSaveHistoryCapture(Bitmap img)
    {
        if (!_settings.AutoSaveCaptures)
            return null;

        try
        {
            var now = DateTime.Now;
            var folder = Path.Combine(_settings.GetHistoryRoot(), now.ToString("yyyy-MM-dd"));
            Directory.CreateDirectory(folder);

            var path = Path.Combine(folder, $"capture_{now:HHmmss_fff}{_settings.GetFileExtension()}");
            var index = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(folder, $"capture_{now:HHmmss_fff}_{index}{_settings.GetFileExtension()}");
                index++;
            }

            SaveImage(img, path);
            return path;
        }
        catch
        {
            return null;
        }
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

            if (_libraryForm is { IsDisposed: false })
            {
                _libraryForm.Close();
                _libraryForm = null;
            }

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

    private void ShowCaptureLibrary()
    {
        if (_libraryForm is { IsDisposed: false })
        {
            _libraryForm.Reload();
            _libraryForm.Activate();
            return;
        }

        _libraryForm = new CaptureLibraryForm(_settings);
        _libraryForm.FormClosed += (_, _) => _libraryForm = null;
        _libraryForm.Show();
    }

    private void RefreshLibraryIfOpen()
    {
        if (_libraryForm is { IsDisposed: false })
            _libraryForm.Reload();
    }

    private void ShowHelp()
    {
        var help =
            "Screen Capture 사용법\n\n" +
            $"단축키 A: {_settings.HotkeyA.ToDisplayString()} - {DescribeAction(_settings.ActionA)}\n" +
            $"단축키 B: {_settings.HotkeyB.ToDisplayString()} - {DescribeAction(_settings.ActionB)}\n" +
            $"스티키 C: {_settings.HotkeyC.ToDisplayString()} - {DescribeAction(_settings.ActionC)} + 스티키 표시\n\n" +
            "기본 조작\n" +
            "- 캡쳐: 단축키를 누른 뒤 원하는 영역 드래그\n" +
            "- 정밀 선택: Shift를 누른 상태로 확대경 사용\n" +
            "- 마지막 영역 반복: 트레이 메뉴 > 마지막 캡쳐 영역 반복\n\n" +
            "스티키 창\n" +
            "- 드래그: 창 이동\n" +
            "- Shift + 드래그: 크기 조절\n" +
            "- 마우스 휠: 투명도 조절\n" +
            "- Shift + 휠: 확대/축소\n" +
            "- 더블클릭: 캡쳐한 위치의 원본 열기\n" +
            "  브라우저는 저장된 주소를 열고, 탐색기는 폴더를 엽니다.\n" +
            "- 우클릭: 저장, OCR, ChatGPT, 배경 제거 등 부가기능\n" +
            "- Esc/Delete: 스티키 닫기";

        MessageBox.Show(help, "사용법 도움말", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void ShowAbout()
    {
        var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0";

        using var form = new Form
        {
            Text = "버전 정보",
            Size = new Size(430, 270),
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            StartPosition = FormStartPosition.CenterScreen,
            TopMost = true,
            BackColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };

        var header = new Panel
        {
            Dock = DockStyle.Top,
            Height = 96,
            BackColor = MenuBackground
        };
        header.Paint += (_, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            using var iconBrush = new SolidBrush(MenuAccent);
            e.Graphics.FillEllipse(iconBrush, 24, 24, 48, 48);

            using var iconPen = new Pen(Color.White, 3)
            {
                LineJoin = LineJoin.Round
            };
            e.Graphics.DrawRectangle(iconPen, 38, 43, 20, 14);
            e.Graphics.DrawEllipse(iconPen, 44, 46, 8, 8);
            e.Graphics.FillRectangle(Brushes.White, 41, 39, 7, 4);
        };

        var title = new Label
        {
            Text = "Screen Capture",
            Location = new Point(88, 24),
            AutoSize = true,
            Font = new Font("Segoe UI", 15, FontStyle.Bold),
            ForeColor = MenuText,
            BackColor = Color.Transparent
        };
        header.Controls.Add(title);

        var versionLabel = new Label
        {
            Text = $"버전 {version}",
            Location = new Point(90, 55),
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5f),
            ForeColor = MenuMutedText,
            BackColor = Color.Transparent
        };
        header.Controls.Add(versionLabel);
        form.Controls.Add(header);

        var description = new Label
        {
            Text = "빠른 영역 캡쳐, 파일 저장, 스티키 캡쳐를 위한 데스크톱 유틸리티입니다.",
            Location = new Point(24, 116),
            Size = new Size(374, 42),
            ForeColor = MenuText
        };
        form.Controls.Add(description);

        var copyright = new Label
        {
            Text = $"Copyright © {DateTime.Now.Year} Screen Capture. All rights reserved.",
            Location = new Point(24, 162),
            Size = new Size(374, 22),
            ForeColor = MenuMutedText
        };
        form.Controls.Add(copyright);

        var btnOk = new Button
        {
            Text = "확인",
            Size = new Size(88, 30),
            Location = new Point(318, 196),
            DialogResult = DialogResult.OK
        };
        form.Controls.Add(btnOk);
        form.AcceptButton = btnOk;
        form.CancelButton = btnOk;

        form.ShowDialog();
    }

    private static string DescribeAction(CaptureAction action) =>
        action switch
        {
            CaptureAction.ClipboardOnly => "클립보드에 이미지 복사",
            CaptureAction.FileAndClipboard => "파일 저장 + 경로 복사",
            CaptureAction.FileOnly => "파일 저장",
            _ => action.ToString()
        };

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
