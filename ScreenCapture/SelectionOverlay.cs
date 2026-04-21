using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace ScreenCapture;

public class SelectionOverlay : Form
{
    private readonly Bitmap _screenshot;
    private Point _start;
    private Point _current;
    private bool _selecting;
    private bool _precisionSelecting;
    private Rectangle _selected;
    private bool _magnifierOn;
    private Point _mouseScreen;
    private readonly System.Windows.Forms.Timer _shiftPollTimer;

    private const int MagZoom = 8;
    private const int MagRadius = 60;
    private const int MagSize = MagRadius * 2 + 1;

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    public Rectangle SelectedRegion => _selected;

    public Bitmap? GetSelectedImage()
    {
        if (_selected.Width < 5 || _selected.Height < 5) return null;

        var vScreen = SystemInformation.VirtualScreen;
        var srcRect = new Rectangle(
            _selected.X - vScreen.X,
            _selected.Y - vScreen.Y,
            _selected.Width,
            _selected.Height);

        var result = new Bitmap(srcRect.Width, srcRect.Height);
        using var g = Graphics.FromImage(result);
        g.DrawImage(_screenshot,
            new Rectangle(0, 0, srcRect.Width, srcRect.Height),
            srcRect, GraphicsUnit.Pixel);
        return result;
    }

    public SelectionOverlay()
    {
        var vScreen = SystemInformation.VirtualScreen;

        _screenshot = new Bitmap(vScreen.Width, vScreen.Height);
        using (var g = Graphics.FromImage(_screenshot))
            g.CopyFromScreen(vScreen.Location, Point.Empty, vScreen.Size);

        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        Location = vScreen.Location;
        Size = vScreen.Size;
        TopMost = true;
        ShowInTaskbar = false;
        Cursor = Cursors.Cross;
        DoubleBuffered = true;
        KeyPreview = true;

        _mouseScreen = ClampToVirtualScreen(Cursor.Position);
        _shiftPollTimer = new System.Windows.Forms.Timer { Interval = 20 };
        _shiftPollTimer.Tick += (_, _) => SyncMagnifierState();
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        Activate();
        Focus();

        _mouseScreen = ClampToVirtualScreen(Cursor.Position);
        SyncMagnifierState();
        _shiftPollTimer.Start();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        var vScreen = SystemInformation.VirtualScreen;

        g.DrawImage(_screenshot, 0, 0);
        using (var dimBrush = new SolidBrush(Color.FromArgb(120, 0, 0, 0)))
            g.FillRectangle(dimBrush, 0, 0, vScreen.Width, vScreen.Height);

        if (_selecting && _selected.Width > 0 && _selected.Height > 0)
        {
            var srcRect = new Rectangle(
                _selected.X - vScreen.X,
                _selected.Y - vScreen.Y,
                _selected.Width, _selected.Height);
            var destRect = new Rectangle(
                _selected.X - Left,
                _selected.Y - Top,
                _selected.Width, _selected.Height);

            g.DrawImage(_screenshot, destRect, srcRect, GraphicsUnit.Pixel);

            using var pen = new Pen(Color.Red, 2);
            g.DrawRectangle(pen, destRect);

            var sizeText = $"{_selected.Width} x {_selected.Height}";
            using var font = new Font("Segoe UI", 10, FontStyle.Bold);
            var textSize = g.MeasureString(sizeText, font);
            var textX = destRect.X;
            var textY = destRect.Y - textSize.Height - 4;
            if (textY < 0) textY = destRect.Bottom + 4;

            using var bgBrush = new SolidBrush(Color.FromArgb(200, 0, 0, 0));
            g.FillRectangle(bgBrush, textX, textY, textSize.Width + 4, textSize.Height);
            using var textBrush = new SolidBrush(Color.White);
            g.DrawString(sizeText, font, textBrush, textX + 2, textY);
        }

        if (_magnifierOn)
            DrawMagnifier(g);
    }

    private void DrawMagnifier(Graphics g)
    {
        var vScreen = SystemInformation.VirtualScreen;
        var localPt = new Point(_mouseScreen.X - Left, _mouseScreen.Y - Top);
        var srcX = _mouseScreen.X - vScreen.X;
        var srcY = _mouseScreen.Y - vScreen.Y;

        var pixelsToShow = MagSize / MagZoom;
        var halfPx = pixelsToShow / 2;

        // Magnifier position: offset from cursor to avoid overlap
        var magX = localPt.X + 20;
        var magY = localPt.Y + 20;
        if (magX + MagSize + 4 > Width) magX = localPt.X - MagSize - 24;
        if (magY + MagSize + 30 > Height) magY = localPt.Y - MagSize - 50;

        // Background
        using var bgBrush = new SolidBrush(Color.FromArgb(220, 30, 30, 30));
        g.FillRectangle(bgBrush, magX - 2, magY - 2, MagSize + 4, MagSize + 28);

        // Draw zoomed pixels
        g.InterpolationMode = InterpolationMode.NearestNeighbor;
        g.PixelOffsetMode = PixelOffsetMode.Half;

        var srcRect = new Rectangle(
            Math.Clamp(srcX - halfPx, 0, _screenshot.Width - pixelsToShow),
            Math.Clamp(srcY - halfPx, 0, _screenshot.Height - pixelsToShow),
            pixelsToShow, pixelsToShow);
        var destRect = new Rectangle(magX, magY, MagSize, MagSize);

        g.DrawImage(_screenshot, destRect, srcRect, GraphicsUnit.Pixel);

        g.InterpolationMode = InterpolationMode.Default;
        g.PixelOffsetMode = PixelOffsetMode.Default;

        // Grid lines
        using var gridPen = new Pen(Color.FromArgb(60, 255, 255, 255), 1);
        for (int i = 0; i <= pixelsToShow; i++)
        {
            var offset = i * MagZoom;
            g.DrawLine(gridPen, magX + offset, magY, magX + offset, magY + MagSize);
            g.DrawLine(gridPen, magX, magY + offset, magX + MagSize, magY + offset);
        }

        // Crosshair at center
        var cx = magX + (MagSize / 2);
        var cy = magY + (MagSize / 2);
        using var crossPen = new Pen(Color.Red, 1);
        g.DrawRectangle(crossPen, cx - MagZoom / 2, cy - MagZoom / 2, MagZoom, MagZoom);

        // Border
        using var borderPen = new Pen(Color.FromArgb(180, 255, 255, 255), 2);
        g.DrawRectangle(borderPen, magX - 1, magY - 1, MagSize + 2, MagSize + 2);

        // Pixel color & coordinates
        var px = Math.Clamp(srcX, 0, _screenshot.Width - 1);
        var py = Math.Clamp(srcY, 0, _screenshot.Height - 1);
        var pixelColor = _screenshot.GetPixel(px, py);
        var infoText = $"({_mouseScreen.X}, {_mouseScreen.Y})  #{pixelColor.R:X2}{pixelColor.G:X2}{pixelColor.B:X2}";

        using var infoFont = new Font("Consolas", 9, FontStyle.Bold);
        using var infoBg = new SolidBrush(Color.FromArgb(200, 0, 0, 0));
        using var infoFg = new SolidBrush(Color.White);
        var infoSize = g.MeasureString(infoText, infoFont);
        g.FillRectangle(infoBg, magX, magY + MagSize + 2, infoSize.Width + 6, infoSize.Height + 2);
        g.DrawString(infoText, infoFont, infoFg, magX + 3, magY + MagSize + 3);

        // Color preview swatch
        using var colorBrush = new SolidBrush(pixelColor);
        g.FillRectangle(colorBrush, magX + MagSize - 18, magY + MagSize + 4, 16, 12);
        g.DrawRectangle(Pens.White, magX + MagSize - 18, magY + MagSize + 4, 16, 12);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left)
            return;

        var screenPoint = PointToScreen(e.Location);
        _mouseScreen = screenPoint;

        if (IsPrecisionModeActive())
        {
            if (!_selecting)
                BeginSelection(screenPoint, precision: true);
            else
                CompleteSelection(screenPoint);

            Invalidate();
            return;
        }

        BeginSelection(screenPoint, precision: false);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        _mouseScreen = PointToScreen(e.Location);

        if (_selecting)
        {
            _current = _mouseScreen;
            _selected = MakeRect(_start, _current);
        }

        if (_magnifierOn || _selecting)
            Invalidate();
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left && _selecting)
        {
            _current = PointToScreen(e.Location);
            _mouseScreen = _current;
            _selected = MakeRect(_start, _current);

            if (_precisionSelecting)
            {
                Invalidate();
                return;
            }

            _selecting = false;
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape)
        {
            DialogResult = DialogResult.Cancel;
            Close();
            return;
        }

        // Shift toggles magnifier
        if (e.KeyCode == Keys.ShiftKey)
        {
            _mouseScreen = ClampToVirtualScreen(Cursor.Position);
            SetMagnifier(true);
            Invalidate();
            return;
        }

        // Arrow keys: 1px fine adjustment
        if (HandleArrowKey(e.KeyCode))
            return;
    }

    protected override void OnKeyUp(KeyEventArgs e)
    {
        if (e.KeyCode == Keys.ShiftKey)
        {
            SetMagnifier(IsShiftDown() || _precisionSelecting);
            Invalidate();
        }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        var keyCode = keyData & Keys.KeyCode;
        if (HandleArrowKey(keyCode))
            return true;

        return base.ProcessCmdKey(ref msg, keyData);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _shiftPollTimer.Stop();
            _shiftPollTimer.Dispose();
            _screenshot.Dispose();
        }
        base.Dispose(disposing);
    }

    private bool HandleArrowKey(Keys keyCode)
    {
        if (keyCode is not (Keys.Left or Keys.Right or Keys.Up or Keys.Down))
            return false;

        int dx = keyCode == Keys.Left ? -1 : keyCode == Keys.Right ? 1 : 0;
        int dy = keyCode == Keys.Up ? -1 : keyCode == Keys.Down ? 1 : 0;

        var target = _selecting ? _current : _mouseScreen;
        if (target == Point.Empty)
            target = Cursor.Position;

        target = ClampToVirtualScreen(new Point(target.X + dx, target.Y + dy));
        _mouseScreen = target;

        if (_selecting)
        {
            _current = target;
            _selected = MakeRect(_start, _current);
        }

        Cursor.Position = target;
        Invalidate();
        return true;
    }

    private void SyncMagnifierState()
    {
        if (SetMagnifier(IsShiftDown() || _precisionSelecting))
            Invalidate();
    }

    private bool SetMagnifier(bool enabled)
    {
        if (_magnifierOn == enabled)
            return false;

        _magnifierOn = enabled;
        return true;
    }

    private static bool IsShiftDown() =>
        (GetAsyncKeyState((int)Keys.ShiftKey) & 0x8000) != 0;

    private bool IsPrecisionModeActive() =>
        _magnifierOn || _precisionSelecting || IsShiftDown();

    private void BeginSelection(Point screenPoint, bool precision)
    {
        _selecting = true;
        _precisionSelecting = precision;
        _start = screenPoint;
        _current = screenPoint;
        _mouseScreen = screenPoint;
        _selected = MakeRect(_start, _current);

        if (precision)
            SetMagnifier(true);
    }

    private void CompleteSelection(Point screenPoint)
    {
        _current = screenPoint;
        _mouseScreen = screenPoint;
        _selected = MakeRect(_start, _current);

        if (_selected.Width < 5 || _selected.Height < 5)
            return;

        _selecting = false;
        _precisionSelecting = false;
        DialogResult = DialogResult.OK;
        Close();
    }

    private static Point ClampToVirtualScreen(Point point)
    {
        var bounds = SystemInformation.VirtualScreen;
        return new Point(
            Math.Clamp(point.X, bounds.Left, bounds.Right - 1),
            Math.Clamp(point.Y, bounds.Top, bounds.Bottom - 1));
    }

    private static Rectangle MakeRect(Point a, Point b) =>
        new(Math.Min(a.X, b.X), Math.Min(a.Y, b.Y),
            Math.Abs(a.X - b.X), Math.Abs(a.Y - b.Y));
}
