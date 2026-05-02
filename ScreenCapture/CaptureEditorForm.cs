using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace ScreenCapture;

public class CaptureEditorForm : Form
{
    private readonly Bitmap _originalImage;
    private readonly CaptureInfo _info;
    private readonly string _saveFolder;
    private readonly EditorCanvas _canvas = new();
    private readonly List<Annotation> _annotations = new();
    private readonly Dictionary<EditorTool, ToolStripButton> _toolButtons = new();

    private ToolStripButton _undoButton = null!;
    private ToolStripButton _clearButton = null!;
    private ToolStripDropDownButton _colorButton = null!;
    private ToolStripComboBox _widthCombo = null!;
    private ToolStripComboBox _fontSizeCombo = null!;
    private Label _statusLabel = null!;
    private RadioButton _saveEditedOnly = null!;
    private RadioButton _saveOriginalAndEdited = null!;
    private RadioButton _saveOriginalOnly = null!;

    private RectangleF _imageRect;
    private EditorTool _currentTool = EditorTool.Highlighter;
    private Color _currentColor = Color.FromArgb(245, 197, 24);
    private int _strokeWidth = 6;
    private float _fontSize = 24f;
    private bool _dragging;
    private bool _resizingText;
    private PointF _startImagePoint;
    private RectangleF _resizeStartBounds;
    private Annotation? _draft;
    private Annotation? _selectedText;
    private TextBox? _activeTextBox;
    private RectangleF _activeTextImageBounds;

    public event Action<Bitmap, CaptureInfo>? NewStickyRequested;
    public event Action? CapturesSaved;

    public CaptureEditorForm(Bitmap image, CaptureInfo info, string? saveFolder = null)
    {
        _originalImage = new Bitmap(image);
        _info = CloneInfo(info, info.CapturedImagePath, info.WindowTitle);
        _saveFolder = string.IsNullOrWhiteSpace(saveFolder)
            ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            : saveFolder;

        Text = "상세 캡쳐 편집";
        StartPosition = FormStartPosition.CenterScreen;
        WindowState = FormWindowState.Maximized;
        MinimumSize = new Size(980, 680);
        KeyPreview = true;
        Font = new Font("Segoe UI", 9.5f);
        BackColor = Color.FromArgb(242, 244, 248);

        BuildUi();
        SelectTool(EditorTool.Highlighter);
        UpdateCommandState();
    }

    private void BuildUi()
    {
        var toolbar = new ToolStrip
        {
            Dock = DockStyle.Top,
            GripStyle = ToolStripGripStyle.Hidden,
            ImageScalingSize = new Size(18, 18),
            Padding = new Padding(6, 5, 6, 5),
            BackColor = Color.White,
            RenderMode = ToolStripRenderMode.System
        };

        AddToolButton(toolbar, "형광펜", EditorTool.Highlighter);
        AddToolButton(toolbar, "펜", EditorTool.Pen);
        AddToolButton(toolbar, "밑줄", EditorTool.Underline);
        AddToolButton(toolbar, "사각형", EditorTool.Rectangle);
        AddToolButton(toolbar, "원", EditorTool.Ellipse);
        AddToolButton(toolbar, "선", EditorTool.Line);
        AddToolButton(toolbar, "화살표", EditorTool.Arrow);
        AddToolButton(toolbar, "텍스트", EditorTool.Text);
        toolbar.Items.Add(new ToolStripSeparator());

        _colorButton = new ToolStripDropDownButton("색상");
        AddColorItem("노랑", Color.FromArgb(245, 197, 24));
        AddColorItem("빨강", Color.FromArgb(239, 68, 68));
        AddColorItem("파랑", Color.FromArgb(37, 99, 235));
        AddColorItem("초록", Color.FromArgb(22, 163, 74));
        AddColorItem("검정", Color.FromArgb(17, 24, 39));
        AddColorItem("흰색", Color.White);
        toolbar.Items.Add(_colorButton);

        toolbar.Items.Add(new ToolStripLabel("두께"));
        _widthCombo = new ToolStripComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 54
        };
        _widthCombo.Items.AddRange(new object[] { "2", "4", "6", "10", "16", "24" });
        _widthCombo.SelectedItem = "6";
        _widthCombo.SelectedIndexChanged += (_, _) =>
        {
            if (int.TryParse(_widthCombo.SelectedItem?.ToString(), out var width))
                _strokeWidth = width;
        };
        toolbar.Items.Add(_widthCombo);

        toolbar.Items.Add(new ToolStripLabel("글자"));
        _fontSizeCombo = new ToolStripComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 58
        };
        _fontSizeCombo.Items.AddRange(new object[] { "14", "18", "24", "32", "44", "60" });
        _fontSizeCombo.SelectedItem = "24";
        _fontSizeCombo.SelectedIndexChanged += (_, _) =>
        {
            if (float.TryParse(_fontSizeCombo.SelectedItem?.ToString(), out var size))
                _fontSize = size;
        };
        toolbar.Items.Add(_fontSizeCombo);

        toolbar.Items.Add(new ToolStripSeparator());
        _undoButton = new ToolStripButton("실행취소") { DisplayStyle = ToolStripItemDisplayStyle.Text };
        _undoButton.Click += (_, _) => UndoLast();
        toolbar.Items.Add(_undoButton);

        _clearButton = new ToolStripButton("전체삭제") { DisplayStyle = ToolStripItemDisplayStyle.Text };
        _clearButton.Click += (_, _) => ClearAnnotations();
        toolbar.Items.Add(_clearButton);

        Controls.Add(toolbar);

        _canvas.Dock = DockStyle.Fill;
        _canvas.BackColor = Color.FromArgb(226, 232, 240);
        _canvas.Cursor = Cursors.Cross;
        _canvas.Paint += Canvas_Paint;
        _canvas.MouseDown += Canvas_MouseDown;
        _canvas.MouseMove += Canvas_MouseMove;
        _canvas.MouseUp += Canvas_MouseUp;
        _canvas.Resize += (_, _) => _canvas.Invalidate();
        Controls.Add(_canvas);

        var bottom = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 58,
            Padding = new Padding(12, 10, 12, 8),
            BackColor = Color.White
        };

        var flow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            AutoScroll = true
        };
        bottom.Controls.Add(flow);

        _saveEditedOnly = MakeRadio("수정본만 저장", true);
        _saveOriginalAndEdited = MakeRadio("원본+수정본 저장", false);
        _saveOriginalOnly = MakeRadio("원본만 저장", false);

        flow.Controls.Add(new Label
        {
            Text = "저장:",
            AutoSize = true,
            Margin = new Padding(0, 8, 8, 0)
        });
        flow.Controls.Add(_saveEditedOnly);
        flow.Controls.Add(_saveOriginalAndEdited);
        flow.Controls.Add(_saveOriginalOnly);
        flow.Controls.Add(MakeButton("수정본 복사", CopyEditedToClipboard));
        flow.Controls.Add(MakeButton("원본 복사", CopyOriginalToClipboard));
        flow.Controls.Add(MakeButton("저장", SaveSelectedImages));
        flow.Controls.Add(MakeButton("스티키로 열기", OpenEditedAsSticky));
        flow.Controls.Add(MakeButton("닫기", Close));

        _statusLabel = new Label
        {
            Dock = DockStyle.Right,
            Width = 360,
            TextAlign = ContentAlignment.MiddleRight,
            ForeColor = Color.FromArgb(75, 85, 99)
        };
        bottom.Controls.Add(_statusLabel);

        Controls.Add(bottom);
    }

    private void AddToolButton(ToolStrip toolbar, string text, EditorTool tool)
    {
        var button = new ToolStripButton(text)
        {
            CheckOnClick = false,
            DisplayStyle = ToolStripItemDisplayStyle.Text
        };
        button.Click += (_, _) => SelectTool(tool);
        _toolButtons[tool] = button;
        toolbar.Items.Add(button);
    }

    private void AddColorItem(string name, Color color)
    {
        var item = new ToolStripMenuItem(name)
        {
            Image = CreateColorSwatch(color)
        };
        item.Click += (_, _) =>
        {
            _currentColor = color;
            _colorButton.Text = $"색상: {name}";
            _colorButton.Image = CreateColorSwatch(color);
        };
        _colorButton.DropDownItems.Add(item);

        if (_colorButton.DropDownItems.Count == 1)
        {
            _colorButton.Text = $"색상: {name}";
            _colorButton.Image = CreateColorSwatch(color);
        }
    }

    private static Image CreateColorSwatch(Color color)
    {
        var bmp = new Bitmap(18, 18);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var brush = new SolidBrush(color);
        using var border = new Pen(Color.FromArgb(120, 31, 41, 55));
        g.FillEllipse(brush, 3, 3, 12, 12);
        g.DrawEllipse(border, 3, 3, 12, 12);
        return bmp;
    }

    private static RadioButton MakeRadio(string text, bool isChecked) =>
        new()
        {
            Text = text,
            Checked = isChecked,
            AutoSize = true,
            Margin = new Padding(0, 8, 12, 0)
        };

    private static Button MakeButton(string text, Action action)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            Height = 31,
            Margin = new Padding(0, 4, 8, 0),
            Padding = new Padding(10, 0, 10, 0)
        };
        button.Click += (_, _) => action();
        return button;
    }

    private void SelectTool(EditorTool tool)
    {
        CommitTextBox(save: true);
        _currentTool = tool;

        if (tool == EditorTool.Highlighter && _currentColor == Color.White)
            _currentColor = Color.FromArgb(245, 197, 24);

        foreach (var pair in _toolButtons)
            pair.Value.Checked = pair.Key == tool;

        _canvas.Cursor = tool == EditorTool.Text ? Cursors.IBeam : Cursors.Cross;
        SetStatus(tool == EditorTool.Text
            ? "텍스트 선택: 드래그로 박스를 만들고, 기존 텍스트는 우하단 핸들로 크기 조절"
            : $"{ToolName(tool)} 선택");
        _canvas.Invalidate();
    }

    private static string ToolName(EditorTool tool) => tool switch
    {
        EditorTool.Highlighter => "형광펜",
        EditorTool.Pen => "펜",
        EditorTool.Underline => "밑줄",
        EditorTool.Rectangle => "사각형",
        EditorTool.Ellipse => "원",
        EditorTool.Line => "선",
        EditorTool.Arrow => "화살표",
        EditorTool.Text => "텍스트",
        _ => "도구"
    };

    private void Canvas_Paint(object? sender, PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(_canvas.BackColor);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;

        _imageRect = CalculateImageRect(_canvas.ClientRectangle);
        if (_imageRect.Width <= 1 || _imageRect.Height <= 1)
            return;

        using (var shadow = new SolidBrush(Color.FromArgb(30, 15, 23, 42)))
            g.FillRectangle(shadow, _imageRect.X + 4, _imageRect.Y + 5, _imageRect.Width, _imageRect.Height);

        g.DrawImage(_originalImage, _imageRect);
        DrawAnnotations(g, _imageRect, _annotations, showSelection: true);

        if (_draft != null)
            DrawAnnotation(g, _imageRect, _draft, showSelection: false);

        using var border = new Pen(Color.FromArgb(148, 163, 184), 1f);
        g.DrawRectangle(border, Rectangle.Round(_imageRect));
    }

    private RectangleF CalculateImageRect(Rectangle client)
    {
        const int margin = 24;
        var availableWidth = Math.Max(1, client.Width - margin * 2);
        var availableHeight = Math.Max(1, client.Height - margin * 2);
        var scale = Math.Min(availableWidth / (float)_originalImage.Width, availableHeight / (float)_originalImage.Height);
        scale = Math.Min(scale, 1.0f);
        var width = _originalImage.Width * scale;
        var height = _originalImage.Height * scale;
        var x = client.X + (client.Width - width) / 2f;
        var y = client.Y + (client.Height - height) / 2f;
        return new RectangleF(x, y, width, height);
    }

    private void Canvas_MouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left)
            return;

        CommitTextBox(save: true);

        if (!TryClientToImage(e.Location, out var imagePoint))
            return;

        if (_currentTool == EditorTool.Text)
        {
            if (TryHitTextResizeHandle(imagePoint, out var resizeTarget))
            {
                _selectedText = resizeTarget;
                _resizingText = true;
                _startImagePoint = imagePoint;
                _resizeStartBounds = resizeTarget.Bounds;
                _canvas.Capture = true;
                _canvas.Invalidate();
                return;
            }

            if (TryHitTextAnnotation(imagePoint, out var selected))
            {
                _selectedText = selected;
                _canvas.Invalidate();
                return;
            }
        }
        else
        {
            _selectedText = null;
        }

        _dragging = true;
        _startImagePoint = imagePoint;
        _draft = CreateDraftAnnotation(imagePoint);
        _canvas.Capture = true;
    }

    private void Canvas_MouseMove(object? sender, MouseEventArgs e)
    {
        if (_resizingText && _selectedText != null)
        {
            if (!TryClientToImage(e.Location, out var imagePoint))
                imagePoint = ClampToImage(e.Location);

            var width = Math.Max(20, imagePoint.X - _resizeStartBounds.X);
            var height = Math.Max(18, imagePoint.Y - _resizeStartBounds.Y);
            _selectedText.Bounds = ClampImageBounds(new RectangleF(
                _resizeStartBounds.X, _resizeStartBounds.Y, width, height));
            _selectedText.End = new PointF(_selectedText.Bounds.Right, _selectedText.Bounds.Bottom);
            _canvas.Invalidate();
            return;
        }

        if (!_dragging || _draft == null)
            return;

        if (!TryClientToImage(e.Location, out var currentPoint))
            currentPoint = ClampToImage(e.Location);

        if (_draft.Tool is EditorTool.Highlighter or EditorTool.Pen)
        {
            var last = _draft.Points.Count > 0 ? _draft.Points[^1] : _startImagePoint;
            if (Distance(last, currentPoint) >= 1.5f)
                _draft.Points.Add(currentPoint);
        }
        else
        {
            if (_draft.Tool == EditorTool.Underline)
                currentPoint = new PointF(currentPoint.X, _startImagePoint.Y);

            _draft.End = currentPoint;
            _draft.Bounds = MakeBounds(_startImagePoint, currentPoint);
        }

        _canvas.Invalidate();
    }

    private void Canvas_MouseUp(object? sender, MouseEventArgs e)
    {
        if (_resizingText)
        {
            _resizingText = false;
            _canvas.Capture = false;
            SetStatus("텍스트 박스 크기 변경");
            _canvas.Invalidate();
            return;
        }

        if (!_dragging || _draft == null)
            return;

        _dragging = false;
        _canvas.Capture = false;

        if (_draft.Tool == EditorTool.Text)
        {
            var textBounds = _draft.Bounds.Width < 8 || _draft.Bounds.Height < 8
                ? new RectangleF(_draft.Start.X, _draft.Start.Y, 280 / GetCanvasScale(), 70 / GetCanvasScale())
                : _draft.Bounds;
            _draft = null;
            BeginTextEdit(textBounds);
            _canvas.Invalidate();
            return;
        }

        if (IsAnnotationUsable(_draft))
        {
            _annotations.Add(_draft);
            SetStatus($"{ToolName(_draft.Tool)} 추가");
        }

        _draft = null;
        UpdateCommandState();
        _canvas.Invalidate();
    }

    private Annotation CreateDraftAnnotation(PointF imagePoint)
    {
        var width = _currentTool == EditorTool.Highlighter
            ? Math.Max(12, _strokeWidth * 3)
            : _strokeWidth;

        var annotation = new Annotation
        {
            Tool = _currentTool,
            Color = _currentColor,
            Width = width,
            FontSize = _fontSize,
            Start = imagePoint,
            End = imagePoint,
            Bounds = new RectangleF(imagePoint, SizeF.Empty)
        };

        if (_currentTool is EditorTool.Highlighter or EditorTool.Pen)
            annotation.Points.Add(imagePoint);

        return annotation;
    }

    private bool IsAnnotationUsable(Annotation annotation)
    {
        if (annotation.Tool is EditorTool.Highlighter or EditorTool.Pen)
            return annotation.Points.Count > 1;

        if (annotation.Tool == EditorTool.Text)
            return !string.IsNullOrWhiteSpace(annotation.Text);

        return Math.Abs(annotation.End.X - annotation.Start.X) >= 3 ||
               Math.Abs(annotation.End.Y - annotation.Start.Y) >= 3;
    }

    private void BeginTextEdit(RectangleF imageBounds)
    {
        _activeTextImageBounds = ClampImageBounds(imageBounds);
        var canvasBounds = Rectangle.Round(ToCanvasRect(_imageRect, _activeTextImageBounds));
        canvasBounds.Width = Math.Max(120, canvasBounds.Width);
        canvasBounds.Height = Math.Max(40, canvasBounds.Height);

        var textBox = new TextBox
        {
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Malgun Gothic", Math.Max(10f, _fontSize * GetCanvasScale())),
            Multiline = true,
            AcceptsReturn = true,
            Bounds = canvasBounds,
            BackColor = Color.White,
            ForeColor = _currentColor == Color.White ? Color.Black : _currentColor
        };

        textBox.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Enter && e.Control)
            {
                e.SuppressKeyPress = true;
                CommitTextBox(save: true);
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                CommitTextBox(save: false);
            }
        };
        textBox.Leave += (_, _) => CommitTextBox(save: true);

        _activeTextBox = textBox;
        _canvas.Controls.Add(textBox);
        textBox.Focus();
    }

    private void CommitTextBox(bool save)
    {
        if (_activeTextBox == null)
            return;

        var textBox = _activeTextBox;
        _activeTextBox = null;

        var text = textBox.Text.Trim();
        _canvas.Controls.Remove(textBox);
        textBox.Dispose();

        if (save && !string.IsNullOrWhiteSpace(text))
        {
            var annotation = new Annotation
            {
                Tool = EditorTool.Text,
                Color = _currentColor,
                Width = _strokeWidth,
                FontSize = _fontSize,
                Start = _activeTextImageBounds.Location,
                End = new PointF(_activeTextImageBounds.Right, _activeTextImageBounds.Bottom),
                Bounds = _activeTextImageBounds,
                Text = text
            };
            _annotations.Add(annotation);
            _selectedText = annotation;
            SetStatus("텍스트 추가");
            UpdateCommandState();
            _canvas.Invalidate();
        }
    }

    private bool TryClientToImage(Point clientPoint, out PointF imagePoint)
    {
        if (!_imageRect.Contains(clientPoint))
        {
            imagePoint = PointF.Empty;
            return false;
        }

        imagePoint = new PointF(
            (clientPoint.X - _imageRect.X) * _originalImage.Width / _imageRect.Width,
            (clientPoint.Y - _imageRect.Y) * _originalImage.Height / _imageRect.Height);
        return true;
    }

    private PointF ClampToImage(Point clientPoint)
    {
        var x = Math.Clamp(clientPoint.X, _imageRect.Left, _imageRect.Right);
        var y = Math.Clamp(clientPoint.Y, _imageRect.Top, _imageRect.Bottom);
        return new PointF(
            (x - _imageRect.X) * _originalImage.Width / _imageRect.Width,
            (y - _imageRect.Y) * _originalImage.Height / _imageRect.Height);
    }

    private float GetCanvasScale()
    {
        if (_imageRect.Width <= 0)
            return 1f;

        return _imageRect.Width / _originalImage.Width;
    }

    private static float Distance(PointF a, PointF b)
    {
        var dx = a.X - b.X;
        var dy = a.Y - b.Y;
        return MathF.Sqrt(dx * dx + dy * dy);
    }

    private static RectangleF MakeBounds(PointF a, PointF b)
    {
        var x = Math.Min(a.X, b.X);
        var y = Math.Min(a.Y, b.Y);
        return new RectangleF(x, y, Math.Abs(a.X - b.X), Math.Abs(a.Y - b.Y));
    }

    private RectangleF ClampImageBounds(RectangleF bounds)
    {
        var x = Math.Clamp(bounds.X, 0, Math.Max(0, _originalImage.Width - 1));
        var y = Math.Clamp(bounds.Y, 0, Math.Max(0, _originalImage.Height - 1));
        var width = Math.Clamp(bounds.Width, 1, Math.Max(1, _originalImage.Width - x));
        var height = Math.Clamp(bounds.Height, 1, Math.Max(1, _originalImage.Height - y));
        return new RectangleF(x, y, width, height);
    }

    private bool TryHitTextAnnotation(PointF imagePoint, out Annotation annotation)
    {
        for (var i = _annotations.Count - 1; i >= 0; i--)
        {
            var candidate = _annotations[i];
            if (candidate.Tool == EditorTool.Text && candidate.Bounds.Contains(imagePoint))
            {
                annotation = candidate;
                return true;
            }
        }

        annotation = null!;
        return false;
    }

    private bool TryHitTextResizeHandle(PointF imagePoint, out Annotation annotation)
    {
        for (var i = _annotations.Count - 1; i >= 0; i--)
        {
            var candidate = _annotations[i];
            if (candidate.Tool == EditorTool.Text && GetTextResizeHandle(candidate.Bounds).Contains(imagePoint))
            {
                annotation = candidate;
                return true;
            }
        }

        annotation = null!;
        return false;
    }

    private RectangleF GetTextResizeHandle(RectangleF bounds)
    {
        var size = Math.Max(8f, 12f / GetCanvasScale());
        return new RectangleF(bounds.Right - size, bounds.Bottom - size, size, size);
    }

    private void DrawAnnotations(Graphics g, RectangleF imageRect, IEnumerable<Annotation> annotations, bool showSelection)
    {
        foreach (var annotation in annotations)
            DrawAnnotation(g, imageRect, annotation, showSelection);
    }

    private void DrawAnnotation(Graphics g, RectangleF imageRect, Annotation annotation, bool showSelection)
    {
        var scale = imageRect.Width / _originalImage.Width;
        var width = Math.Max(1f, annotation.Width * scale);

        switch (annotation.Tool)
        {
            case EditorTool.Highlighter:
                DrawFreehand(g, imageRect, annotation, width, highlighter: true);
                break;

            case EditorTool.Pen:
                DrawFreehand(g, imageRect, annotation, width, highlighter: false);
                break;

            case EditorTool.Rectangle:
                using (var pen = MakePen(annotation.Color, width))
                    g.DrawRectangle(pen, Rectangle.Round(ToCanvasRect(imageRect, annotation.Bounds)));
                break;

            case EditorTool.Ellipse:
                using (var pen = MakePen(annotation.Color, width))
                    g.DrawEllipse(pen, ToCanvasRect(imageRect, annotation.Bounds));
                break;

            case EditorTool.Line:
            case EditorTool.Underline:
                using (var pen = MakePen(annotation.Color, width))
                    g.DrawLine(pen, ToCanvasPoint(imageRect, annotation.Start), ToCanvasPoint(imageRect, annotation.End));
                break;

            case EditorTool.Arrow:
                using (var pen = MakePen(annotation.Color, width))
                {
                    pen.CustomEndCap = new AdjustableArrowCap(width * 2.2f, width * 2.8f, true);
                    g.DrawLine(pen, ToCanvasPoint(imageRect, annotation.Start), ToCanvasPoint(imageRect, annotation.End));
                }
                break;

            case EditorTool.Text:
                if (string.IsNullOrWhiteSpace(annotation.Text))
                    DrawTextDraft(g, imageRect, annotation, scale);
                else
                    DrawTextAnnotation(g, imageRect, annotation, scale, showSelection);
                break;
        }
    }

    private static Pen MakePen(Color color, float width) =>
        new(color, width)
        {
            StartCap = LineCap.Round,
            EndCap = LineCap.Round,
            LineJoin = LineJoin.Round
        };

    private void DrawFreehand(Graphics g, RectangleF imageRect, Annotation annotation, float width, bool highlighter)
    {
        if (annotation.Points.Count < 2)
            return;

        using var pen = new Pen(highlighter ? Color.FromArgb(105, annotation.Color) : annotation.Color, width)
        {
            StartCap = LineCap.Round,
            EndCap = LineCap.Round,
            LineJoin = LineJoin.Round
        };

        var points = annotation.Points.Select(p => ToCanvasPoint(imageRect, p)).ToArray();
        g.DrawLines(pen, points);
    }

    private void DrawTextAnnotation(Graphics g, RectangleF imageRect, Annotation annotation, float scale, bool showSelection)
    {
        if (string.IsNullOrWhiteSpace(annotation.Text))
            return;

        var box = ToCanvasRect(imageRect, annotation.Bounds);
        box.Width = Math.Max(24, box.Width);
        box.Height = Math.Max(24, box.Height);
        var content = RectangleF.Inflate(box, -6, -4);

        using var font = new Font("Malgun Gothic", Math.Max(8f, annotation.FontSize * scale), FontStyle.Bold, GraphicsUnit.Pixel);
        using var format = new StringFormat
        {
            Alignment = StringAlignment.Near,
            LineAlignment = StringAlignment.Near,
            Trimming = StringTrimming.EllipsisWord
        };

        using var background = new SolidBrush(Color.FromArgb(232, Color.White));
        using var border = new Pen(Color.FromArgb(170, annotation.Color), Math.Max(1f, annotation.Width * scale / 2f));
        using var textBrush = new SolidBrush(annotation.Color == Color.White ? Color.Black : annotation.Color);

        g.FillRectangle(background, box);
        g.DrawRectangle(border, Rectangle.Round(box));
        g.DrawString(annotation.Text, font, textBrush, content, format);

        if (!showSelection || !ReferenceEquals(annotation, _selectedText))
            return;

        using var selectPen = new Pen(Color.FromArgb(37, 99, 235), 1.5f)
        {
            DashStyle = DashStyle.Dash
        };
        g.DrawRectangle(selectPen, Rectangle.Round(box));

        var handle = ToCanvasRect(imageRect, GetTextResizeHandle(annotation.Bounds));
        using var handleBrush = new SolidBrush(Color.FromArgb(37, 99, 235));
        using var handleBorder = new Pen(Color.White, 1f);
        g.FillRectangle(handleBrush, handle);
        g.DrawRectangle(handleBorder, Rectangle.Round(handle));
    }

    private void DrawTextDraft(Graphics g, RectangleF imageRect, Annotation annotation, float scale)
    {
        var box = ToCanvasRect(imageRect, annotation.Bounds);
        if (box.Width < 2 || box.Height < 2)
            return;

        using var pen = new Pen(Color.FromArgb(180, annotation.Color), Math.Max(1f, annotation.Width * scale / 2f))
        {
            DashStyle = DashStyle.Dash
        };
        g.DrawRectangle(pen, Rectangle.Round(box));
    }

    private PointF ToCanvasPoint(RectangleF imageRect, PointF imagePoint) =>
        new(
            imageRect.X + imagePoint.X * imageRect.Width / _originalImage.Width,
            imageRect.Y + imagePoint.Y * imageRect.Height / _originalImage.Height);

    private RectangleF ToCanvasRect(RectangleF imageRect, RectangleF imageBounds)
    {
        var start = ToCanvasPoint(imageRect, imageBounds.Location);
        return new RectangleF(
            start.X,
            start.Y,
            imageBounds.Width * imageRect.Width / _originalImage.Width,
            imageBounds.Height * imageRect.Height / _originalImage.Height);
    }

    private Bitmap RenderEditedImage()
    {
        var result = new Bitmap(_originalImage.Width, _originalImage.Height, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(result);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.DrawImageUnscaled(_originalImage, Point.Empty);
        DrawAnnotations(g, new RectangleF(0, 0, _originalImage.Width, _originalImage.Height), _annotations, showSelection: false);
        return result;
    }

    private void CopyEditedToClipboard()
    {
        CommitTextBox(save: true);
        using var result = RenderEditedImage();
        Clipboard.SetDataObject(result, true);
        SetStatus("수정본을 클립보드에 복사했습니다.");
    }

    private void CopyOriginalToClipboard()
    {
        using var original = new Bitmap(_originalImage);
        Clipboard.SetDataObject(original, true);
        SetStatus("원본을 클립보드에 복사했습니다.");
    }

    private void SaveSelectedImages()
    {
        CommitTextBox(save: true);

        try
        {
            Directory.CreateDirectory(_saveFolder);
            var now = DateTime.Now;
            var baseName = now.ToString("yyyyMMddHHmmss");
            var saved = new List<string>();

            if (_saveOriginalOnly.Checked || _saveOriginalAndEdited.Checked)
            {
                var originalPath = MakeUniquePath(_saveFolder, baseName, "original");
                _originalImage.Save(originalPath, ImageFormat.Png);
                CaptureHistoryMetadata.Save(originalPath, CloneInfo(_info, originalPath, "상세 캡쳐 원본"));
                saved.Add(Path.GetFileName(originalPath));
            }

            if (_saveEditedOnly.Checked || _saveOriginalAndEdited.Checked)
            {
                using var result = RenderEditedImage();
                var editedPath = MakeUniquePath(_saveFolder, baseName, "edited");
                result.Save(editedPath, ImageFormat.Png);
                CaptureHistoryMetadata.Save(editedPath, CloneInfo(_info, editedPath, "상세 캡쳐 수정본"));
                saved.Add(Path.GetFileName(editedPath));
            }

            CapturesSaved?.Invoke();
            SetStatus($"저장 완료: {string.Join(", ", saved)}");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"캡쳐 이미지를 저장하지 못했습니다.\n\n{ex.Message}",
                "상세 캡쳐 저장", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void OpenEditedAsSticky()
    {
        CommitTextBox(save: true);
        using var result = RenderEditedImage();
        NewStickyRequested?.Invoke(new Bitmap(result), CloneInfo(_info, null, "상세 캡쳐 수정본"));
        SetStatus("수정본을 스티키로 열었습니다.");
    }

    private void UndoLast()
    {
        CommitTextBox(save: false);
        if (_annotations.Count == 0)
            return;

        var removed = _annotations[^1];
        _annotations.RemoveAt(_annotations.Count - 1);
        if (ReferenceEquals(_selectedText, removed))
            _selectedText = null;

        SetStatus("마지막 수정내용을 취소했습니다.");
        UpdateCommandState();
        _canvas.Invalidate();
    }

    private void ClearAnnotations()
    {
        CommitTextBox(save: false);
        if (_annotations.Count == 0)
            return;

        _annotations.Clear();
        _selectedText = null;
        SetStatus("수정내용을 모두 삭제했습니다.");
        UpdateCommandState();
        _canvas.Invalidate();
    }

    private void UpdateCommandState()
    {
        var hasAnnotations = _annotations.Count > 0;
        _undoButton.Enabled = hasAnnotations;
        _clearButton.Enabled = hasAnnotations;
    }

    private void SetStatus(string status)
    {
        if (_statusLabel != null)
            _statusLabel.Text = status;
    }

    private static string MakeUniquePath(string folder, string baseName, string suffix)
    {
        var path = Path.Combine(folder, $"{baseName}_{suffix}.png");
        var index = 1;
        while (File.Exists(path))
        {
            path = Path.Combine(folder, $"{baseName}_{suffix}_{index}.png");
            index++;
        }

        return path;
    }

    private static CaptureInfo CloneInfo(CaptureInfo source, string? capturedImagePath, string? title)
    {
        return new CaptureInfo
        {
            ProcessName = source.ProcessName,
            WindowTitle = string.IsNullOrWhiteSpace(title) ? source.WindowTitle : title,
            Url = source.Url,
            FilePath = source.FilePath,
            FolderPath = source.FolderPath,
            ExePath = source.ExePath,
            ClipboardText = source.ClipboardText,
            OcrText = source.OcrText,
            CapturedImagePath = capturedImagePath,
            CapturedRegion = source.CapturedRegion,
            SourceHwnd = source.SourceHwnd,
            SourceKind = source.SourceKind,
            SourceAnchor = source.SourceAnchor,
            Sources = CaptureInfo.CloneSources(source.Sources)
        };
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == (Keys.Control | Keys.Z))
        {
            UndoLast();
            return true;
        }

        if (keyData == (Keys.Control | Keys.C))
        {
            CopyEditedToClipboard();
            return true;
        }

        if (keyData == (Keys.Control | Keys.S))
        {
            SaveSelectedImages();
            return true;
        }

        if (keyData == Keys.Escape)
        {
            if (_activeTextBox != null)
                CommitTextBox(save: false);
            else
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
        }

        base.Dispose(disposing);
    }

    private enum EditorTool
    {
        Highlighter,
        Pen,
        Underline,
        Rectangle,
        Ellipse,
        Line,
        Arrow,
        Text
    }

    private sealed class Annotation
    {
        public EditorTool Tool { get; init; }
        public Color Color { get; init; }
        public int Width { get; init; }
        public float FontSize { get; init; }
        public PointF Start { get; set; }
        public PointF End { get; set; }
        public RectangleF Bounds { get; set; }
        public string Text { get; init; } = "";
        public List<PointF> Points { get; } = new();
    }

    private sealed class EditorCanvas : Panel
    {
        public EditorCanvas()
        {
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw |
                     ControlStyles.UserPaint, true);
        }
    }
}
