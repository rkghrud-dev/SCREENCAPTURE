using System.Diagnostics;
using System.Drawing.Drawing2D;

namespace ScreenCapture;

public class CaptureLibraryForm : Form
{
    private static readonly string[] ImageExtensions = { ".jpg", ".jpeg", ".png", ".bmp" };

    private readonly AppSettings _settings;
    private readonly ListBox _dateList;
    private readonly ListView _captureList;
    private readonly ImageList _thumbs;
    private readonly PictureBox _preview;
    private readonly Label _detailTitle;
    private readonly Label _detailMeta;
    private readonly Label _emptyLabel;
    private readonly Button _openButton;
    private readonly Button _folderButton;
    private readonly Button _copyButton;
    private readonly Button _deleteButton;
    private readonly List<DateBucket> _buckets = new();

    private Image? _previewImage;
    private string? _selectedPath;

    public CaptureLibraryForm(AppSettings settings)
    {
        _settings = settings.Clone();

        Text = "캡쳐 보관함";
        Size = new Size(980, 640);
        MinimumSize = new Size(820, 520);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9);
        BackColor = Color.FromArgb(246, 248, 251);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Padding = new Padding(14)
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 210));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        Controls.Add(root);

        var leftPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(12)
        };
        leftPanel.Paint += DrawPanelBorder;
        root.Controls.Add(leftPanel, 0, 0);

        var leftLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            RowCount = 3,
            ColumnCount = 1
        };
        leftLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        leftLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        leftLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        leftPanel.Controls.Add(leftLayout);

        var title = new Label
        {
            Text = "날짜별 캡쳐",
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            ForeColor = Color.FromArgb(17, 24, 39),
            TextAlign = ContentAlignment.MiddleLeft
        };
        leftLayout.Controls.Add(title, 0, 0);

        _dateList = new ListBox
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            DrawMode = DrawMode.OwnerDrawFixed,
            ItemHeight = 34,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(31, 41, 55)
        };
        _dateList.DrawItem += DateList_DrawItem;
        _dateList.SelectedIndexChanged += (_, _) => LoadSelectedDate();
        leftLayout.Controls.Add(_dateList, 0, 1);

        var refreshButton = MakeButton("새로고침");
        refreshButton.Dock = DockStyle.Bottom;
        refreshButton.Click += (_, _) => Reload();
        leftLayout.Controls.Add(refreshButton, 0, 2);

        var mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(14)
        };
        mainPanel.Paint += DrawPanelBorder;
        root.Controls.Add(mainPanel, 1, 0);

        var mainLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3
        };
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
        mainPanel.Controls.Add(mainLayout);

        var toolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false
        };
        mainLayout.Controls.Add(toolbar, 0, 0);

        _openButton = MakeButton("열기");
        _openButton.Click += (_, _) => OpenSelected();
        toolbar.Controls.Add(_openButton);

        _folderButton = MakeButton("위치 열기");
        _folderButton.Click += (_, _) => OpenSelectedFolder();
        toolbar.Controls.Add(_folderButton);

        _copyButton = MakeButton("경로 복사");
        _copyButton.Click += (_, _) => CopySelectedPath();
        toolbar.Controls.Add(_copyButton);

        _deleteButton = MakeButton("삭제");
        _deleteButton.Click += (_, _) => DeleteSelected();
        toolbar.Controls.Add(_deleteButton);

        var openRootButton = MakeButton("보관함 폴더");
        openRootButton.Click += (_, _) => OpenHistoryRoot();
        toolbar.Controls.Add(openRootButton);

        var bodySplit = new SplitContainer
        {
            Dock = DockStyle.Fill,
            SplitterDistance = 480,
            BackColor = Color.White
        };
        mainLayout.Controls.Add(bodySplit, 0, 1);

        _thumbs = new ImageList
        {
            ColorDepth = ColorDepth.Depth32Bit,
            ImageSize = new Size(168, 104)
        };

        _captureList = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.LargeIcon,
            LargeImageList = _thumbs,
            BorderStyle = BorderStyle.None,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(31, 41, 55),
            MultiSelect = false,
            HideSelection = false
        };
        _captureList.SelectedIndexChanged += (_, _) => ShowSelectedPreview();
        _captureList.DoubleClick += (_, _) => OpenSelected();
        bodySplit.Panel1.Controls.Add(_captureList);

        _emptyLabel = new Label
        {
            Text = "저장된 캡쳐가 없습니다.",
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = Color.FromArgb(107, 114, 128),
            Visible = false
        };
        bodySplit.Panel1.Controls.Add(_emptyLabel);
        _emptyLabel.BringToFront();

        _preview = new PictureBox
        {
            Dock = DockStyle.Fill,
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.FromArgb(243, 244, 246)
        };
        bodySplit.Panel2.Controls.Add(_preview);

        var detailPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 10, 0, 0)
        };
        mainLayout.Controls.Add(detailPanel, 0, 2);

        _detailTitle = new Label
        {
            Dock = DockStyle.Top,
            Height = 28,
            Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(17, 24, 39)
        };
        detailPanel.Controls.Add(_detailTitle);

        _detailMeta = new Label
        {
            Dock = DockStyle.Fill,
            ForeColor = Color.FromArgb(107, 114, 128)
        };
        detailPanel.Controls.Add(_detailMeta);

        Reload();
        UpdateActionButtons();
    }

    public void Reload()
    {
        _buckets.Clear();
        _dateList.Items.Clear();

        var root = _settings.GetHistoryRoot();
        Directory.CreateDirectory(root);

        var directories = Directory.GetDirectories(root)
            .OrderByDescending(path => path)
            .ToList();

        foreach (var directory in directories)
        {
            var files = GetImageFiles(directory).ToList();
            if (files.Count == 0) continue;

            var key = Path.GetFileName(directory);
            _buckets.Add(new DateBucket(key, MakeDateDisplay(key, files.Count), directory));
        }

        foreach (var bucket in _buckets)
            _dateList.Items.Add(bucket);

        if (_dateList.Items.Count > 0)
            _dateList.SelectedIndex = 0;
        else
            LoadSelectedDate();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _previewImage?.Dispose();
            _thumbs.Dispose();
        }

        base.Dispose(disposing);
    }

    private void LoadSelectedDate()
    {
        _captureList.Items.Clear();
        _thumbs.Images.Clear();
        SetPreview(null);

        var bucket = _dateList.SelectedItem as DateBucket;
        if (bucket == null)
        {
            _emptyLabel.Visible = true;
            _detailTitle.Text = "캡쳐를 시작하면 여기에 자동으로 쌓입니다.";
            _detailMeta.Text = _settings.GetHistoryRoot();
            UpdateActionButtons();
            return;
        }

        var files = GetImageFiles(bucket.DirectoryPath)
            .OrderByDescending(File.GetLastWriteTime)
            .ToList();

        for (var i = 0; i < files.Count; i++)
        {
            var path = files[i];
            _thumbs.Images.Add(CreateThumbnail(path));

            var item = new ListViewItem(Path.GetFileNameWithoutExtension(path), i)
            {
                Tag = path,
                ToolTipText = path
            };
            _captureList.Items.Add(item);
        }

        _emptyLabel.Visible = files.Count == 0;
        if (_captureList.Items.Count > 0)
            _captureList.Items[0].Selected = true;

        UpdateActionButtons();
    }

    private void ShowSelectedPreview()
    {
        _selectedPath = _captureList.SelectedItems.Count > 0
            ? _captureList.SelectedItems[0].Tag as string
            : null;

        SetPreview(_selectedPath);
        UpdateActionButtons();
    }

    private void SetPreview(string? path)
    {
        _previewImage?.Dispose();
        _previewImage = null;
        _preview.Image = null;

        if (string.IsNullOrEmpty(path) || !File.Exists(path))
        {
            _detailTitle.Text = "선택된 캡쳐 없음";
            _detailMeta.Text = "";
            return;
        }

        try
        {
            _previewImage = LoadImageCopy(path);
            _preview.Image = _previewImage;

            var info = new FileInfo(path);
            _detailTitle.Text = Path.GetFileName(path);
            _detailMeta.Text =
                $"{info.LastWriteTime:yyyy-MM-dd HH:mm:ss}   " +
                $"{_previewImage.Width}x{_previewImage.Height}   " +
                $"{FormatBytes(info.Length)}";
        }
        catch
        {
            _detailTitle.Text = Path.GetFileName(path);
            _detailMeta.Text = "미리보기를 열 수 없습니다.";
        }
    }

    private static IEnumerable<string> GetImageFiles(string directory)
    {
        if (!Directory.Exists(directory))
            return Enumerable.Empty<string>();

        return Directory.EnumerateFiles(directory)
            .Where(path => ImageExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase));
    }

    private static Image CreateThumbnail(string path)
    {
        using var source = LoadImageCopy(path);
        var thumb = new Bitmap(168, 104);

        using var g = Graphics.FromImage(thumb);
        g.Clear(Color.FromArgb(243, 244, 246));
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;

        var scale = Math.Min(150.0 / source.Width, 86.0 / source.Height);
        var width = Math.Max(1, (int)Math.Round(source.Width * scale));
        var height = Math.Max(1, (int)Math.Round(source.Height * scale));
        var x = (thumb.Width - width) / 2;
        var y = (thumb.Height - height) / 2;

        g.DrawImage(source, new Rectangle(x, y, width, height));

        using var pen = new Pen(Color.FromArgb(229, 231, 235));
        g.DrawRectangle(pen, 0, 0, thumb.Width - 1, thumb.Height - 1);

        return thumb;
    }

    private static Bitmap LoadImageCopy(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var source = Image.FromStream(stream);
        return new Bitmap(source);
    }

    private void OpenSelected()
    {
        if (string.IsNullOrEmpty(_selectedPath) || !File.Exists(_selectedPath))
            return;

        Process.Start(new ProcessStartInfo(_selectedPath) { UseShellExecute = true });
    }

    private void OpenSelectedFolder()
    {
        if (string.IsNullOrEmpty(_selectedPath) || !File.Exists(_selectedPath))
            return;

        Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{_selectedPath}\"") { UseShellExecute = true });
    }

    private void CopySelectedPath()
    {
        if (string.IsNullOrEmpty(_selectedPath) || !File.Exists(_selectedPath))
            return;

        Clipboard.SetText(_selectedPath);
    }

    private void DeleteSelected()
    {
        if (string.IsNullOrEmpty(_selectedPath) || !File.Exists(_selectedPath))
            return;

        var fileName = Path.GetFileName(_selectedPath);
        var result = MessageBox.Show(
            $"{fileName}\n\n이 캡쳐 이미지를 삭제할까요?",
            "캡쳐 삭제",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (result != DialogResult.Yes)
            return;

        File.Delete(_selectedPath);
        Reload();
    }

    private void OpenHistoryRoot()
    {
        var root = _settings.GetHistoryRoot();
        Directory.CreateDirectory(root);
        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{root}\"") { UseShellExecute = true });
    }

    private void UpdateActionButtons()
    {
        var enabled = !string.IsNullOrEmpty(_selectedPath) && File.Exists(_selectedPath);
        _openButton.Enabled = enabled;
        _folderButton.Enabled = enabled;
        _copyButton.Enabled = enabled;
        _deleteButton.Enabled = enabled;
    }

    private static Button MakeButton(string text)
    {
        var button = new Button
        {
            Text = text,
            Size = new Size(92, 32),
            Margin = new Padding(0, 0, 8, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(249, 250, 251),
            ForeColor = Color.FromArgb(31, 41, 55)
        };
        button.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(239, 246, 255);
        return button;
    }

    private static void DrawPanelBorder(object? sender, PaintEventArgs e)
    {
        if (sender is not Control control) return;

        using var pen = new Pen(Color.FromArgb(229, 231, 235));
        e.Graphics.DrawRectangle(pen, 0, 0, control.Width - 1, control.Height - 1);
    }

    private void DateList_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0) return;

        var selected = (e.State & DrawItemState.Selected) != 0;
        var bounds = e.Bounds;
        var bucket = (DateBucket)_dateList.Items[e.Index];

        using var bg = new SolidBrush(selected ? Color.FromArgb(232, 240, 254) : Color.White);
        e.Graphics.FillRectangle(bg, bounds);

        if (selected)
        {
            using var accent = new SolidBrush(Color.FromArgb(37, 99, 235));
            e.Graphics.FillRectangle(accent, bounds.X, bounds.Y + 7, 3, bounds.Height - 14);
        }

        using var textBrush = new SolidBrush(Color.FromArgb(31, 41, 55));
        using var font = new Font("Segoe UI", 9.5f, selected ? FontStyle.Bold : FontStyle.Regular);
        e.Graphics.DrawString(bucket.Display, font, textBrush, bounds.X + 12, bounds.Y + 8);
    }

    private static string MakeDateDisplay(string key, int count)
    {
        if (DateTime.TryParse(key, out var date))
        {
            var label = date.Date == DateTime.Today
                ? "오늘"
                : date.Date == DateTime.Today.AddDays(-1)
                    ? "어제"
                    : date.ToString("yyyy-MM-dd");
            return $"{label}  ({count})";
        }

        return $"{key}  ({count})";
    }

    private static string FormatBytes(long bytes)
    {
        if (bytes >= 1024 * 1024)
            return $"{bytes / 1024d / 1024d:0.0} MB";

        return $"{Math.Max(1, bytes / 1024d):0.0} KB";
    }

    private sealed record DateBucket(string Key, string Display, string DirectoryPath)
    {
        public override string ToString() => Display;
    }
}
