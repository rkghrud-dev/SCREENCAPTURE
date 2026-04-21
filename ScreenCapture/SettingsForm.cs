using Microsoft.Win32;

namespace ScreenCapture;

public class SettingsForm : Form
{
    private readonly ListBox _navList;
    private readonly Panel _contentPanel;
    private readonly List<Panel> _pages = new();

    // --- General ---
    private CheckBox _chkStartWithWindows = null!;
    private CheckBox _chkStartMinimized = null!;

    // --- Capture ---
    private CheckBox _chkIncludeCursor = null!;
    private ComboBox _cmbDelay = null!;
    private CheckBox _chkPlaySound = null!;
    private CheckBox _chkPinToDesktop = null!;
    private CheckBox _chkEnableOcr = null!;

    // --- Save ---
    private TextBox _txtFolder = null!;
    private ComboBox _cmbFormat = null!;
    private TrackBar _trkQuality = null!;
    private Label _lblQuality = null!;
    private TextBox _txtFilePattern = null!;
    private Label _lblPreview = null!;

    // --- Hotkeys ---
    private CheckBox _chkCtrlA = null!, _chkShiftA = null!, _chkAltA = null!;
    private TextBox _txtKeyA = null!;
    private Keys _keyA;
    private ComboBox _cmbActionA = null!;
    private CheckBox _chkCtrlB = null!, _chkShiftB = null!, _chkAltB = null!;
    private TextBox _txtKeyB = null!;
    private Keys _keyB;
    private ComboBox _cmbActionB = null!;
    private CheckBox _chkCtrlC = null!, _chkShiftC = null!, _chkAltC = null!;
    private TextBox _txtKeyC = null!;
    private Keys _keyC;
    private ComboBox _cmbActionC = null!;

    public AppSettings Result { get; private set; }

    public SettingsForm(AppSettings current)
    {
        Result = current.Clone();

        Text = "Options";
        Size = new Size(560, 420);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        TopMost = true;
        Font = new Font("Segoe UI", 9);

        // Left nav list
        _navList = new ListBox
        {
            Location = new Point(0, 0),
            Size = new Size(120, 340),
            BorderStyle = BorderStyle.None,
            BackColor = Color.FromArgb(245, 245, 245),
            Font = new Font("Segoe UI", 10),
            ItemHeight = 32,
            DrawMode = DrawMode.OwnerDrawFixed
        };
        _navList.Items.AddRange(new object[] { "일반", "캡쳐", "저장", "단축키" });
        _navList.DrawItem += NavList_DrawItem;
        _navList.SelectedIndexChanged += (_, _) => ShowPage(_navList.SelectedIndex);
        Controls.Add(_navList);

        // Right content panel
        _contentPanel = new Panel
        {
            Location = new Point(121, 0),
            Size = new Size(423, 340),
            BorderStyle = BorderStyle.None
        };
        Controls.Add(_contentPanel);

        // Separator line
        var sep = new Label
        {
            Location = new Point(120, 0),
            Size = new Size(1, 340),
            BackColor = Color.FromArgb(220, 220, 220)
        };
        Controls.Add(sep);

        // Bottom buttons
        var btnOk = new Button
        {
            Text = "확인(O)",
            Location = new Point(340, 350),
            Size = new Size(90, 30),
            DialogResult = DialogResult.OK
        };
        btnOk.Click += BtnSave_Click;
        Controls.Add(btnOk);

        var btnCancel = new Button
        {
            Text = "취소(C)",
            Location = new Point(440, 350),
            Size = new Size(90, 30),
            DialogResult = DialogResult.Cancel
        };
        Controls.Add(btnCancel);

        AcceptButton = btnOk;
        CancelButton = btnCancel;

        // Build pages
        _pages.Add(BuildGeneralPage());
        _pages.Add(BuildCapturePage());
        _pages.Add(BuildSavePage());
        _pages.Add(BuildHotkeyPage());

        foreach (var p in _pages)
        {
            p.Visible = false;
            _contentPanel.Controls.Add(p);
        }

        LoadFromSettings(current);
        _navList.SelectedIndex = 0;
    }

    // --- Page builders ---

    private Panel BuildGeneralPage()
    {
        var p = MakePage();
        int y = 10;

        AddGroupLabel(p, "프로그램 시작 설정", ref y);

        _chkStartWithWindows = AddCheck(p, "윈도우 시작시 자동 실행(R)", 20, ref y);
        _chkStartMinimized = AddCheck(p, "시작 시 트레이로 최소화", 20, ref y);

        return p;
    }

    private Panel BuildCapturePage()
    {
        var p = MakePage();
        int y = 10;

        AddGroupLabel(p, "캡쳐 설정", ref y);

        _chkIncludeCursor = AddCheck(p, "마우스 커서 포함(M)", 20, ref y);

        y += 5;
        var lblDelay = new Label { Text = "캡쳐 지연 시간:", Location = new Point(20, y + 3), AutoSize = true };
        p.Controls.Add(lblDelay);

        _cmbDelay = new ComboBox
        {
            Location = new Point(140, y),
            Size = new Size(100, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cmbDelay.Items.AddRange(new object[] { "없음", "1초", "2초", "3초", "5초" });
        p.Controls.Add(_cmbDelay);
        y += 35;

        _chkPlaySound = AddCheck(p, "캡쳐 시 효과음 재생", 20, ref y);

        y += 10;
        _chkPinToDesktop = AddCheck(p, "캡쳐 후 스티키로 고정(P)", 20, ref y);
        _chkEnableOcr = AddCheck(p, "OCR 텍스트 인식(O)", 20, ref y);

        return p;
    }

    private Panel BuildSavePage()
    {
        var p = MakePage();
        int y = 10;

        AddGroupLabel(p, "저장 설정", ref y);

        // Folder
        var lblFolder = new Label { Text = "저장 폴더:", Location = new Point(20, y + 3), AutoSize = true };
        p.Controls.Add(lblFolder);
        y += 25;

        _txtFolder = new TextBox
        {
            Location = new Point(20, y),
            Size = new Size(310, 25),
            ReadOnly = true,
            BackColor = SystemColors.Window
        };
        p.Controls.Add(_txtFolder);

        var btnBrowse = new Button
        {
            Text = "...",
            Location = new Point(338, y - 1),
            Size = new Size(50, 27)
        };
        btnBrowse.Click += (_, _) =>
        {
            using var dlg = new FolderBrowserDialog
            {
                Description = "저장 폴더 선택",
                SelectedPath = _txtFolder.Text,
                UseDescriptionForTitle = true
            };
            if (dlg.ShowDialog() == DialogResult.OK)
                _txtFolder.Text = dlg.SelectedPath;
        };
        p.Controls.Add(btnBrowse);
        y += 40;

        // Image format
        var lblFormat = new Label { Text = "이미지 형식:", Location = new Point(20, y + 3), AutoSize = true };
        p.Controls.Add(lblFormat);

        _cmbFormat = new ComboBox
        {
            Location = new Point(140, y),
            Size = new Size(80, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cmbFormat.Items.AddRange(new object[] { "JPG", "PNG", "BMP" });
        _cmbFormat.SelectedIndexChanged += (_, _) =>
        {
            bool isJpg = _cmbFormat.SelectedItem?.ToString() == "JPG";
            _trkQuality.Enabled = isJpg;
            _lblQuality.Enabled = isJpg;
            UpdateFilePreview();
        };
        p.Controls.Add(_cmbFormat);
        y += 35;

        // JPEG quality
        var lblQ = new Label { Text = "JPG 품질:", Location = new Point(20, y + 3), AutoSize = true };
        p.Controls.Add(lblQ);

        _trkQuality = new TrackBar
        {
            Location = new Point(140, y - 5),
            Size = new Size(180, 30),
            Minimum = 10,
            Maximum = 100,
            TickFrequency = 10,
            SmallChange = 5,
            LargeChange = 10
        };
        _trkQuality.ValueChanged += (_, _) => _lblQuality.Text = $"{_trkQuality.Value}%";
        p.Controls.Add(_trkQuality);

        _lblQuality = new Label
        {
            Text = "95%",
            Location = new Point(325, y + 3),
            AutoSize = true,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        p.Controls.Add(_lblQuality);
        y += 40;

        // Filename pattern
        var lblPat = new Label { Text = "파일 이름:", Location = new Point(20, y + 3), AutoSize = true };
        p.Controls.Add(lblPat);

        _txtFilePattern = new TextBox
        {
            Location = new Point(140, y),
            Size = new Size(200, 25)
        };
        _txtFilePattern.TextChanged += (_, _) => UpdateFilePreview();
        p.Controls.Add(_txtFilePattern);
        y += 30;

        _lblPreview = new Label
        {
            Text = "",
            Location = new Point(140, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        p.Controls.Add(_lblPreview);

        return p;
    }

    private static readonly string[] ActionLabels = {
        "클립보드만 (이미지 복사)",
        "파일 저장 + 경로 클립보드",
        "파일만 저장"
    };

    private Panel BuildHotkeyPage()
    {
        var p = MakePage();
        int y = 10;

        AddGroupLabel(p, "단축키 설정", ref y);

        // Hotkey A
        var lblA = new Label
        {
            Text = "단축키 A",
            Location = new Point(20, y),
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold)
        };
        p.Controls.Add(lblA);
        y += 25;

        _chkCtrlA = AddModCheck(p, "Ctrl", 20, y);
        _chkShiftA = AddModCheck(p, "Shift", 80, y);
        _chkAltA = AddModCheck(p, "Alt", 145, y);
        var lblPlus1 = new Label { Text = "+", Location = new Point(192, y + 3), AutoSize = true };
        p.Controls.Add(lblPlus1);
        _txtKeyA = AddKeyBox(p, 210, y);
        y += 30;

        var lblActA = new Label { Text = "캡쳐 결과:", Location = new Point(20, y + 3), AutoSize = true };
        p.Controls.Add(lblActA);
        _cmbActionA = new ComboBox
        {
            Location = new Point(100, y),
            Size = new Size(230, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cmbActionA.Items.AddRange(ActionLabels);
        p.Controls.Add(_cmbActionA);
        y += 45;

        // Hotkey B
        var lblB = new Label
        {
            Text = "단축키 B",
            Location = new Point(20, y),
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold)
        };
        p.Controls.Add(lblB);
        y += 25;

        _chkCtrlB = AddModCheck(p, "Ctrl", 20, y);
        _chkShiftB = AddModCheck(p, "Shift", 80, y);
        _chkAltB = AddModCheck(p, "Alt", 145, y);
        var lblPlus2 = new Label { Text = "+", Location = new Point(192, y + 3), AutoSize = true };
        p.Controls.Add(lblPlus2);
        _txtKeyB = AddKeyBox(p, 210, y);
        y += 30;

        var lblActB = new Label { Text = "캡쳐 결과:", Location = new Point(20, y + 3), AutoSize = true };
        p.Controls.Add(lblActB);
        _cmbActionB = new ComboBox
        {
            Location = new Point(100, y),
            Size = new Size(230, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cmbActionB.Items.AddRange(ActionLabels);
        p.Controls.Add(_cmbActionB);
        y += 45;

        // Hotkey C (Sticky)
        var lblC = new Label
        {
            Text = "단축키 C (스티키)",
            Location = new Point(20, y),
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold)
        };
        p.Controls.Add(lblC);
        y += 25;

        _chkCtrlC = AddModCheck(p, "Ctrl", 20, y);
        _chkShiftC = AddModCheck(p, "Shift", 80, y);
        _chkAltC = AddModCheck(p, "Alt", 145, y);
        var lblPlus3 = new Label { Text = "+", Location = new Point(192, y + 3), AutoSize = true };
        p.Controls.Add(lblPlus3);
        _txtKeyC = AddKeyBox(p, 210, y);
        y += 30;

        var lblActC = new Label { Text = "캡쳐 결과:", Location = new Point(20, y + 3), AutoSize = true };
        p.Controls.Add(lblActC);
        _cmbActionC = new ComboBox
        {
            Location = new Point(100, y),
            Size = new Size(230, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cmbActionC.Items.AddRange(ActionLabels);
        p.Controls.Add(_cmbActionC);

        return p;
    }

    // --- Load / Save ---

    private void LoadFromSettings(AppSettings s)
    {
        // General
        _chkStartWithWindows.Checked = s.StartWithWindows;
        _chkStartMinimized.Checked = s.StartMinimized;

        // Capture
        _chkIncludeCursor.Checked = s.IncludeCursor;
        _cmbDelay.SelectedIndex = s.CaptureDelay switch
        {
            1000 => 1, 2000 => 2, 3000 => 3, 5000 => 4, _ => 0
        };
        _chkPlaySound.Checked = s.PlaySound;
        _chkPinToDesktop.Checked = s.PinToDesktop;
        _chkEnableOcr.Checked = s.EnableOcr;

        // Save
        _txtFolder.Text = s.SaveFolder;
        _cmbFormat.SelectedItem = s.ImageFormat;
        _trkQuality.Value = Math.Clamp(s.JpegQuality, 10, 100);
        _lblQuality.Text = $"{_trkQuality.Value}%";
        _txtFilePattern.Text = s.FileNamePattern;
        UpdateFilePreview();

        // Hotkeys
        _chkCtrlA.Checked = s.HotkeyA.Ctrl;
        _chkShiftA.Checked = s.HotkeyA.Shift;
        _chkAltA.Checked = s.HotkeyA.Alt;
        _keyA = s.HotkeyA.GetKey();
        _txtKeyA.Text = s.HotkeyA.Key;
        _cmbActionA.SelectedIndex = (int)s.ActionA;

        _chkCtrlB.Checked = s.HotkeyB.Ctrl;
        _chkShiftB.Checked = s.HotkeyB.Shift;
        _chkAltB.Checked = s.HotkeyB.Alt;
        _keyB = s.HotkeyB.GetKey();
        _txtKeyB.Text = s.HotkeyB.Key;
        _cmbActionB.SelectedIndex = (int)s.ActionB;

        _chkCtrlC.Checked = s.HotkeyC.Ctrl;
        _chkShiftC.Checked = s.HotkeyC.Shift;
        _chkAltC.Checked = s.HotkeyC.Alt;
        _keyC = s.HotkeyC.GetKey();
        _txtKeyC.Text = s.HotkeyC.Key;
        _cmbActionC.SelectedIndex = (int)s.ActionC;
    }

    private void BtnSave_Click(object? sender, EventArgs e)
    {
        // General
        Result.StartWithWindows = _chkStartWithWindows.Checked;
        Result.StartMinimized = _chkStartMinimized.Checked;

        // Apply startup registry
        ApplyStartup(Result.StartWithWindows);

        // Capture
        Result.IncludeCursor = _chkIncludeCursor.Checked;
        Result.CaptureDelay = _cmbDelay.SelectedIndex switch
        {
            1 => 1000, 2 => 2000, 3 => 3000, 4 => 5000, _ => 0
        };
        Result.PlaySound = _chkPlaySound.Checked;
        Result.PinToDesktop = _chkPinToDesktop.Checked;
        Result.EnableOcr = _chkEnableOcr.Checked;

        // Save
        Result.SaveFolder = _txtFolder.Text;
        Result.ImageFormat = _cmbFormat.SelectedItem?.ToString() ?? "JPG";
        Result.JpegQuality = _trkQuality.Value;
        Result.FileNamePattern = _txtFilePattern.Text;

        // Hotkeys
        Result.HotkeyA = new HotkeyConfig
        {
            Ctrl = _chkCtrlA.Checked,
            Shift = _chkShiftA.Checked,
            Alt = _chkAltA.Checked,
            Key = _keyA.ToString()
        };
        Result.ActionA = (CaptureAction)_cmbActionA.SelectedIndex;
        Result.HotkeyB = new HotkeyConfig
        {
            Ctrl = _chkCtrlB.Checked,
            Shift = _chkShiftB.Checked,
            Alt = _chkAltB.Checked,
            Key = _keyB.ToString()
        };
        Result.ActionB = (CaptureAction)_cmbActionB.SelectedIndex;
        Result.HotkeyC = new HotkeyConfig
        {
            Ctrl = _chkCtrlC.Checked,
            Shift = _chkShiftC.Checked,
            Alt = _chkAltC.Checked,
            Key = _keyC.ToString()
        };
        Result.ActionC = (CaptureAction)_cmbActionC.SelectedIndex;
    }

    private static void ApplyStartup(bool enable)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Run", true);
            if (key == null) return;

            if (enable)
            {
                var exePath = Application.ExecutablePath;
                key.SetValue("ScreenCapture", exePath);
            }
            else
            {
                key.DeleteValue("ScreenCapture", false);
            }
        }
        catch { }
    }

    private void UpdateFilePreview()
    {
        try
        {
            var ext = _cmbFormat.SelectedItem?.ToString()?.ToLower() ?? "jpg";
            var pattern = _txtFilePattern.Text;
            var sample = DateTime.Now.ToString(pattern);
            _lblPreview.Text = $"미리보기: {sample}.{ext}";
        }
        catch
        {
            _lblPreview.Text = "미리보기: (잘못된 패턴)";
        }
    }

    // --- Navigation ---

    private void ShowPage(int index)
    {
        for (int i = 0; i < _pages.Count; i++)
            _pages[i].Visible = i == index;
    }

    private void NavList_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0) return;
        var g = e.Graphics;
        var text = _navList.Items[e.Index].ToString()!;
        bool selected = (e.State & DrawItemState.Selected) != 0;

        var bgColor = selected ? Color.FromArgb(0, 120, 215) : Color.FromArgb(245, 245, 245);
        var fgColor = selected ? Color.White : Color.Black;

        using var bgBrush = new SolidBrush(bgColor);
        g.FillRectangle(bgBrush, e.Bounds);

        using var font = new Font("Segoe UI", 10);
        using var fgBrush = new SolidBrush(fgColor);
        var textRect = new RectangleF(e.Bounds.X + 15, e.Bounds.Y, e.Bounds.Width - 15, e.Bounds.Height);
        var sf = new StringFormat { LineAlignment = StringAlignment.Center };
        g.DrawString(text, font, fgBrush, textRect, sf);
    }

    // --- UI Helpers ---

    private static Panel MakePage()
    {
        return new Panel
        {
            Location = new Point(0, 0),
            Size = new Size(423, 340),
            AutoScroll = true
        };
    }

    private static void AddGroupLabel(Panel p, string text, ref int y)
    {
        var grp = new Label
        {
            Text = text,
            Location = new Point(15, y),
            AutoSize = true,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 90, 180)
        };
        p.Controls.Add(grp);
        y += 30;

        var line = new Label
        {
            Location = new Point(15, y - 5),
            Size = new Size(380, 1),
            BackColor = Color.FromArgb(200, 200, 200)
        };
        p.Controls.Add(line);
        y += 5;
    }

    private static CheckBox AddCheck(Panel p, string text, int x, ref int y)
    {
        var chk = new CheckBox
        {
            Text = text,
            Location = new Point(x, y),
            AutoSize = true
        };
        p.Controls.Add(chk);
        y += 28;
        return chk;
    }

    private static CheckBox AddModCheck(Panel p, string text, int x, int y)
    {
        var chk = new CheckBox
        {
            Text = text,
            Location = new Point(x, y),
            AutoSize = true
        };
        p.Controls.Add(chk);
        return chk;
    }

    private TextBox AddKeyBox(Panel p, int x, int y)
    {
        var txt = new TextBox
        {
            Location = new Point(x, y),
            Size = new Size(120, 25),
            ReadOnly = true,
            BackColor = SystemColors.Window,
            TextAlign = HorizontalAlignment.Center,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        txt.KeyDown += KeyBox_KeyDown;
        txt.PreviewKeyDown += (_, e) => e.IsInputKey = true;
        p.Controls.Add(txt);
        return txt;
    }

    private void KeyBox_KeyDown(object? sender, KeyEventArgs e)
    {
        e.SuppressKeyPress = true;
        e.Handled = true;

        if (e.KeyCode is Keys.ControlKey or Keys.ShiftKey or Keys.Menu)
            return;

        var txt = (TextBox)sender!;
        txt.Text = e.KeyCode.ToString();

        if (txt == _txtKeyA) _keyA = e.KeyCode;
        else if (txt == _txtKeyB) _keyB = e.KeyCode;
        else if (txt == _txtKeyC) _keyC = e.KeyCode;
    }
}
