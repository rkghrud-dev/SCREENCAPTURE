using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace ScreenCapture;

public class AiToolsForm : Form
{
    private const string DefaultImageEditModel = "Qwen/Qwen-Image-Edit";
    private const string FluxKontextImageEditModel = "black-forest-labs/FLUX.1-Kontext-dev";

    private readonly CaptureInfo _info;
    private readonly string _saveFolder;
    private readonly Bitmap _originalImage;
    private Bitmap _workingImage;
    private readonly List<Bitmap> _history = new();
    private const int MaxHistory = 2;

    private PictureBox _beforePreview = null!;
    private PictureBox _afterPreview = null!;
    private Label _beforeLabel = null!;
    private Label _afterLabel = null!;
    private TextBox _promptBox = null!;
    private TextBox _widthBox = null!;
    private TextBox _heightBox = null!;
    private ComboBox _modelBox = null!;
    private Label _statusLabel = null!;
    private Button _undoButton = null!;
    private readonly List<Button> _actionButtons = new();

    public event Action<Bitmap, CaptureInfo>? NewStickyRequested;

    public AiToolsForm(Bitmap image, CaptureInfo info, string saveFolder)
    {
        _originalImage = new Bitmap(image);
        _workingImage = new Bitmap(image);
        _info = info;
        _saveFolder = saveFolder;

        Text = "AI 도구";
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(920, 720);
        MinimumSize = new Size(760, 600);
        Font = new Font("Segoe UI", 9);
        ShowIcon = false;
        ShowInTaskbar = false;

        BuildUi();
        UpdatePreviews();
    }

    private void BuildUi()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Padding = new Padding(12)
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
        Controls.Add(root);

        var previewPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            BackColor = Color.FromArgb(38, 38, 38),
            Padding = new Padding(6)
        };
        previewPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        previewPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        previewPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        previewPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        root.Controls.Add(previewPanel, 0, 0);

        _beforeLabel = new Label
        {
            Text = "▶ Before (원본)",
            ForeColor = Color.FromArgb(180, 180, 180),
            AutoSize = true,
            Padding = new Padding(4, 2, 0, 2),
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
        };
        previewPanel.Controls.Add(_beforeLabel, 0, 0);

        _beforePreview = new PictureBox
        {
            Dock = DockStyle.Fill,
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.FromArgb(24, 24, 24),
            Margin = new Padding(4)
        };
        previewPanel.Controls.Add(_beforePreview, 0, 1);

        _afterLabel = new Label
        {
            Text = "▶ After (변경 후)",
            ForeColor = Color.FromArgb(120, 200, 120),
            AutoSize = true,
            Padding = new Padding(4, 2, 0, 2),
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
        };
        previewPanel.Controls.Add(_afterLabel, 0, 2);

        _afterPreview = new PictureBox
        {
            Dock = DockStyle.Fill,
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.FromArgb(24, 24, 24),
            Margin = new Padding(4)
        };
        previewPanel.Controls.Add(_afterPreview, 0, 3);

        var right = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 9,
            Padding = new Padding(12, 0, 0, 0)
        };
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        right.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(right, 1, 0);

        right.Controls.Add(new Label
        {
            AutoSize = true,
            Text = "이미지 변경 프롬프트",
            Font = new Font(Font, FontStyle.Bold)
        }, 0, 0);

        _promptBox = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Text = "Make this image cleaner and more detailed while preserving the original layout."
        };
        right.Controls.Add(_promptBox, 0, 1);

        right.Controls.Add(new Label
        {
            AutoSize = true,
            Padding = new Padding(0, 8, 0, 0),
            Text = "모델",
            Font = new Font(Font, FontStyle.Bold)
        }, 0, 2);

        _modelBox = new ComboBox
        {
            Dock = DockStyle.Top,
            DropDownStyle = ComboBoxStyle.DropDown
        };
        _modelBox.Items.Add(DefaultImageEditModel);
        _modelBox.Items.Add(FluxKontextImageEditModel);
        _modelBox.SelectedIndex = 0;
        right.Controls.Add(_modelBox, 0, 3);

        var sizePanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Padding = new Padding(0, 8, 0, 0),
            WrapContents = false
        };
        sizePanel.Controls.Add(new Label
        {
            Text = "크기",
            AutoSize = true,
            Width = 38,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(0, 6, 0, 0)
        });
        _widthBox = new TextBox { Width = 70, Text = _workingImage.Width.ToString() };
        _heightBox = new TextBox { Width = 70, Text = _workingImage.Height.ToString() };
        sizePanel.Controls.Add(_widthBox);
        sizePanel.Controls.Add(new Label
        {
            Text = "x",
            AutoSize = true,
            Padding = new Padding(4, 6, 4, 0)
        });
        sizePanel.Controls.Add(_heightBox);
        sizePanel.Controls.Add(new Label
        {
            Text = "px",
            AutoSize = true,
            Padding = new Padding(4, 6, 0, 0)
        });
        right.Controls.Add(sizePanel, 0, 4);

        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            RowCount = 5,
            Padding = new Padding(0, 10, 0, 0),
            AutoSize = true
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        right.Controls.Add(grid, 0, 5);

        AddActionButton(grid, "이미지 변경 (미구현)", 0, 0, EditImageWithAi);
        AddActionButton(grid, "배경 제거", 1, 0, RemoveBackground);
        AddActionButton(grid, "ChatGPT로 보내기", 0, 1, OpenInChatGpt);
        AddActionButton(grid, "번역하기", 1, 1, TranslateImageText);
        AddActionButton(grid, "크기 변경", 0, 2, ResizeImage);
        AddActionButton(grid, "아이콘 저장", 1, 2, SaveIcon);
        AddActionButton(grid, "결과 저장", 0, 3, SaveImageAs);
        AddActionButton(grid, "닫기", 1, 3, Close);

        _undoButton = new Button
        {
            Text = "↩ 돌아가기",
            Dock = DockStyle.Fill,
            Height = 38,
            Enabled = false,
            Margin = new Padding(0, 0, 0, 8)
        };
        _undoButton.Click += (_, _) => Undo();
        _actionButtons.Add(_undoButton);
        grid.Controls.Add(_undoButton, 0, 4);
        grid.SetColumnSpan(_undoButton, 2);

        var historyLabel = new Label
        {
            AutoSize = true,
            ForeColor = Color.FromArgb(100, 100, 100),
            Padding = new Padding(0, 4, 0, 0)
        };
        right.Controls.Add(historyLabel, 0, 6);
        _historyLabel = historyLabel;

        _statusLabel = new Label
        {
            Dock = DockStyle.Top,
            AutoSize = false,
            Height = 48,
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Color.FromArgb(70, 70, 70),
            Text = "준비됨"
        };
        right.Controls.Add(_statusLabel, 0, 8);
    }

    private Label _historyLabel = null!;

    private void AddActionButton(TableLayoutPanel grid, string text, int col, int row, Action action)
    {
        var button = new Button
        {
            Text = text,
            Dock = DockStyle.Fill,
            Height = 36,
            Margin = new Padding(0, 0, col == 0 ? 6 : 0, 6)
        };
        button.Click += (_, _) => action();
        _actionButtons.Add(button);
        grid.Controls.Add(button, col, row);
    }

    private void UpdatePreviews()
    {
        var oldBefore = _beforePreview.Image;
        _beforePreview.Image = _history.Count > 0
            ? new Bitmap(_history[^1])
            : new Bitmap(_originalImage);
        oldBefore?.Dispose();

        var oldAfter = _afterPreview.Image;
        _afterPreview.Image = new Bitmap(_workingImage);
        oldAfter?.Dispose();

        _beforeLabel.Text = _history.Count > 0
            ? $"▶ Before ({_history.Count}차 변형)"
            : "▶ Before (원본)";
        _afterLabel.Text = _history.Count > 0
            ? $"▶ After ({_history.Count + 1}차 변형)"
            : "▶ After (변경 후)";

        _undoButton.Enabled = _history.Count > 0;
        _historyLabel.Text = _history.Count > 0
            ? $"히스토리: {_history.Count}개 (최대 {MaxHistory}개 보관)"
            : "";
    }

    private void PushHistory()
    {
        if (_history.Count >= MaxHistory)
        {
            _history[0].Dispose();
            _history.RemoveAt(0);
        }
        _history.Add(new Bitmap(_workingImage));
    }

    private void Undo()
    {
        if (_history.Count == 0) return;

        _workingImage.Dispose();
        _workingImage = _history[^1];
        _history.RemoveAt(_history.Count - 1);

        UpdatePreviews();
        SetStatus($"돌아가기 완료 (히스토리: {_history.Count}개 남음)");
    }

    private async void EditImageWithAi()
    {
        await Task.CompletedTask;
        MessageBox.Show(
            "AI 이미지 편집은 아직 미구현입니다.\n현재는 스티키 우클릭 메뉴의 부가기능에서 ChatGPT로 보내기와 배경 제거만 사용할 수 있습니다.",
            "AI 이미지 편집", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private async void RemoveBackground()
    {
        await RunBusyAsync("로컬 rembg 배경 제거 실행 중...", async () =>
        {
            using var result = await RunRemoveBackgroundAsync();
            PublishResult(result, "nobg", "배경 제거 결과");
        });
    }

    private async void OpenInChatGpt()
    {
        try
        {
            using var clipboardImage = new Bitmap(_workingImage);
            Clipboard.SetDataObject(clipboardImage, true);

            Process.Start(new ProcessStartInfo("https://chatgpt.com/") { UseShellExecute = true });
            await Task.Delay(3500);
            SendKeys.SendWait("^v");
            SetStatus("ChatGPT로 이미지를 보냈습니다. 붙지 않았다면 ChatGPT 입력창에서 Ctrl+V를 눌러주세요.");
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
        await RunBusyAsync("OCR 텍스트 확인 중...", async () =>
        {
            var text = _info.OcrText ?? _info.ClipboardText;
            if (string.IsNullOrWhiteSpace(text))
            {
                text = await OcrHelper.ExtractTextAsync(_workingImage);
                _info.OcrText = text;
            }

            if (string.IsNullOrWhiteSpace(text))
                throw new InvalidOperationException("번역할 텍스트를 찾지 못했습니다.");

            text = text.Trim();
            Clipboard.SetText(text);

            var urlText = text.Length > 4500 ? text[..4500] : text;
            var url = "https://translate.google.com/?sl=auto&tl=ko&text=" +
                      Uri.EscapeDataString(urlText) +
                      "&op=translate";
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            SetStatus("번역 페이지를 열었습니다. 원문도 클립보드에 복사했습니다.");
        });
    }

    private void SaveImageAs()
    {
        using var dlg = new SaveFileDialog
        {
            Filter = "PNG|*.png|JPG|*.jpg|BMP|*.bmp",
            FileName = $"ai_result_{DateTime.Now:yyyyMMddHHmmss}"
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        SaveBitmap(_workingImage, dlg.FileName);
        SetStatus($"저장됨: {dlg.FileName}");
    }

    private void ResizeImage()
    {
        if (!TryReadSize(out var width, out var height))
            return;

        using var result = ResizeBitmap(_workingImage, width, height);
        PublishResult(result, $"resize_{width}x{height}", $"크기 변경 {width}x{height}");
    }

    private void SaveIcon()
    {
        using var dlg = new SaveFileDialog
        {
            Filter = "Icon|*.ico|PNG|*.png",
            FileName = $"icon_{DateTime.Now:yyyyMMddHHmmss}"
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        using var iconImage = CreateIconImage(_workingImage, 256);
        if (Path.GetExtension(dlg.FileName).Equals(".ico", StringComparison.OrdinalIgnoreCase))
            SavePngIco(iconImage, dlg.FileName);
        else
            iconImage.Save(dlg.FileName, ImageFormat.Png);

        SetStatus($"아이콘 저장됨: {dlg.FileName}");
    }

    private void PublishResult(Bitmap result, string suffix, string title)
    {
        var ts = DateTime.Now.ToString("yyyyMMddHHmmss");
        Directory.CreateDirectory(_saveFolder);
        var path = Path.Combine(_saveFolder, $"{ts}_{suffix}.png");
        result.Save(path, ImageFormat.Png);

        PushHistory();
        _workingImage.Dispose();
        _workingImage = new Bitmap(result);
        UpdatePreviews();

        var newInfo = new CaptureInfo
        {
            ProcessName = _info.ProcessName,
            WindowTitle = title,
            Url = _info.Url,
            FilePath = _info.FilePath,
            ClipboardText = _info.ClipboardText,
            CapturedImagePath = path,
            CapturedRegion = _info.CapturedRegion,
            SourceHwnd = _info.SourceHwnd,
            SourceKind = _info.SourceKind,
            SourceAnchor = _info.SourceAnchor
        };

        NewStickyRequested?.Invoke(new Bitmap(result), newInfo);
        SetStatus($"완료됨: {path}");
    }

    private bool TryReadSize(out int width, out int height)
    {
        width = 0;
        height = 0;

        if (!int.TryParse(_widthBox.Text.Trim(), out width) ||
            !int.TryParse(_heightBox.Text.Trim(), out height) ||
            width < 1 || height < 1 || width > 8192 || height > 8192)
        {
            MessageBox.Show("크기는 1부터 8192 사이의 픽셀 값으로 입력해주세요.",
                "크기 변경", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return false;
        }

        return true;
    }

    private async Task<Bitmap> RunHuggingFaceImageEditAsync(string prompt, string model)
    {
        var tokenPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Desktop", "key", "huggingface.txt");
        if (!File.Exists(tokenPath))
            throw new FileNotFoundException($"HuggingFace 토큰 파일을 찾지 못했습니다: {tokenPath}");

        var token = (await File.ReadAllTextAsync(tokenPath)).Trim();
        if (string.IsNullOrWhiteSpace(token))
            throw new InvalidOperationException("HuggingFace 토큰 파일이 비어 있습니다.");

        var tmpDir = Path.Combine(Path.GetTempPath(), "ScreenCaptureAi");
        Directory.CreateDirectory(tmpDir);
        var stamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
        var inputPath = Path.Combine(tmpDir, $"input_{stamp}.png");
        var outputPath = Path.Combine(tmpDir, $"output_{stamp}.png");
        var scriptPath = Path.Combine(tmpDir, $"hf_edit_{stamp}.py");

        _workingImage.Save(inputPath, ImageFormat.Png);
        await File.WriteAllTextAsync(scriptPath, """
import sys, re
from huggingface_hub import InferenceClient

token, input_path, output_path, prompt, model = sys.argv[1:6]
image_client = InferenceClient(provider="fal-ai", api_key=token)
text_client = InferenceClient(api_key=token)

# Auto-translate Korean prompt to English
if re.search(r'[\uac00-\ud7a3]', prompt):
    try:
        translated = text_client.translation(prompt, model="Helsinki-NLP/opus-mt-ko-en")
        if isinstance(translated, str) and translated.strip():
            prompt = translated.strip()
        elif hasattr(translated, 'translation_text'):
            prompt = translated.translation_text.strip()
    except Exception:
        pass

with open(input_path, "rb") as f:
    img_bytes = f.read()

try:
    result = image_client.image_to_image(img_bytes, prompt=prompt, model=model)
except Exception as e:
    message = str(e).strip()
    if "Cannot POST" in message or "404" in message:
        raise RuntimeError(
            "HuggingFace 이미지 변경 API 호출이 실패했습니다. "
            f"모델 '{model}'이 현재 Inference Provider에서 image-to-image로 제공되지 않을 수 있습니다. "
            "모델을 Qwen/Qwen-Image-Edit 또는 black-forest-labs/FLUX.1-Kontext-dev로 바꾸고, "
            "토큰에 Inference Providers 권한이 있는지 확인해주세요.\n\n"
            f"원본 오류: {message}"
        ) from e
    raise
result.save(output_path)
""");

        try
        {
            await RunPythonAsync(scriptPath, token, inputPath, outputPath, prompt, model);
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("HuggingFace 결과 이미지가 생성되지 않았습니다.");

            using var fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read);
            return new Bitmap(fs);
        }
        finally
        {
            TryDelete(inputPath);
            TryDelete(outputPath);
            TryDelete(scriptPath);
        }
    }

    private static string NormalizeHuggingFaceImageEditModel(string model)
    {
        model = model.Trim();
        if (string.IsNullOrWhiteSpace(model))
            return DefaultImageEditModel;

        return model switch
        {
            "diffusers/instruct-pix2pix" => DefaultImageEditModel,
            "instruct-pix2pix" => DefaultImageEditModel,
            "black-forest-labs/FLUX.1-schnell" => DefaultImageEditModel,
            _ => model
        };
    }

    private async Task<Bitmap> RunRemoveBackgroundAsync()
    {
        var tmpDir = Path.Combine(Path.GetTempPath(), "ScreenCaptureAi");
        Directory.CreateDirectory(tmpDir);
        var stamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
        var inputPath = Path.Combine(tmpDir, $"rembg_input_{stamp}.png");
        var outputPath = Path.Combine(tmpDir, $"rembg_output_{stamp}.png");
        var scriptPath = Path.Combine(tmpDir, $"rembg_{stamp}.py");

        _workingImage.Save(inputPath, ImageFormat.Png);
        await File.WriteAllTextAsync(scriptPath, """
import sys
from rembg import remove
from PIL import Image

input_path, output_path = sys.argv[1:3]
result = remove(Image.open(input_path))
result.save(output_path)
""");

        try
        {
            await RunPythonAsync(scriptPath, inputPath, outputPath);
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("배경 제거 결과 이미지가 생성되지 않았습니다.");

            using var fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read);
            return new Bitmap(fs);
        }
        finally
        {
            TryDelete(inputPath);
            TryDelete(outputPath);
            TryDelete(scriptPath);
        }
    }

    private static async Task RunPythonAsync(string scriptPath, params string[] args)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "python",
            CreateNoWindow = true,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        psi.ArgumentList.Add(scriptPath);
        foreach (var arg in args)
            psi.ArgumentList.Add(arg);

        using var proc = Process.Start(psi);
        if (proc == null)
            throw new FileNotFoundException("python");

        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();
        await proc.WaitForExitAsync();

        if (proc.ExitCode != 0)
        {
            var stdout = await stdoutTask;
            var stderr = await stderrTask;
            var message = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
            throw new InvalidOperationException(message.Trim());
        }
    }

    private async Task RunBusyAsync(string status, Func<Task> action)
    {
        SetBusy(true, status);
        try
        {
            await action();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "AI 도구", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            SetStatus("실패");
        }
        finally
        {
            SetBusy(false, _statusLabel.Text);
        }
    }

    private void SetBusy(bool busy, string status)
    {
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        foreach (var button in _actionButtons)
            button.Enabled = !busy;
        if (!busy)
            _undoButton.Enabled = _history.Count > 0;
        SetStatus(status);
    }

    private void SetStatus(string status) => _statusLabel.Text = status;

    private static void SaveBitmap(Bitmap bitmap, string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        var format = ext switch
        {
            ".jpg" or ".jpeg" => ImageFormat.Jpeg,
            ".bmp" => ImageFormat.Bmp,
            _ => ImageFormat.Png
        };
        bitmap.Save(path, format);
    }

    private static Bitmap ResizeBitmap(Bitmap source, int width, int height)
    {
        var result = new Bitmap(width, height);
        result.SetResolution(source.HorizontalResolution, source.VerticalResolution);
        using var g = Graphics.FromImage(result);
        g.CompositingQuality = CompositingQuality.HighQuality;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        g.SmoothingMode = SmoothingMode.HighQuality;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.Clear(Color.Transparent);
        g.DrawImage(source, new Rectangle(0, 0, width, height));
        return result;
    }

    private static Bitmap CreateIconImage(Bitmap source, int size)
    {
        var result = new Bitmap(size, size, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(result);
        g.CompositingQuality = CompositingQuality.HighQuality;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        g.SmoothingMode = SmoothingMode.HighQuality;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.Clear(Color.Transparent);

        var scale = Math.Min(size / (double)source.Width, size / (double)source.Height);
        var width = (int)Math.Round(source.Width * scale);
        var height = (int)Math.Round(source.Height * scale);
        var x = (size - width) / 2;
        var y = (size - height) / 2;
        g.DrawImage(source, new Rectangle(x, y, width, height));
        return result;
    }

    private static void SavePngIco(Bitmap image, string path)
    {
        using var png = new MemoryStream();
        image.Save(png, ImageFormat.Png);
        var bytes = png.ToArray();

        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        using var writer = new BinaryWriter(fs);
        writer.Write((ushort)0);
        writer.Write((ushort)1);
        writer.Write((ushort)1);
        writer.Write((byte)0);
        writer.Write((byte)0);
        writer.Write((byte)0);
        writer.Write((byte)0);
        writer.Write((ushort)1);
        writer.Write((ushort)32);
        writer.Write(bytes.Length);
        writer.Write(22);
        writer.Write(bytes);
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _beforePreview.Image?.Dispose();
            _afterPreview.Image?.Dispose();
            _originalImage.Dispose();
            _workingImage.Dispose();
            foreach (var h in _history) h.Dispose();
            _history.Clear();
        }
        base.Dispose(disposing);
    }
}
