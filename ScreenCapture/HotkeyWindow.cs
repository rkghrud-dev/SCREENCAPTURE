using System.Runtime.InteropServices;

namespace ScreenCapture;

public class HotkeyWindow : NativeWindow, IDisposable
{
    public const int MOD_ALT = 0x0001;
    public const int MOD_CTRL = 0x0002;
    public const int MOD_SHIFT = 0x0004;
    private const int WM_HOTKEY = 0x0312;

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, int modifiers, int vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private readonly List<int> _ids = new();
    public event Action<int>? HotkeyPressed;

    public HotkeyWindow()
    {
        CreateHandle(new CreateParams());
    }

    public bool Register(int id, int modifiers, Keys key)
    {
        if (RegisterHotKey(Handle, id, modifiers, (int)key))
        {
            _ids.Add(id);
            return true;
        }
        return false;
    }

    public void UnregisterAll()
    {
        foreach (var id in _ids)
            UnregisterHotKey(Handle, id);
        _ids.Clear();
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY)
            HotkeyPressed?.Invoke(m.WParam.ToInt32());
        base.WndProc(ref m);
    }

    public void Dispose()
    {
        UnregisterAll();
        DestroyHandle();
    }
}
