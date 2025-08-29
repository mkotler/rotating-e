using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Configuration;
using System.Windows;
using System.Windows.Threading;

namespace RotatingEOverlay
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
    private OutlookWatcher? _watcher;
    private MainWindow? _overlay;
    private bool _armed;
    private NotifyIcon? _trayIcon;
    private ToolStripMenuItem? _pauseItem;
    private bool _paused;
    private DateTime _lastInputUtc;
    private int _idleThresholdSeconds = 60; // default; overridable via config
    private IntPtr _keyboardHookId = IntPtr.Zero;
    private IntPtr _mouseHookId = IntPtr.Zero;
    private bool _verbose = false; // toggle for extra diagnostics (disabled by default now)
    private DateTime _overlayShownUtc;
    private const int MouseMoveSuppressMs = 200; // don't treat initial synthetic move as real input within this window

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            System.Diagnostics.Debug.WriteLine("[App] Startup");
            _overlay = new MainWindow();
            _overlay.Hide();
            _lastInputUtc = DateTime.UtcNow;

            _watcher = new OutlookWatcher();
            if (_watcher.Initialize())
            {
                System.Diagnostics.Debug.WriteLine("[App] OutlookWatcher initialized");
                _watcher.NewMail += (_, _) =>
                {
                    System.Diagnostics.Debug.WriteLine("[App] NewMail event received");
                    Dispatcher.Invoke(ShowOverlay);
                };
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[App] OutlookWatcher failed to initialize");
            }

            HookGlobalInput();
            LoadConfig();
            InitTray();
        }

        private void LoadConfig()
        {
            try
            {
                var raw = ConfigurationManager.AppSettings["IdleThresholdSeconds"];
                if (!string.IsNullOrWhiteSpace(raw) && int.TryParse(raw, out int v) && v >= 0)
                {
                    _idleThresholdSeconds = v;
                }
                System.Diagnostics.Debug.WriteLine($"[App] Config IdleThresholdSeconds={_idleThresholdSeconds}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[App] LoadConfig failed: " + ex.Message);
            }
        }

        private void InitTray()
        {
            try
            {
                _trayIcon = new NotifyIcon
                {
                    Icon = LoadTrayIcon(),
                    Visible = true,
                    Text = "Disclosure"
                };

                var menu = new ContextMenuStrip();
                _pauseItem = new ToolStripMenuItem("Pause") { CheckOnClick = true, Checked = false };
                _pauseItem.CheckedChanged += (_, _) =>
                {
                    _paused = _pauseItem.Checked;
                    System.Diagnostics.Debug.WriteLine("[App] Pause toggled -> " + _paused);
                    if (_paused)
                    {
                        // Optionally hide current overlay if pausing while visible
                        try { _overlay?.HideUnity(); } catch { }
                    }
                };
                var showItem = new ToolStripMenuItem("Show", null, (_, _) => ForceShowOverlay());
                var exitItem = new ToolStripMenuItem("Exit", null, (_, _) => ExitFromTray());
                menu.Items.Add(_pauseItem);
                menu.Items.Add(showItem);
                menu.Items.Add(new ToolStripSeparator());
                menu.Items.Add(exitItem);
                _trayIcon.ContextMenuStrip = menu;

                // Optional: double-click to test showing overlay manually
                _trayIcon.DoubleClick += (_, _) => Dispatcher.Invoke(ShowOverlay);

                System.Diagnostics.Debug.WriteLine("[App] Tray icon initialized");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[App] Tray init failed: " + ex.Message);
            }
        }

        private void ForceShowOverlay()
        {
            System.Diagnostics.Debug.WriteLine("[App] ForceShowOverlay requested (tray menu)");
            if (_overlay == null)
            {
                System.Diagnostics.Debug.WriteLine("[App] _overlay is null (ForceShowOverlay)");
                return;
            }
            _overlay.ShowUnityWindow();
            _overlay.Activate();
            _armed = true;
            _overlayShownUtc = DateTime.UtcNow;
            System.Diagnostics.Debug.WriteLine("[App] Overlay forcibly shown (tray menu)");
        }

    private Icon LoadTrayIcon()
        {
            try
            {
                var pngPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "e-icon.png");
                if (File.Exists(pngPath))
                {
                    using var bmp = new Bitmap(pngPath);
                    // NOTE: Icon.FromHandle should ideally release handle via DestroyIcon; negligible for single icon lifetime
                    var icon = Icon.FromHandle(bmp.GetHicon());
                    return icon;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[App] LoadTrayIcon error: " + ex.Message);
            }
            return SystemIcons.Application;
        }

        private void ExitFromTray()
        {
            System.Diagnostics.Debug.WriteLine("[App] Exit requested from tray");
            try { _overlay?.HideUnity(kill: true); } catch { }
            try
            {
                if (_trayIcon != null)
                {
                    _trayIcon.Visible = false;
                    _trayIcon.Dispose();
                    _trayIcon = null;
                }
            }
            catch { }
            Shutdown();
        }

        private void ShowOverlay()
        {
            System.Diagnostics.Debug.WriteLine("[App] ShowOverlay requested");
            if (_verbose) LogState("ShowOverlay.enter");
            if (_paused)
            {
                System.Diagnostics.Debug.WriteLine("[App] Suppressed (paused)");
                return;
            }
            bool idle = IsUserIdle();
            if (!idle)
            {
                System.Diagnostics.Debug.WriteLine("[App] Suppressed (user active / not idle)");
                return;
            }
            if (_armed) // already showing for prior mail
            {
                System.Diagnostics.Debug.WriteLine("[App] Suppressed (already armed/showing)");
                return;
            }
            if (_overlay == null)
            {
                System.Diagnostics.Debug.WriteLine("[App] _overlay is null");
                return;
            }
            _overlay.ShowUnityWindow();
            _overlay.Activate();
            _armed = true;
            _overlayShownUtc = DateTime.UtcNow;
            System.Diagnostics.Debug.WriteLine("[App] Overlay shown and armed");
            if (_verbose) LogState("ShowOverlay.afterShow");
        }

        private void HideOverlay()
        {
            if (_overlay is { IsVisible: true })
            {
                _overlay.HideUnity(kill: true);
                System.Diagnostics.Debug.WriteLine("[App] Overlay hidden");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[App] HideOverlay called but overlay not visible");
            }
        }

        #region Global Input
        private delegate IntPtr LowLevelProc(int nCode, IntPtr wParam, IntPtr lParam);
        private LowLevelProc? _proc;

        private void HookGlobalInput()
        {
            _proc = HookCallback;
            using var curProcess = System.Diagnostics.Process.GetCurrentProcess();
            using var curModule = curProcess.MainModule!;
            _keyboardHookId = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(curModule.ModuleName), 0);
            _mouseHookId = SetWindowsHookEx(WH_MOUSE_LL, _proc, GetModuleHandle(curModule.ModuleName), 0);
            System.Diagnostics.Debug.WriteLine("[App] Hooks installed keyboard=" + _keyboardHookId + " mouse=" + _mouseHookId);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try
            {
                if (_keyboardHookId != IntPtr.Zero) UnhookWindowsHookEx(_keyboardHookId);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[App] Unhook keyboard failed: " + ex.Message); }
            try
            {
                if (_mouseHookId != IntPtr.Zero) UnhookWindowsHookEx(_mouseHookId);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[App] Unhook mouse failed: " + ex.Message); }
            _watcher?.Dispose();
            if (_trayIcon != null)
            {
                try
                {
                    _trayIcon.Visible = false;
                    _trayIcon.Dispose();
                }
                catch { }
                _trayIcon = null;
            }
            try { _overlay?.HideUnity(kill: true); } catch { }
            base.OnExit(e);
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                bool isKeyboardMsg = IsKeyboardMessage((int)wParam);
                bool isMouseMsg = !isKeyboardMsg; // since only two hooks share callback
                bool injected = false;
                int msg = (int)wParam;
                if (isKeyboardMsg)
                {
                    try
                    {
                        KBDLLHOOKSTRUCT kb = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                        injected = (kb.flags & LLKHF_INJECTED) != 0; // synthetic keyboard
                    }
                    catch { }
                }
                else if (isMouseMsg)
                {
                    try
                    {
                        MSLLHOOKSTRUCT ms = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                        injected = (ms.flags & (LLMHF_INJECTED | LLMHF_LOWER_IL_INJECTED)) != 0; // synthetic mouse
                    }
                    catch { }
                }

                TimeSpan sinceShown = _overlayShownUtc == default ? TimeSpan.MaxValue : DateTime.UtcNow - _overlayShownUtc;
                bool isMeaningful = IsMeaningfulUserAction(msg, sinceShown);

                if (_verbose)
                {
                    System.Diagnostics.Debug.WriteLine($"[App] Input event msg=0x{msg:X} injected={injected} meaningful={isMeaningful} armed={_armed} sinceShownMs={sinceShown.TotalMilliseconds:F0}");
                }

                if (injected)
                {
                    if (_verbose) System.Diagnostics.Debug.WriteLine("[App] Ignored injected event");
                    return CallNextHookEx(_keyboardHookId, nCode, wParam, lParam);
                }

                if (!isMeaningful)
                {
                    if (_verbose) System.Diagnostics.Debug.WriteLine("[App] Ignored non-meaningful (e.g., early mouse move)");
                    return CallNextHookEx(_keyboardHookId, nCode, wParam, lParam);
                }

                _lastInputUtc = DateTime.UtcNow; // track input time
                if (_armed)
                {
                    _armed = false;
                    System.Diagnostics.Debug.WriteLine("[App] Meaningful input -> hiding overlay & Unity");
                    Dispatcher.Invoke(() =>
                    {
                        try
                        {
                            HideOverlay();
                            _overlay?.HideUnity(kill: true);
                            if (_verbose) LogState("AfterHideOnInput");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("[App] Exception hiding on input: " + ex.Message);
                        }
                    });
                }
                else if (_verbose)
                {
                    System.Diagnostics.Debug.WriteLine("[App] Meaningful input while not armed (no hide)");
                }
            }
            // Use keyboard hook handle for chaining; if null fallback to IntPtr.Zero
            return CallNextHookEx(_keyboardHookId, nCode, wParam, lParam);
        }

        private bool IsUserIdle()
        {
            try
            {
                int thresholdMs = _idleThresholdSeconds * 1000;
                if (thresholdMs <= 0) return true; // allow forced immediate show for testing
                // Hook-based idle
                double hookIdleMs = (DateTime.UtcNow - _lastInputUtc).TotalMilliseconds;
                // Win32 last input
                uint win32IdleMs = GetLastInputMilliseconds();
                bool idle = hookIdleMs >= thresholdMs || win32IdleMs >= thresholdMs;
                if (_verbose)
                    System.Diagnostics.Debug.WriteLine($"[App] Idle check hook={hookIdleMs:F0}ms win32={win32IdleMs}ms thresh={thresholdMs} -> {idle}");
                return idle;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[App] IsUserIdle exception: " + ex.Message);
                return true; // fail open
            }
        }

        private static uint GetLastInputMilliseconds()
        {
            if (GetLastInputInfo(out LASTINPUTINFO info))
            {
                uint tick = GetTickCount();
                if (info.dwTime <= tick) return tick - info.dwTime;
            }
            return 0;
        }


        private void LogState(string tag)
        {
            try
            {
                bool overlayVisible = _overlay?.IsVisible ?? false;
                double lastInputMs = (DateTime.UtcNow - _lastInputUtc).TotalMilliseconds;
                System.Diagnostics.Debug.WriteLine($"[App] State[{tag}] armed={_armed} paused={_paused} overlayVisible={overlayVisible} lastInputAgeMs={lastInputMs:F0} idleThresholdSec={_idleThresholdSeconds}");
            }
            catch { }
        }

        private static bool IsKeyboardMessage(int msg)
        {
            return msg == WM_KEYDOWN || msg == WM_KEYUP || msg == WM_SYSKEYDOWN || msg == WM_SYSKEYUP;
        }

        private bool IsMeaningfulUserAction(int msg, TimeSpan sinceShown)
        {
            // Keyboard key press
            if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN) return true;
            // Mouse button / wheel events
            switch (msg)
            {
                case WM_LBUTTONDOWN:
                case WM_RBUTTONDOWN:
                case WM_MBUTTONDOWN:
                case WM_XBUTTONDOWN:
                case WM_MOUSEWHEEL:
                case WM_MOUSEHWHEEL:
                    return true;
                case WM_MOUSEMOVE:
                    // Treat as meaningful only if overlay has been visible beyond suppression window
                    return sinceShown.TotalMilliseconds > MouseMoveSuppressMs;
            }
            return false; // ignore other events (keyup, button up, etc.)
        }

        // Low-level hook structs
        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode; public uint scanCode; public uint flags; public uint time; public UIntPtr dwExtraInfo;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt; public uint mouseData; public uint flags; public uint time; public UIntPtr dwExtraInfo;
        }
        [StructLayout(LayoutKind.Sequential)] private struct POINT { public int x; public int y; }

        // Message & flag constants
        private const int WM_MOUSEMOVE = 0x0200;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_MBUTTONDOWN = 0x0207;
        private const int WM_XBUTTONDOWN = 0x020B;
        private const int WM_MOUSEWHEEL = 0x020A;
        private const int WM_MOUSEHWHEEL = 0x020E;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;
        private const uint LLKHF_INJECTED = 0x00000010;
        private const uint LLMHF_INJECTED = 0x00000001;
        private const uint LLMHF_LOWER_IL_INJECTED = 0x00000002;

        private const int WH_KEYBOARD_LL = 13;
        private const int WH_MOUSE_LL = 14; // can add mouse if desired

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

    // Idle detection
    [StructLayout(LayoutKind.Sequential)]
    private struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
    [DllImport("user32.dll")] private static extern bool GetLastInputInfo(out LASTINPUTINFO plii);
    [DllImport("kernel32.dll")] private static extern uint GetTickCount();

    // (Removed fullscreen detection code as no longer used)
        #endregion
    }
}

