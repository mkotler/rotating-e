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
                var exitItem = new ToolStripMenuItem("Exit", null, (_, _) => ExitFromTray());
                menu.Items.Add(_pauseItem);
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
            if (_paused)
            {
                System.Diagnostics.Debug.WriteLine("[App] Suppressed (paused)");
                return;
            }
            if (!IsUserIdle())
            {
                System.Diagnostics.Debug.WriteLine("[App] Suppressed (user active / not idle)");
                return;
            }
            if (IsForegroundFullscreen())
            {
                System.Diagnostics.Debug.WriteLine("[App] Suppressed (foreground fullscreen)");
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
            System.Diagnostics.Debug.WriteLine("[App] Overlay shown and armed");
        }

        private void HideOverlay()
        {
            if (_overlay is { IsVisible: true })
            {
                _overlay.HideUnity(kill: false);
                System.Diagnostics.Debug.WriteLine("[App] Overlay hidden");
            }
        }

        #region Global Input
        private delegate IntPtr LowLevelProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static IntPtr _hookId = IntPtr.Zero;
        private LowLevelProc? _proc;

        private void HookGlobalInput()
        {
            _proc = HookCallback;
            _hookId = SetHook(_proc);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            UnhookWindowsHookEx(_hookId);
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
                _lastInputUtc = DateTime.UtcNow; // track input time always
                if (_armed)
                {
                    _armed = false;
                    System.Diagnostics.Debug.WriteLine("[App] Input -> hiding overlay & Unity");
                    Dispatcher.Invoke(() =>
                    {
                        try
                        {
                            HideOverlay();
                            _overlay?.HideUnity();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("[App] Exception hiding on input: " + ex.Message);
                        }
                    });
                }
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
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

        private bool IsForegroundFullscreen()
        {
            try
            {
                IntPtr fg = GetForegroundWindow();
                if (fg == IntPtr.Zero) return false;
                if (!GetWindowRect(fg, out RECT rect)) return false;
                IntPtr monitor = MonitorFromWindow(fg, MONITOR_DEFAULTTONEAREST);
                if (monitor == IntPtr.Zero) return false;
                MONITORINFO mi = new MONITORINFO { cbSize = (uint)Marshal.SizeOf<MONITORINFO>() };
                if (!GetMonitorInfo(monitor, ref mi)) return false;
                int w = rect.Right - rect.Left;
                int h = rect.Bottom - rect.Top;
                int mw = (int)(mi.rcMonitor.Right - mi.rcMonitor.Left);
                int mh = (int)(mi.rcMonitor.Bottom - mi.rcMonitor.Top);
                bool fullscreen = Math.Abs(w - mw) <= 2 && Math.Abs(h - mh) <= 2; // tolerances
                if (fullscreen)
                    System.Diagnostics.Debug.WriteLine($"[App] Foreground fullscreen detected size={w}x{h} monitor={mw}x{mh}");
                return fullscreen;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[App] IsForegroundFullscreen exception: " + ex.Message);
                return false;
            }
        }

        private static IntPtr SetHook(LowLevelProc proc)
        {
            using var curProcess = System.Diagnostics.Process.GetCurrentProcess();
            using var curModule = curProcess.MainModule!;
            IntPtr h = SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            IntPtr hm = SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            System.Diagnostics.Debug.WriteLine("[App] Keyboard hook=" + h + " mouse hook=" + hm);
            return h; // store only keyboard for unhook simplicity
        }

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

    // Fullscreen detection
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", SetLastError = true)] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);
    private const uint MONITOR_DEFAULTTONEAREST = 2;
    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [StructLayout(LayoutKind.Sequential)] private struct MONITORINFO { public uint cbSize; public RECT rcMonitor; public RECT rcWork; public uint dwFlags; }
        #endregion
    }
}

