using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace RotatingEOverlay
{
    public partial class MainWindow : Window
    {
        private Process? _unityProcess;
        private bool _unityRunning;
        private IntPtr _unityHwnd = IntPtr.Zero;
        private System.Windows.Threading.DispatcherTimer? _overlayFocusTimer;

        public MainWindow()
        {
            InitializeComponent();
            // Don't auto-start Unity; start when overlay shown
            Deactivated += (_, _) => Hide();
            PreviewKeyDown += OnInput;
            PreviewMouseMove += OnInput;
            PreviewMouseDown += OnInput;
            System.Diagnostics.Debug.WriteLine("[MainWindow] Constructed");
        }

        private void OnInput(object? sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[MainWindow] Input -> HideUnity");
            HideUnity(true);
        }

        public void ShowUnity()
        {
            System.Diagnostics.Debug.WriteLine("[MainWindow] ShowUnity called");
            if (!_unityRunning)
            {
                StartUnity();
            }
        }

        private void StartUnity()
        {
            if (_unityProcess != null && !_unityProcess.HasExited)
            {
                System.Diagnostics.Debug.WriteLine("[MainWindow] Unity already running PID=" + _unityProcess.Id);
                if (_unityHwnd == IntPtr.Zero)
                {
                    _unityHwnd = _unityProcess.MainWindowHandle;
                    System.Diagnostics.Debug.WriteLine("[MainWindow] Cached handle=" + _unityHwnd);
                }
                return;
            }
            string exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Unity", "Disclosure.exe");
            if (!File.Exists(exePath))
            {
                System.Diagnostics.Debug.WriteLine("[MainWindow] Unity exe missing: " + exePath);
                return;
            }
            _unityProcess = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = exePath,
                    UseShellExecute = true, // allow normal window creation
                    CreateNoWindow = false,
                    WindowStyle = ProcessWindowStyle.Normal
                }
            };
            bool started = _unityProcess.Start();
            _unityRunning = started;
            System.Diagnostics.Debug.WriteLine("[MainWindow] Started Unity process success=" + started + (started ? (" PID=" + _unityProcess.Id) : ""));
            if (started)
            {
                try { _unityProcess.WaitForInputIdle(5000); } catch { }
                _unityHwnd = _unityProcess.MainWindowHandle;
                if (_unityHwnd == IntPtr.Zero)
                {
                    for (int i = 0; i < 20 && _unityHwnd == IntPtr.Zero; i++)
                    {
                        System.Threading.Thread.Sleep(100);
                        _unityHwnd = _unityProcess.MainWindowHandle;
                    }
                }
                System.Diagnostics.Debug.WriteLine("[MainWindow] Initial main window handle=" + _unityHwnd);
            }
        }

        public void ShowUnityWindow()
        {
            System.Diagnostics.Debug.WriteLine("[MainWindow] ShowUnityWindow called");
            if (_unityProcess == null || _unityProcess.HasExited)
            {
                StartUnity();
            }
            if (_unityHwnd == IntPtr.Zero && _unityProcess != null && !_unityProcess.HasExited)
            {
                _unityHwnd = _unityProcess.MainWindowHandle;
                if (_unityHwnd == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine("[MainWindow] MainWindowHandle still zero, attempting EnumWindows refresh");
                    RefreshUnityWindowHandle();
                }
            }
            if (_unityHwnd != IntPtr.Zero)
            {
                bool visible = IsWindowVisible(_unityHwnd);
                WINDOWPLACEMENT wp = new WINDOWPLACEMENT();
                wp.length = (uint)Marshal.SizeOf<WINDOWPLACEMENT>();
                bool gotWp = GetWindowPlacement(_unityHwnd, ref wp);
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Pre-show state visible={visible} gotPlacement={gotWp} showCmd={(gotWp?wp.showCmd:-1)} hwnd={_unityHwnd}");

                bool showOk = true;
                if (!visible)
                {
                    showOk = ShowWindow(_unityHwnd, SW_SHOW); // if hidden (SW_HIDE used previously)
                }
                else if (gotWp && (wp.showCmd == SW_SHOWMINIMIZED))
                {
                    showOk = ShowWindow(_unityHwnd, SW_RESTORE);
                }
                else
                {
                    // Not hidden or minimized; ensure top-most + foreground anyway
                    ShowWindow(_unityHwnd, SW_SHOW); // idempotent
                }

                bool posOk = SetWindowPos(_unityHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
                bool bringOk = BringWindowToTop(_unityHwnd);
                bool fgOk = SetForegroundWindow(_unityHwnd);
                System.Diagnostics.Debug.WriteLine($"[MainWindow] ShowUnityWindow actions showOk={showOk} posOk={posOk} bringOk={bringOk} fgOk={fgOk}");

                if (!fgOk)
                {
                    System.Threading.Tasks.Task.Run(async () =>
                    {
                        await System.Threading.Tasks.Task.Delay(150);
                        bool retry = SetForegroundWindow(_unityHwnd);
                        System.Diagnostics.Debug.WriteLine("[MainWindow] Foreground retry result=" + retry);
                    });
                }
            }
            StartOverlayFocusEnforcement();
        }

        public void HideUnity(bool kill = false)
        {
            if (IsVisible) Hide();
            if (_unityHwnd != IntPtr.Zero)
            {
                ShowWindow(_unityHwnd, SW_HIDE);
                System.Diagnostics.Debug.WriteLine("[MainWindow] HideUnity -> SW_HIDE handle=" + _unityHwnd);
            }
            StopOverlayFocusEnforcement();
            if (kill && _unityProcess != null && !_unityProcess.HasExited)
            {
                try
                {
                    _unityProcess.Kill(true);
                    _unityProcess.WaitForExit(2000);
                    System.Diagnostics.Debug.WriteLine("[MainWindow] Unity process killed per request");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[MainWindow] Kill error: " + ex.Message);
                }
                _unityProcess = null;
                _unityHwnd = IntPtr.Zero;
                _unityRunning = false;
            }
        }

        private void StartOverlayFocusEnforcement()
        {
            if (_overlayFocusTimer != null) return;
            _overlayFocusTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _overlayFocusTimer.Tick += (_, _) => EnforceOverlayForeground();
            _overlayFocusTimer.Start();
        }

        private void StopOverlayFocusEnforcement()
        {
            if (_overlayFocusTimer != null)
            {
                _overlayFocusTimer.Stop();
                _overlayFocusTimer = null;
            }
        }

        private void EnforceOverlayForeground()
        {
            if (_unityHwnd != IntPtr.Zero && IsWindowVisible(_unityHwnd))
            {
                IntPtr fg = GetForegroundWindow();
                if (fg != _unityHwnd)
                {
                    SetForegroundWindow(_unityHwnd);
                    BringWindowToTop(_unityHwnd);
                    System.Diagnostics.Debug.WriteLine("[MainWindow] EnforceOverlayForeground: Restored overlay to foreground");
                }
            }
        }

        #region Win32
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,int X,int Y,int cx,int cy,uint uFlags);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;
        private const int SW_RESTORE = 9;
        private const int SW_SHOWMINIMIZED = 2;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        [StructLayout(LayoutKind.Sequential)]
        private struct WINDOWPLACEMENT
        {
            public uint length;
            public uint flags;
            public uint showCmd;
            public POINT ptMinPosition;
            public POINT ptMaxPosition;
            public RECT rcNormalPosition;
        }
        [StructLayout(LayoutKind.Sequential)] private struct POINT { public int X; public int Y; }
        [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        private void RefreshUnityWindowHandle()
        {
            if (_unityProcess == null || _unityProcess.HasExited) return;
            uint targetPid = (uint)_unityProcess.Id;
            IntPtr found = IntPtr.Zero;
            EnumWindows((h, l) =>
            {
                GetWindowThreadProcessId(h, out uint pid);
                if (pid == targetPid)
                {
                    // pick first visible top-level window
                    if (IsWindowVisible(h))
                    {
                        found = h;
                        return false; // stop enumeration
                    }
                }
                return true;
            }, IntPtr.Zero);
            if (found != IntPtr.Zero)
            {
                _unityHwnd = found;
                System.Diagnostics.Debug.WriteLine("[MainWindow] RefreshUnityWindowHandle found hwnd=" + _unityHwnd);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[MainWindow] RefreshUnityWindowHandle did not find a visible window");
            }
        }
        #endregion
    }
}
