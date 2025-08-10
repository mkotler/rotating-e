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
            HideUnity();
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
                    UseShellExecute = false,
                    CreateNoWindow = true,
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
            }
            if (_unityHwnd != IntPtr.Zero)
            {
                ShowWindow(_unityHwnd, SW_SHOW);
                SetWindowPos(_unityHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                System.Diagnostics.Debug.WriteLine("[MainWindow] ShowUnityWindow -> SW_SHOW");
            }
        }

        public void HideUnity(bool kill = false)
        {
            if (IsVisible) Hide();
            if (_unityHwnd != IntPtr.Zero)
            {
                ShowWindow(_unityHwnd, SW_HIDE);
                System.Diagnostics.Debug.WriteLine("[MainWindow] HideUnity -> SW_HIDE handle=" + _unityHwnd);
            }
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

        #region Win32
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,int X,int Y,int cx,int cy,uint uFlags);
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOACTIVATE = 0x0010;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        #endregion
    }
}
