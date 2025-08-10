using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;

namespace RotatingEOverlay
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private OutlookWatcher? _watcher;
        private MainWindow? _overlay;
        private bool _armed;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            System.Diagnostics.Debug.WriteLine("[App] Startup");
            _overlay = new MainWindow();
            _overlay.Hide();

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
        }

        private void ShowOverlay()
        {
            System.Diagnostics.Debug.WriteLine("[App] ShowOverlay called");
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
            base.OnExit(e);
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && _armed)
            {
                _armed = false;
                System.Diagnostics.Debug.WriteLine("[App] Input detected -> hiding overlay & Unity");
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        HideOverlay();
                        _overlay?.HideUnity(); // force ensure unity hides even if overlay check failed
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("[App] Exception hiding on input: " + ex.Message);
                    }
                });
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
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
        #endregion
    }
}

