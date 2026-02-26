using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace outlook_extension
{
    public class HotkeyService : IDisposable
    {
        private const int HotkeyIdMain = 0x1000;
        private const int HotkeyIdSettings = 0x1001;
        private const int WmHotkey = 0x0312;

        private readonly Outlook.Application _application;
        private readonly SettingsService _settingsService;
        private readonly Action _hotkeyActionMain;
        private readonly Action _hotkeyActionSettings;
        private readonly LoggingService _loggingService;
        private readonly HotkeyWindow _hotkeyWindow;
        private bool _isRegisteredMain;
        private bool _isRegisteredSettings;
        private readonly Timer _retryTimer;
        private int _retryAttempts;
        private const int MaxRetryAttempts = 15;

        public bool IsRegistered => _isRegisteredMain || _isRegisteredSettings;

        public HotkeyService(
            Outlook.Application application,
            SettingsService settingsService,
            Action hotkeyActionMain,
            LoggingService loggingService,
            Action hotkeyActionSettings = null)
        {
            _application = application;
            _settingsService = settingsService;
            _hotkeyActionMain = hotkeyActionMain;
            _hotkeyActionSettings = hotkeyActionSettings;
            _loggingService = loggingService;
            _hotkeyWindow = new HotkeyWindow(OnHotkeyPressed);
            // Create a dedicated (invisible) message window immediately.
            _hotkeyWindow.CreateMessageWindow();

            _retryTimer = new Timer { Interval = 1000 };
            _retryTimer.Tick += (sender, args) => RetryRegister();
        }

        public void RegisterShortcut()
        {
            _retryAttempts = 0;
            AttemptRegister();
        }

        public void UnregisterShortcut()
        {
            _retryTimer.Stop();
            try
            {
                var handle = _hotkeyWindow.WindowHandle;
                if (handle != IntPtr.Zero)
                {
                    if (_isRegisteredMain)
                    {
                        UnregisterHotKey(handle, HotkeyIdMain);
                    }
                    if (_isRegisteredSettings)
                    {
                        UnregisterHotKey(handle, HotkeyIdSettings);
                    }
                }
            }
            catch (Exception ex)
            {
                _loggingService.LogError("HotkeyUnregister", ex);
            }
            finally
            {
                _isRegisteredMain = false;
                _isRegisteredSettings = false;
            }
        }

        public void Dispose()
        {
            UnregisterShortcut();
            _retryTimer.Stop();
            _hotkeyWindow.DestroyMessageWindow();
        }

        private void OnHotkeyPressed(int id)
        {
            try
            {
                if (id == HotkeyIdMain)
                {
                    _hotkeyActionMain?.Invoke();
                }
                else if (id == HotkeyIdSettings)
                {
                    _hotkeyActionSettings?.Invoke();
                }
            }
            catch (Exception ex)
            {
                try { _loggingService.LogError("HotkeyCallback", ex); } catch { }
            }
        }

        private void RetryRegister()
        {
            if (IsRegistered)
            {
                _retryTimer.Stop();
                return;
            }

            _retryAttempts++;
            if (_retryAttempts > MaxRetryAttempts)
            {
                _retryTimer.Stop();
                return;
            }

            AttemptRegister();
        }

        private void AttemptRegister()
        {
            UnregisterShortcut();

            var handle = _hotkeyWindow.WindowHandle;
            if (handle == IntPtr.Zero)
            {
                if (!_retryTimer.Enabled)
                {
                    _retryTimer.Start();
                }
                return;
            }

            // Register main shortcut
            _isRegisteredMain = false;
            _isRegisteredSettings = false;

            try
            {
                if (ShortcutParser.TryParse(_settingsService.Current.Shortcut, out var modifiersMain, out var keyMain))
                {
                    if (!RegisterHotKey(handle, HotkeyIdMain, modifiersMain, keyMain))
                    {
                        var err = Marshal.GetLastWin32Error();
                        _loggingService.LogInfo($"Hotkey Registrierung fehlgeschlagen für Haupt-Shortcut ({_settingsService.Current.Shortcut}). Win32Error={err}");
                    }
                    else
                    {
                        _isRegisteredMain = true;
                        _loggingService.LogInfo($"Hotkey registriert: Main={_settingsService.Current.Shortcut}");
                    }
                }
            }
            catch (Exception ex)
            {
                _loggingService.LogError("HotkeyRegisterMain", ex);
            }

            // Register settings shortcut if configured
            try
            {
                var settingsShortcut = _settingsService.Current.SettingsShortcut;
                if (!string.IsNullOrWhiteSpace(settingsShortcut) && ShortcutParser.TryParse(settingsShortcut, out var modifiersSettings, out var keySettings))
                {
                    if (!RegisterHotKey(handle, HotkeyIdSettings, modifiersSettings, keySettings))
                    {
                        var err = Marshal.GetLastWin32Error();
                        _loggingService.LogInfo($"Hotkey Registrierung fehlgeschlagen für Settings-Shortcut ({settingsShortcut}). Win32Error={err}");
                    }
                    else
                    {
                        _isRegisteredSettings = true;
                        _loggingService.LogInfo($"Hotkey registriert: Settings={settingsShortcut}");
                    }
                }
            }
            catch (Exception ex)
            {
                _loggingService.LogError("HotkeyRegisterSettings", ex);
            }

            if ((!_isRegisteredMain && !_isRegisteredSettings) && !_retryTimer.Enabled)
            {
                _retryTimer.Start();
            }
            else
            {
                _retryTimer.Stop();
            }
        }

        private class HotkeyWindow : NativeWindow
        {
            private readonly Action<int> _callback;

            public HotkeyWindow(Action<int> callback)
            {
                _callback = callback;
            }

            public IntPtr WindowHandle => base.Handle;

            public void CreateMessageWindow()
            {
                try
                {
                    var cp = new CreateParams
                    {
                        Caption = "QuickMoveHotkeyWindow"
                    };
                    CreateHandle(cp);
                }
                catch
                {
                    // swallow - retries handled by HotkeyService
                }
            }

            public void DestroyMessageWindow()
            {
                try
                {
                    if (base.Handle != IntPtr.Zero)
                    {
                        DestroyHandle();
                    }
                }
                catch
                {
                    // ignore
                }
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WmHotkey)
                {
                    try
                    {
                        var id = m.WParam.ToInt32();
                        _callback?.Invoke(id);
                    }
                    catch { }
                }

                base.WndProc(ref m);
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    }
}
