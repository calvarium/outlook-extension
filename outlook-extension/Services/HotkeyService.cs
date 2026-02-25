using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace outlook_extension
{
    public class HotkeyService : IDisposable
    {
        private const int HotkeyId = 0x1000;
        private const int WmHotkey = 0x0312;

        private readonly Outlook.Application _application;
        private readonly SettingsService _settingsService;
        private readonly Action _hotkeyAction;
        private readonly LoggingService _loggingService;
        private readonly HotkeyWindow _hotkeyWindow;
        private bool _isRegistered;
        private readonly Timer _retryTimer;
        private int _retryAttempts;
        private const int MaxRetryAttempts = 15;

        public bool IsRegistered => _isRegistered;
        public HotkeyService(
            Outlook.Application application,
            SettingsService settingsService,
            Action hotkeyAction,
            LoggingService loggingService)
        {
            _application = application;
            _settingsService = settingsService;
            _hotkeyAction = hotkeyAction;
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
            if (!_isRegistered)
            {
                return;
            }

            try
            {
                var handle = _hotkeyWindow.WindowHandle;
                if (handle != IntPtr.Zero)
                {
                    UnregisterHotKey(handle, HotkeyId);
                }
            }
            catch (Exception ex)
            {
                _loggingService.LogError("HotkeyUnregister", ex);
            }
            finally
            {
                _isRegistered = false;
            }
        }

        public void Dispose()
        {
            UnregisterShortcut();
            _retryTimer.Stop();
            _hotkeyWindow.DestroyMessageWindow();
        }

        private void OnHotkeyPressed()
        {
            _hotkeyAction?.Invoke();
        }

        private void RetryRegister()
        {
            if (_isRegistered)
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

            if (!ShortcutParser.TryParse(_settingsService.Current.Shortcut, out var modifiers, out var key))
            {
                return;
            }

            if (!RegisterHotKey(handle, HotkeyId, modifiers, key))
            {
                _loggingService.LogInfo("Hotkey Registrierung fehlgeschlagen.");
                if (!_retryTimer.Enabled)
                {
                    _retryTimer.Start();
                }
                return;
            }

            _isRegistered = true;
            _retryTimer.Stop();
        }

        private class HotkeyWindow : NativeWindow
        {
            private readonly Action _callback;

            public HotkeyWindow(Action callback)
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
                    _callback?.Invoke();
                }

                base.WndProc(ref m);
            }
        }

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    }
}
