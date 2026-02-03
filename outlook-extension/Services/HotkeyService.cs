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
        }

        public void RegisterShortcut()
        {
            UnregisterShortcut();

            var handle = GetOutlookWindowHandle();
            if (handle == IntPtr.Zero)
            {
                return;
            }

            if (_hotkeyWindow.Handle != IntPtr.Zero)
            {
                _hotkeyWindow.ReleaseHandle();
            }

            _hotkeyWindow.AssignHandle(handle);

            if (!ShortcutParser.TryParse(_settingsService.Current.Shortcut, out var modifiers, out var key))
            {
                return;
            }

            if (!RegisterHotKey(handle, HotkeyId, modifiers, key))
            {
                _loggingService.LogInfo("Hotkey Registrierung fehlgeschlagen.");
                return;
            }

            _isRegistered = true;
        }

        public void UnregisterShortcut()
        {
            if (!_isRegistered)
            {
                return;
            }

            try
            {
                var handle = GetOutlookWindowHandle();
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
            _hotkeyWindow.ReleaseHandle();
        }

        private void OnHotkeyPressed()
        {
            _hotkeyAction?.Invoke();
        }

        private static IntPtr GetOutlookWindowHandle()
        {
            var handle = Process.GetCurrentProcess().MainWindowHandle;
            if (handle != IntPtr.Zero)
            {
                return handle;
            }

            return GetForegroundWindow();
        }

        private class HotkeyWindow : NativeWindow
        {
            private readonly Action _callback;

            public HotkeyWindow(Action callback)
            {
                _callback = callback;
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

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
    }
}
