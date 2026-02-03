using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace outlook_extension
{
    public static class ShortcutParser
    {
        private const uint ModifierAlt = 0x0001;
        private const uint ModifierControl = 0x0002;
        private const uint ModifierShift = 0x0004;
        private const uint ModifierWin = 0x0008;

        private static readonly HashSet<string> ForbiddenCombos = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Alt+F4"
        };

        public static bool TryParse(string shortcut, out uint modifiers, out uint key)
        {
            modifiers = 0;
            key = 0;

            if (string.IsNullOrWhiteSpace(shortcut))
            {
                return false;
            }

            if (ForbiddenCombos.Contains(shortcut.Trim()))
            {
                return false;
            }

            var parts = shortcut.Split('+');
            Keys parsedKey = Keys.None;

            foreach (var part in parts.Select(item => item.Trim()))
            {
                if (part.Equals("Ctrl", StringComparison.OrdinalIgnoreCase) || part.Equals("Strg", StringComparison.OrdinalIgnoreCase))
                {
                    modifiers |= ModifierControl;
                    continue;
                }

                if (part.Equals("Shift", StringComparison.OrdinalIgnoreCase))
                {
                    modifiers |= ModifierShift;
                    continue;
                }

                if (part.Equals("Alt", StringComparison.OrdinalIgnoreCase))
                {
                    modifiers |= ModifierAlt;
                    continue;
                }

                if (part.Equals("Win", StringComparison.OrdinalIgnoreCase))
                {
                    modifiers |= ModifierWin;
                    continue;
                }

                if (!Enum.TryParse(part, true, out parsedKey))
                {
                    return false;
                }
            }

            if (parsedKey == Keys.None || parsedKey == Keys.ControlKey || parsedKey == Keys.ShiftKey || parsedKey == Keys.Menu)
            {
                return false;
            }

            key = (uint)parsedKey;
            return true;
        }

        public static string Format(Keys keyData)
        {
            var parts = new List<string>();
            if (keyData.HasFlag(Keys.Control))
            {
                parts.Add("Ctrl");
            }

            if (keyData.HasFlag(Keys.Shift))
            {
                parts.Add("Shift");
            }

            if (keyData.HasFlag(Keys.Alt))
            {
                parts.Add("Alt");
            }

            var key = keyData & Keys.KeyCode;
            if (key != Keys.ControlKey && key != Keys.ShiftKey && key != Keys.Menu)
            {
                parts.Add(key.ToString());
            }

            return string.Join("+", parts);
        }
    }
}
