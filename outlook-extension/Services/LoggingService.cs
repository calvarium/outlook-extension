using System;
using System.IO;

namespace outlook_extension
{
    public class LoggingService
    {
        private readonly string _logPath;

        public LoggingService()
        {
            var folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "QuickMoveOutlook");
            Directory.CreateDirectory(folder);
            _logPath = Path.Combine(folder, "quickmove.log");
        }

        public void LogInfo(string message)
        {
            Write("INFO", message);
        }

        public void LogError(string context, Exception ex)
        {
            Write("ERROR", $"{context}: {ex}");
        }

        private void Write(string level, string message)
        {
            try
            {
                File.AppendAllText(_logPath, $"{DateTime.Now:O} [{level}] {message}{Environment.NewLine}");
            }
            catch
            {
                // Logging ist optional und darf den Ablauf nicht stören.
            }
        }
    }
}
