using System;

namespace Document_converter
{
    public interface ILogger
    {
        void Info(string message);
        void Error(string message, Exception ex);
    }

    public sealed class Logger : ILogger
    {
        private static readonly Lazy<Logger> _instance = new Lazy<Logger>(() => new Logger());
        private readonly string _logPath;
        private readonly object _lock = new object();

        public static Logger Instance => _instance.Value;

        private Logger()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var logDir = System.IO.Path.Combine(appData, "DocumentConverter", "Logs");
            System.IO.Directory.CreateDirectory(logDir);
            _logPath = System.IO.Path.Combine(logDir, $"app_{DateTime.Now:yyyyMMdd}.log");
        }

        public void Info(string message)
        {
            WriteLog("INFO", message);
        }

        public void Error(string message, Exception ex)
        {
            WriteLog("ERROR", $"{message}\n{ex}");
        }

        private void WriteLog(string level, string message)
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var logLine = $"[{timestamp}] [{level}] {message}{Environment.NewLine}";

            lock (_lock)
            {
                try
                {
                    System.IO.File.AppendAllText(_logPath, logLine);
                }
                catch
                {
                    // Silently fail if logging fails
                }
            }
        }
    }
}