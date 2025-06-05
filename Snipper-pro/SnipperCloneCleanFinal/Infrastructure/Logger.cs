using System;
using System.IO;
using System.Diagnostics;

namespace SnipperCloneCleanFinal.Infrastructure
{
    public static class Logger
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SnipperPro",
            "log.txt"
        );

        private static readonly object LogLock = new object();

        static Logger()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to create log directory: {ex.Message}");
            }
        }

        public static void Log(string message, LogLevel level = LogLevel.Info)
        {
            try
            {
                var logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}{Environment.NewLine}";
                
                lock (LogLock)
                {
                    File.AppendAllText(LogPath, logEntry);
                }
                
                System.Diagnostics.Debug.WriteLine($"SnipperPro: {message}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to write log: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Original message: {message}");
            }
        }

        public static void Error(string message, Exception ex = null)
        {
            var logMessage = ex == null ? message : $"{message} - {ex.GetType().Name}: {ex.Message}";
            Log(logMessage, LogLevel.Error);
            
            if (ex?.StackTrace != null)
            {
                Log($"Stack trace: {ex.StackTrace}", LogLevel.Debug);
            }
        }

        public static void Warning(string message) => Log(message, LogLevel.Warning);
        public static void Info(string message) => Log(message, LogLevel.Info);
        public static void DebugLog(string message) => Log(message, LogLevel.Debug);
    }

    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }
} 