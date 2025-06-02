using System;
using System.IO;

namespace SnipperCloneCleanFinal.Infrastructure
{
    public static class AppConfig
    {
        private static readonly string ConfigPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SnipperPro",
            "config.json");

        public static bool EnableOcrOnOpen { get; set; } = true;
        public static int OcrTimeoutSeconds { get; set; } = 30;

        static AppConfig() => Load();

        public static void Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    var json = File.ReadAllText(ConfigPath);
                    // Simple JSON parsing without external dependencies
                    if (json.Contains("\"EnableOcrOnOpen\":false"))
                        EnableOcrOnOpen = false;
                    if (json.Contains("\"OcrTimeoutSeconds\":"))
                    {
                        var start = json.IndexOf("\"OcrTimeoutSeconds\":") + 20;
                        var end = json.IndexOfAny(new char[] { ',', '}' }, start);
                        if (end > start && int.TryParse(json.Substring(start, end - start).Trim(), out int timeout))
                            OcrTimeoutSeconds = timeout;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading config: {ex.Message}");
            }
        }

        public static void Save()
        {
            try
            {
                var directory = Path.GetDirectoryName(ConfigPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var json = $"{{\"EnableOcrOnOpen\":{EnableOcrOnOpen.ToString().ToLower()},\"OcrTimeoutSeconds\":{OcrTimeoutSeconds}}}";
                File.WriteAllText(ConfigPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving config: {ex.Message}");
            }
        }
    }
} 