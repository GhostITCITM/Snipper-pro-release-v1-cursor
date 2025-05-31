using System;
using System.IO;
using Newtonsoft.Json;

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
                    dynamic cfg = JsonConvert.DeserializeObject(json);
                    EnableOcrOnOpen = cfg.EnableOcrOnOpen ?? true;
                    OcrTimeoutSeconds = cfg.OcrTimeoutSeconds ?? 30;
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

                File.WriteAllText(ConfigPath, JsonConvert.SerializeObject(new { EnableOcrOnOpen, OcrTimeoutSeconds }));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving config: {ex.Message}");
            }
        }
    }
} 