using System;
using System.IO;
using System.Text.Json;

namespace LinHPMonitor
{
    public class MonitorConfig
    {
        public string WindowTitle { get; set; } = "Lineage";
        public double HPThreshold { get; set; } = 70;
        public double MPThreshold { get; set; } = 30;
        public string HPHotkey { get; set; } = "F5";
        public string MPHotkey { get; set; } = "F6";
        public int KeyPressDuration { get; set; } = 50;
        public int CheckInterval { get; set; } = 150;
    }

    public static class ConfigManager
    {
        private static readonly string ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");

        public static MonitorConfig Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    return JsonSerializer.Deserialize<MonitorConfig>(json) ?? new MonitorConfig();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading config: {ex.Message}");
            }
            return new MonitorConfig();
        }

        public static void Save(MonitorConfig config)
        {
            try
            {
                string json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving config: {ex.Message}");
            }
        }
    }
}
