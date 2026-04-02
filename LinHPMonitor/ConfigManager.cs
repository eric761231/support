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
        // Saved capture regions (top-left). Width/Height default to the bar size.
        public int HPRegionX { get; set; } = 0;
        public int HPRegionY { get; set; } = 0;
        public int MPRegionX { get; set; } = 0;
        public int MPRegionY { get; set; } = 0;
        public int RegionWidth { get; set; } = 260;
        public int RegionHeight { get; set; } = 40;
        // 是否在啟動時跳過自動校準，改用目前預覽區域（由 UI 控制）
        public bool SkipCalibrationForPreview { get; set; } = true;
        // 是否啟用 PaddleOCR 做為 Tesseract 的備援
        public bool UsePaddleOcr { get; set; } = true;
        // 使用者手動指定的最大值，可讓右側數字固定為此值（0 表示自動偵測）
        public int UserMaxHp { get; set; } = 0;
        public int UserMaxMp { get; set; } = 0;
        // 每次觸發時連按熱鍵的次數（預設 1；調高至 2 可應對需雙擊才生效的遊戲技能）
        public int PressCount { get; set; } = 1;
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
