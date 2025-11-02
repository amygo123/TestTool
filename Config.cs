using System;
using System.IO;
using System.Text.Json;

namespace StyleWatcherWin
{
    public class AppConfig
    {
        public string api_url { get; set; } = "http://47.111.189.27:8889/qrcode/saleVolumeParser";
        public string method { get; set; } = "POST";
        public string json_key { get; set; } = "code";
        public int timeout_seconds { get; set; } = 6;
        public string hotkey { get; set; } = "Alt+S";
        public WindowCfg window { get; set; } = new WindowCfg();
        public Headers headers { get; set; } = new Headers();

        // === 库存相关配置（新增） ===
        public string inventory_api_url { get; set; } = "http://192.168.40.97:8000/inventory";
        public int inventory_timeout_seconds { get; set; } = 4;
        public int inventory_cache_ttl_seconds { get; set; } = 300;
        public int inventory_low_threshold { get; set; } = 10;

        public class WindowCfg { public int width { get; set; } = 900; public int height { get; set; } = 600; public int fontSize { get; set; } = 12; public bool alwaysOnTop { get; set; } = true; }
        public class Headers { public string Content_Type { get; set; } = "application/json"; }

        public static string ConfigDirectory => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "StyleWatcher");
        public static string ConfigPath => Path.Combine(ConfigDirectory, "appsettings.json");

        public static AppConfig Load()
        {
            try
            {
                if (!Directory.Exists(ConfigDirectory)) Directory.CreateDirectory(ConfigDirectory);
                if (File.Exists(ConfigPath))
                {
                    var json = File.ReadAllText(ConfigPath);
                    return JsonSerializer.Deserialize<AppConfig>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new AppConfig();
                }
            }
            catch { }
            return new AppConfig();
        }

        public void Save()
        {
            try
            {
                if (!Directory.Exists(ConfigDirectory)) Directory.CreateDirectory(ConfigDirectory);
                var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, json);
            }
            catch { }
        }
    }
}
