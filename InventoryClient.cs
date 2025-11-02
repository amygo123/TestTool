using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace StyleWatcherWin
{
    public record InventoryRow(string Warehouse, string Color, string Size, int QtyIn, int QtyOut);

    public class InventoryClient
    {
        private readonly string _apiUrl;
        private readonly int _timeoutSeconds;

        public InventoryClient(AppConfig cfg)
        {
            _apiUrl = string.IsNullOrWhiteSpace(cfg.inventory_api_url) ? "http://192.168.40.97:8000/inventory" : cfg.inventory_api_url;
            _timeoutSeconds = cfg.inventory_timeout_seconds > 0 ? cfg.inventory_timeout_seconds : 4;
        }

        public async Task<List<InventoryRow>> FetchAsync(string styleName, CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(styleName)) return new List<InventoryRow>();
            var url = $"{_apiUrl}?style_name={Uri.EscapeDataString(styleName)}";
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(_timeoutSeconds) };
            using var resp = await http.GetAsync(url, ct);
            resp.EnsureSuccessStatusCode();
            var raw = await resp.Content.ReadAsStringAsync(ct);
            // 期望为 JSON 数组，元素是形如 "品名，颜色，尺码，仓库，x，y" 的字符串
            var list = JsonSerializer.Deserialize<List<string>>(raw, new JsonSerializerOptions { AllowTrailingCommas = true }) ?? new();
            var rows = new List<InventoryRow>();
            foreach (var line in list)
            {
                var parts = line.Split('，');
                if (parts.Length < 6) continue;
                var color = parts[1].Trim();
                var size = parts[2].Trim();
                var wh = parts[3].Trim();
                if (!int.TryParse(parts[4].Trim(), out var inQty)) continue;
                if (!int.TryParse(parts[5].Trim(), out var outQty)) continue;
                rows.Add(new InventoryRow(wh, color, size, inQty, outQty));
            }
            return rows;
        }
    }
}
