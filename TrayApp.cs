using System;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace StyleWatcherWin
{
    public static class Formatter
    {
        /// <summary>
        /// 将接口 msg 文本美化为易读格式：
        /// 1) 顶部：标题（选中文本/商品线） + 昨日销量
        /// 2) 第二行：近7天汇总
        /// 3) 后续：按“日期 款名 尺码 颜色：X件”逐行展示
        /// </summary>
        public static string Prettify(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return raw;

            var s = raw.Replace("\\n", "\n").Trim();
            s = Regex.Replace(s, @"[ \t]+", " ");

            var lines = new StringBuilder();
            string titleLine = null;
            string yestLine = null;
            string weekSum = null;

            // 标题 + 昨日销量
            var titleMatch = Regex.Match(s, @"^(?<title>.*?)(?:[:：]\s*)?(?<yest>昨日[^\n]*)", RegexOptions.Singleline);
            if (titleMatch.Success)
            {
                titleLine = titleMatch.Groups["title"].Value.Trim();
                yestLine = titleMatch.Groups["yest"].Value.Trim();
            }

            // 近7天汇总
            var weekMatch = Regex.Match(s, @"近\s*7\s*天\s*销量\s*汇\s*总[:：]\s*(?<sum>\d+)");
            if (weekMatch.Success)
            {
                weekSum = $"近7天销量汇总：{weekMatch.Groups["sum"].Value}";
            }

            if (!string.IsNullOrEmpty(titleLine)) lines.AppendLine(titleLine);
            if (!string.IsNullOrEmpty(yestLine)) lines.AppendLine(yestLine);
            if (!string.IsNullOrEmpty(weekSum))  lines.AppendLine(weekSum);

            // 按日期拆分：2025-10-24 开头的行
            var segs = Regex.Split(s, @"(?=(20\d{2}-\d{2}-\d{2})\s+)");
            foreach (var seg in segs)
            {
                var m = Regex.Match(seg, @"^(20\d{2}-\d{2}-\d{2})\s+(?<rest>.+)$");
                if (!m.Success) continue;

                var date = m.Groups[1].Value;
                var rest = m.Groups["rest"].Value.Trim();

                // : X件
                var qtyMatch = Regex.Match(rest, @"[:：]\s*(\d+)\s*件");
                var qtyText = qtyMatch.Success ? $"{qtyMatch.Groups[1].Value}件" : "";
                if (qtyMatch.Success) rest = rest[..qtyMatch.Index].Trim();

                // 尺码
                var sizeMatch = Regex.Match(rest, @"\b(S|M|L|XL|2XL|3XL|4XL)\b");
                string left = rest, size = "", color = "";
                if (sizeMatch.Success)
                {
                    left = rest[..sizeMatch.Index].Trim();
                    size = sizeMatch.Value;
                    var after = rest[(sizeMatch.Index + sizeMatch.Length)..].Trim();
                    var colorMatch = Regex.Match(after, @"^([^\s：:]+)");
                    color = colorMatch.Success ? colorMatch.Groups[1].Value : after;
                }

                var prettyLine = new StringBuilder();
                prettyLine.Append(date).Append("  ");
                if (!string.IsNullOrEmpty(left))  prettyLine.Append(left).Append("  ");
                if (!string.IsNullOrEmpty(size))  prettyLine.Append(size).Append("  ");
                if (!string.IsNullOrEmpty(color)) prettyLine.Append(color).Append("  ");
                if (!string.IsNullOrEmpty(qtyText)) prettyLine.Append("：").Append(qtyText);

                var final = prettyLine.ToString().Trim();
                if (!string.IsNullOrEmpty(final))
                    lines.AppendLine(final);
            }

            var output = lines.ToString().TrimEnd();
            if (string.IsNullOrWhiteSpace(output))
                output = Regex.Replace(s, @"\s*\n\s*", "\n").Trim();

            return output;
        }
    }

    public class TrayApp : Form
    {
        [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        const int WM_HOTKEY = 0x0312;
        const uint MOD_ALT = 0x0001, MOD_CONTROL = 0x0002, MOD_SHIFT = 0x0004, MOD_WIN = 0x0008;
        const int KEYEVENTF_KEYUP = 0x0002;
        const byte VK_MENU = 0x12; // Alt

        static void ReleaseAlt()
        {
            if ((GetAsyncKeyState(VK_MENU) & 0x8000) != 0)
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        }

        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly HttpClient _http = new HttpClient();
        readonly AppConfig _cfg;
        int _hotkeyId = 1;
        uint _mod; uint _vk;

        public TrayApp()
        {
            _cfg = AppConfig.Load();
            ShowInTaskbar = false;
            WindowState = FormWindowState.Minimized;
            Visible = false;

            _tray.Text = "StyleWatcher";
            _tray.Icon = SystemIcons.Information;
            _tray.Visible = true;

            var itemQuery = new ToolStripMenuItem("手动输入查询", null, (s,e)=> ShowResultWindow("", "请输入要查询的文本后回车"));
            var itemConfig = new ToolStripMenuItem("打开配置文件", null, (s,e)=> { try { System.Diagnostics.Process.Start("notepad.exe", AppConfig.ConfigPath); } catch {} });
            var itemExit = new ToolStripMenuItem("退出", null, (s,e)=> { Application.Exit(); });
            _menu.Items.Add(itemQuery);
            _menu.Items.Add(itemConfig);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(itemExit);
            _tray.ContextMenuStrip = _menu;

            _http.Timeout = TimeSpan.FromSeconds(_cfg.timeout_seconds);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            ParseHotkey(_cfg.hotkey, out _mod, out _vk);
            if (!RegisterHotKey(Handle, _hotkeyId, _mod, _vk))
                MessageBox.Show($"热键 {_cfg.hotkey} 注册失败，可能被占用。可在 appsettings.json 修改为 Ctrl+Shift+S / Ctrl+Alt+F12。", "StyleWatcher", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            _tray.BalloonTipTitle = "StyleWatcher 已启动";
            _tray.BalloonTipText = $"选中文本后按 {_cfg.hotkey} 查询；右键托盘图标可手动输入或退出。";
            _tray.ShowBalloonTip(2500);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY) _ = OnHotkeyAsync();
            base.WndProc(ref m);
        }

        private async Task OnHotkeyAsync()
        {
            try
            {
                ReleaseAlt();
                SendKeys.SendWait("^c");
                await Task.Delay(120);
                string txt = "";
                try { txt = Clipboard.GetText()?.Trim() ?? ""; } catch {}
                if (string.IsNullOrEmpty(txt))
                {
                    _tray.ShowBalloonTip(2000, "StyleWatcher", "未检测到剪贴板文本，请先选中文本再按热键。", ToolTipIcon.Info);
                    return;
                }
                string result = await QueryAsync(txt);
                ShowResultWindow(txt, result);
                ReleaseAlt();
            }
            catch (Exception ex)
            {
                _tray.ShowBalloonTip(2000, "错误", ex.Message, ToolTipIcon.Error);
            }
        }

        private async Task<string> QueryAsync(string text)
        {
            try
            {
                var method = (_cfg.method ?? "POST").ToUpperInvariant();
                HttpRequestMessage req;
                if (method == "GET")
                    req = new HttpRequestMessage(HttpMethod.Get, $"{_cfg.api_url}?{_cfg.json_key}={Uri.EscapeDataString(text)}");
                else
                {
                    req = new HttpRequestMessage(HttpMethod.Post, _cfg.api_url);
                    var json = JsonSerializer.Serialize(new System.Collections.Generic.Dictionary<string,string> { { _cfg.json_key, text } });
                    req.Content = new StringContent(json, Encoding.UTF8, "application/json");
                }
                var resp = await _http.SendAsync(req);
                var raw = await resp.Content.ReadAsStringAsync();
                try
                {
                    var doc = JsonDocument.Parse(raw);
                    if (doc.RootElement.TryGetProperty("msg", out var msgEl))
                        return Formatter.Prettify(msgEl.ToString());
                    return Formatter.Prettify(raw);
                }
                catch { return Formatter.Prettify(raw); }
            }
            catch (Exception ex) { return $"请求失败：{ex.Message}"; }
        }

        protected override void Dispose(bool disposing)
        {
            try { UnregisterHotKey(Handle, _hotkeyId, _mod, _vk); } catch {}
            if (disposing)
            {
                _tray.Visible = false;
                _tray.Dispose();
                _menu?.Dispose();
                _http?.Dispose();
            }
            base.Dispose(disposing);
        }

        private void ParseHotkey(string s, out uint mod, out uint vk)
        {
            mod = 0; vk = 0;
            if (string.IsNullOrWhiteSpace(s)) { mod = MOD_ALT; vk = (uint)Keys.S; return; }
            var parts = s.Split(new[]{'+'}, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var t = p.Trim().ToUpperInvariant();
                if (t == "CTRL" || t == "CONTROL") mod |= MOD_CONTROL;
                else if (t == "SHIFT") mod |= MOD_SHIFT;
                else if (t == "ALT") mod |= MOD_ALT;
                else if (t == "WIN" || t == "WINDOWS") mod |= MOD_WIN;
                else if (Enum.TryParse<Keys>(t, true, out var key)) vk = (uint)key;
            }
            if (vk == 0) { mod = MOD_ALT; vk = (uint)Keys.S; }
        }

        private void ShowResultWindow(string input, string result)
        {
            using (var f = new ResultForm(_cfg, input, result))
            {
                f.StartPosition = FormStartPosition.Manual;
                var p = Cursor.Position;
                f.Location = new Point(p.X + 12, p.Y + 12);
                f.TopMost = _cfg.window.alwaysOnTop;
                f.ShowDialog();
            }
        }
    }
}
