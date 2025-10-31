using System;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace StyleWatcherWin
{
    public class TrayApp : Form
    {
        // Win32 Hotkey
        [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        const int WM_HOTKEY = 0x0312;
        const uint MOD_ALT = 0x0001, MOD_CONTROL = 0x0002, MOD_SHIFT = 0x0004, MOD_WIN = 0x0008;

        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly HttpClient _http = new HttpClient();
        readonly AppConfig _cfg;
        int _hotkeyId = 1;
        uint _mod; uint _vk;

        public TrayApp()
        {
            _cfg = AppConfig.Load();
            this.ShowInTaskbar = false;
            this.WindowState = FormWindowState.Minimized;
            this.Visible = false;

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

            // http timeout
            _http.Timeout = TimeSpan.FromSeconds(_cfg.timeout_seconds);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            ParseHotkey(_cfg.hotkey, out _mod, out _vk);
            if (!RegisterHotKey(this.Handle, _hotkeyId, _mod, _vk))
            {
                MessageBox.Show($"热键 {_cfg.hotkey} 注册失败，可能被其他程序占用。\n可在 appsettings.json 修改，例如 Ctrl+Shift+S / Ctrl+Alt+F12。", "StyleWatcher", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            _tray.BalloonTipTitle = "StyleWatcher 已启动";
            _tray.BalloonTipText = $"选中文本后按 {_cfg.hotkey} 查询；右键托盘图标可手动输入或退出。";
            _tray.ShowBalloonTip(3000);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                _ = OnHotkeyAsync();
            }
            base.WndProc(ref m);
        }

        private async Task OnHotkeyAsync()
        {
            try
            {
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
                // headers
                var req = new HttpRequestMessage(
                    _cfg.method?.ToUpperInvariant() == "GET" ? HttpMethod.Get : HttpMethod.Post,
                    _cfg.api_url
                );

                if (_cfg.method?.ToUpperInvariant() == "GET")
                {
                    // simple querystring for GET
                    req = new HttpRequestMessage(HttpMethod.Get, $"{_cfg.api_url}?{_cfg.json_key}={Uri.EscapeDataString(text)}");
                }
                else
                {
                    var json = JsonSerializer.Serialize(new System.Collections.Generic.Dictionary<string,string> { { _cfg.json_key, text } });
                    req.Content = new StringContent(json, Encoding.UTF8, "application/json");
                }

                if (!string.IsNullOrEmpty(_cfg.headers?.Content_Type))
                {
                    // HttpClient will set Content-Type for StringContent; keep default
                }

                var resp = await _http.SendAsync(req);
                var raw = await resp.Content.ReadAsStringAsync();
                try
                {
                    var doc = JsonDocument.Parse(raw);
                    if (doc.RootElement.TryGetProperty("msg", out var msgEl))
                    {
                        var msg = msgEl.ToString();
                        if (msg.Contains("\\n")) msg = msg.Replace("\\n", "\n");
                        return msg;
                    }
                    return raw;
                }
                catch { return raw; }
            }
            catch (Exception ex)
            {
                return $"请求失败：{ex.Message}";
            }
        }

        protected override void Dispose(bool disposing)
        {
            try { UnregisterHotKey(this.Handle, _hotkeyId); } catch {}
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
                else
                {
                    if (Enum.TryParse<Keys>(t, true, out var key)) vk = (uint)key;
                }
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
