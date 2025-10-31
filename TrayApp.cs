using System;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Windows.Automation;
using System.Collections.Generic;

namespace StyleWatcherWin
{
    public static class Formatter
    {
        public static string Prettify(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return raw;
            var s = raw.Replace("\\n", "\n").Replace("\r\n", "\n");
            var lines = s.Split('\n');
            for (int i = 0; i < lines.Length; i++) lines[i] = lines[i].Trim();
            s = string.Join("\n", lines);
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\n{3,}", "\n\n");
            return s.Trim();
        }
    }

    public class TrayApp : Form
    {
        [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll")] static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")] static extern IntPtr GetFocus();
        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern IntPtr SendMessage(IntPtr hWnd, int msg, ref int wParam, ref int lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, StringBuilder lParam);

        const int WM_GETTEXTLENGTH = 0x000E;
        const int WM_GETTEXT = 0x000D;
        const int EM_GETSEL = 0x00B0;
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
        uint _mod;
        uint _vk;

        public TrayApp()
        {
            _cfg = AppConfig.Load();
            ShowInTaskbar = false;
            WindowState = FormWindowState.Minimized;
            Visible = false;

            _tray.Text = "StyleWatcher";
            _tray.Icon = SystemIcons.Information;
            _tray.Visible = true;

            var itemQuery = new ToolStripMenuItem("手动输入查询", null, (s, e) => ShowResultWindow("", "请输入要查询的文本后回车"));
            var itemConfig = new ToolStripMenuItem("打开配置文件", null, (s, e) =>
            {
                try { System.Diagnostics.Process.Start("notepad.exe", AppConfig.ConfigPath); } catch { }
            });
            var itemExit = new ToolStripMenuItem("退出", null, (s, e) => { Application.Exit(); });

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
                MessageBox.Show($"热键 {_cfg.hotkey} 注册失败，可能被占用。", "StyleWatcher",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);

            _tray.BalloonTipTitle = "StyleWatcher 已启动";
            _tray.BalloonTipText = $"选中文本后按 {_cfg.hotkey} 查询；右键托盘图标可手动输入或退出。";
            _tray.ShowBalloonTip(2500);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
                _ = OnHotkeyAsync();
            base.WndProc(ref m);
        }

        // ---- 获取选中文本：优先 UIA，其次 Win32，最后剪贴板兜底 ----
        private string TryGetSelectedTextUsingUIA()
        {
            try
            {
                var element = AutomationElement.FocusedElement;
                if (element == null) return null;

                if (element.TryGetCurrentPattern(TextPattern.Pattern, out object tpObj) && tpObj is TextPattern tp)
                {
                    var ranges = tp.GetSelection();
                    if (ranges != null && ranges.Length > 0)
                    {
                        var txt = ranges[0].GetText(-1) ?? "";
                        txt = txt.Trim();
                        if (!string.IsNullOrEmpty(txt)) return txt;
                    }
                }

                if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object vpObj) && vpObj is ValuePattern vp2)
                {
                    var val = (vp2.Current.Value ?? "").Trim();
                    if (!string.IsNullOrEmpty(val)) return val;
                }
            }
            catch { }
            return null;
        }

        private string TryGetSelectedTextUsingWin32()
        {
            try
            {
                var fg = GetForegroundWindow();
                if (fg == IntPtr.Zero) return null;

                uint fgThread = GetWindowThreadProcessId(fg, out _);
                uint curThread = GetCurrentThreadId();
                bool attached = false;

                try
                {
                    attached = AttachThreadInput(curThread, fgThread, true);
                    var hFocus = GetFocus();
                    if (hFocus == IntPtr.Zero) return null;

                    int start = 0, end = 0;
                    SendMessage(hFocus, EM_GETSEL, ref start, ref end);

                    int len = (int)SendMessage(hFocus, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);
                    if (len <= 0) return null;

                    var sb = new StringBuilder(len + 1);
                    SendMessage(hFocus, WM_GETTEXT, sb.Capacity, sb);
                    var full = sb.ToString();

                    if (start < 0 || end < 0 || start > full.Length) return null;
                    if (end > full.Length) end = full.Length;
                    if (end > start) return full.Substring(start, end - start).Trim();
                }
                finally
                {
                    if (attached) AttachThreadInput(curThread, fgThread, false);
                }
            }
            catch { }
            return null;
        }

        private async Task<string> GetSelectionByClipboardRoundTripAsync()
        {
            IDataObject backup = null;
            try { backup = Clipboard.GetDataObject(); } catch { }

            SendKeys.SendWait("^c");
            await Task.Delay(120);

            string txt = "";
            try { txt = Clipboard.GetText()?.Trim() ?? ""; } catch { }

            if (backup != null)
            {
                try { Clipboard.SetDataObject(backup, true); } catch { }
            }
            return txt;
        }

        private async Task OnHotkeyAsync()
        {
            try
            {
                // 防止 Alt 状态卡住
                ReleaseAlt();

                string txt = TryGetSelectedTextUsingUIA();
                if (string.IsNullOrEmpty(txt)) txt = TryGetSelectedTextUsingWin32();
                if (string.IsNullOrEmpty(txt)) txt = await GetSelectionByClipboardRoundTripAsync();

                if (string.IsNullOrEmpty(txt))
                {
                    _tray.ShowBalloonTip(2000, "StyleWatcher", "未检测到选中文本，请先选中一段文字再按热键。", ToolTipIcon.Info);
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
                {
                    req = new HttpRequestMessage(HttpMethod.Get,
                        $"{_cfg.api_url}?{_cfg.json_key}={Uri.EscapeDataString(text)}");
                }
                else
                {
                    req = new HttpRequestMessage(HttpMethod.Post, _cfg.api_url);
                    var payload = new Dictionary<string, string> { { _cfg.json_key, text } };
                    var json = JsonSerializer.Serialize(payload);
                    req.Content = new StringContent(json, Encoding.UTF8, "application/json");
                }

                var resp = await _http.SendAsync(req);
                var raw = await resp.Content.ReadAsStringAsync();

                // 兼容 { "msg": "..." } 以及纯文本
                try
                {
                    using var doc = JsonDocument.Parse(raw);
                    if (doc.RootElement.TryGetProperty("msg", out var msgEl))
                        return Formatter.Prettify(msgEl.ToString());
                    return Formatter.Prettify(raw);
                }
                catch
                {
                    return Formatter.Prettify(raw);
                }
            }
            catch (Exception ex)
            {
                return $"请求失败：{ex.Message}";
            }
        }

        protected override void Dispose(bool disposing)
        {
            try { UnregisterHotKey(Handle, _hotkeyId, _mod, _vk); } catch { }
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
            mod = 0;
            vk  = 0;
            if (string.IsNullOrWhiteSpace(s))
            {
                mod = MOD_ALT; vk = (uint)Keys.S; return;
            }
            var parts = s.Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var t = p.Trim().ToUpperInvariant();
                if (t == "CTRL" || t == "CONTROL") mod |= MOD_CONTROL;
                else if (t == "SHIFT") mod |= MOD_SHIFT;
                else if (t == "ALT") mod |= MOD_ALT;
                else if (t == "WIN" || t == "WINDOWS") mod |= MOD_WIN;
                else if (Enum.TryParse(t, true, out Keys key)) vk = (uint)key;
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
