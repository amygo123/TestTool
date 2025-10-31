using System;
using System.Drawing;
using System.Windows.Forms;

namespace StyleWatcherWin
{
    public class ResultForm : Form
    {
        readonly TextBox _boxInput = new TextBox();
        readonly TextBox _boxResult = new TextBox();
        readonly Button _btnQuery = new Button();
        readonly Button _btnCopy = new Button();
        readonly Button _btnClose = new Button();
        readonly AppConfig _cfg;

        public ResultForm(AppConfig cfg, string input, string result)
        {
            _cfg = cfg;
            Text = "StyleWatcher";
            Width = cfg.window.width;
            Height = cfg.window.height;
            KeyPreview = true;

            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 32 };
            var lbl = new Label { Text = "选中文本：", AutoSize = true, Margin = new Padding(6,8,0,0) };
            _boxInput.Width = cfg.window.width - 260;
            _boxInput.Text = input;
            _btnQuery.Text = "查询(Enter)";
            _btnCopy.Text = "复制结果";
            _btnClose.Text = "关闭(Esc)";
            top.Controls.Add(lbl); top.Controls.Add(_boxInput);
            top.Controls.Add(_btnQuery); top.Controls.Add(_btnCopy); top.Controls.Add(_btnClose);
            Controls.Add(top);

            _boxResult.Multiline = true;
            _boxResult.ReadOnly = true;
            _boxResult.ScrollBars = ScrollBars.Vertical;
            _boxResult.Dock = DockStyle.Fill;
            _boxResult.Font = new Font("Consolas", cfg.window.fontSize);
            _boxResult.Text = result;
            Controls.Add(_boxResult);

            var hint = new Label { Dock = DockStyle.Bottom, Height = 22, ForeColor = Color.Gray, Text = $"提示：全局热键 {cfg.hotkey}；Esc 关闭；Ctrl+C 复制结果；回车可再次查询。" };
            Controls.Add(hint);

            _btnQuery.Click += async (s,e)=> {
                _boxResult.Text = "查询中...";
                var textNow = _boxInput.Text.Trim();
                _boxResult.Text = await ApiHelper.QueryAsync(cfg, textNow);
            };
            _btnCopy.Click += (s,e)=> { try { Clipboard.SetText(_boxResult.Text); } catch {} };
            _btnClose.Click += (s,e)=> Close();

            KeyDown += (s,e)=> {
                if (e.KeyCode == Keys.Escape) Close();
                else if (e.KeyCode == Keys.Enter) _btnQuery.PerformClick();
                else if (e.Control && e.KeyCode == Keys.C) { try { Clipboard.SetText(_boxResult.Text); } catch {} }
            };
        }
    }
}
