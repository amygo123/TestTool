using System;
using System.Drawing;
using System.Windows.Forms;

namespace StyleWatcherWin
{
    public class ResultForm : Form
    {
        readonly TextBox _boxInput = new TextBox();
        readonly RichTextBox _boxResult = new RichTextBox();
        readonly Button _btnQuery = new Button();
        readonly Button _btnCopy = new Button();
        readonly Button _btnClose = new Button();
        readonly AppConfig _cfg;

        public ResultForm(AppConfig cfg, string input, string result)
        {
            _cfg = cfg;

            Width = Math.Max(cfg.window.width, 900);
            Height = Math.Max(cfg.window.height, 600);
            Text = "StyleWatcher";
            KeyPreview = true;

            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, Padding = new Padding(6,6,6,0) };
            var lbl = new Label { Text = "选中文本：", AutoSize = true, Margin = new Padding(0,8,0,0) };
            _boxInput.Width = Width - 330;
            _boxInput.Text = input;
            _btnQuery.Text = "查询(Enter)";
            _btnCopy.Text = "复制结果";
            _btnClose.Text = "关闭(Esc)";
            top.Controls.AddRange(new Control[]{ lbl, _boxInput, _btnQuery, _btnCopy, _btnClose });
            Controls.Add(top);

            _boxResult.ReadOnly = true;
            _boxResult.BorderStyle = BorderStyle.None;
            _boxResult.Dock = DockStyle.Fill;
            _boxResult.WordWrap = true;
            _boxResult.DetectUrls = false;
            _boxResult.Font = new Font("Microsoft YaHei UI", _cfg.window.fontSize);
            _boxResult.Text = result;
            Controls.Add(_boxResult);

            var hint = new Label {
                Dock = DockStyle.Bottom, Height = 22, ForeColor = Color.Gray,
                Text = $"提示：全局热键 {_cfg.hotkey}；Esc 关闭；Ctrl+C 复制结果；回车可再次查询。"
            };
            Controls.Add(hint);

            _btnQuery.Click += async (s,e)=> {
                _boxResult.Text = "查询中...";
                var textNow = _boxInput.Text.Trim();
                var raw = await ApiHelper.QueryAsync(_cfg, textNow);
                _boxResult.Text = Formatter.Prettify(raw);
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
