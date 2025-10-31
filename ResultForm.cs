using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace StyleWatcherWin
{
    public class ResultForm : Form
    {
        readonly TextBox _boxInput = new TextBox();
        readonly RichTextBox _boxRaw = new RichTextBox();
        readonly Button _btnQuery = new Button();
        readonly Button _btnCopy = new Button();
        readonly Button _btnClose = new Button();
        readonly AppConfig _cfg;

        readonly TabControl _tabs = new TabControl();

        // 概览页
        readonly Label _lblTitle = new Label();
        readonly Label _lblYesterday = new Label();
        readonly Label _lblSum7d = new Label();
        readonly ListView _lvSizeTop = new ListView();
        readonly ListView _lvColorTop = new ListView();

        // 明细
        readonly DataGridView _grid = new DataGridView();

        ParsedPayload _parsed = new ParsedPayload();

        public ResultForm(AppConfig cfg, string input, string result)
        {
            _cfg = cfg;
            Width = Math.Max(cfg.window.width, 1000);
            Height = Math.Max(cfg.window.height, 700);
            Text = "StyleWatcher";
            KeyPreview = true;

            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, Padding = new Padding(6, 6, 6, 0) };
            var lbl = new Label { Text = "选中文本：", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            _boxInput.Width = Width - 430;
            _boxInput.Text = input;
            _btnQuery.Text = "查询(Enter)";
            _btnCopy.Text = "复制原文";
            _btnClose.Text = "关闭(Esc)";
            top.Controls.AddRange(new Control[] { lbl, _boxInput, _btnQuery, _btnCopy, _btnClose });
            Controls.Add(top);

            _tabs.Dock = DockStyle.Fill;
            Controls.Add(_tabs);

            BuildOverviewTab();
            BuildDetailsTab();
            BuildRawTab();

            ApplyRawText(result);

            var hint = new Label
            {
                Dock = DockStyle.Bottom, Height = 22, ForeColor = Color.Gray,
                Text = $"提示：全局热键 {_cfg.hotkey}；Esc 关闭；Ctrl+C 复制原文；回车可再次查询。"
            };
            Controls.Add(hint);

            _btnQuery.Click += async (s, e) =>
            {
                SetLoading();
                var textNow = _boxInput.Text.Trim();
                var raw = await ApiHelper.QueryAsync(_cfg, textNow);
                ApplyRawText(raw);
            };
            _btnCopy.Click += (s, e) => { try { Clipboard.SetText(_boxRaw.Text); } catch { } };
            _btnClose.Click += (s, e) => Close();

            KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape) Close();
                else if (e.KeyCode == Keys.Enter) _btnQuery.PerformClick();
                else if (e.Control && e.KeyCode == Keys.C) { try { Clipboard.SetText(_boxRaw.Text); } catch { } }
            };
        }

        void BuildOverviewTab()
        {
            var page = new TabPage("概览");
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(8) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // 头部信息
            var header = new Panel { Dock = DockStyle.Fill };
            _lblTitle.Font = new Font("Microsoft YaHei UI", 12, FontStyle.Bold);
            _lblYesterday.Font = new Font("Microsoft YaHei UI", 10);
            _lblSum7d.Font = new Font("Microsoft YaHei UI", 10);
            _lblTitle.AutoEllipsis = true;
            _lblTitle.Dock = DockStyle.Top;
            _lblYesterday.Dock = DockStyle.Top;
            _lblSum7d.Dock = DockStyle.Top;
            header.Controls.Add(_lblSum7d);
            header.Controls.Add(_lblYesterday);
            header.Controls.Add(_lblTitle);

            // Top10 列表
            SetupList(_lvSizeTop, "尺码 Top10（7天）");
            SetupList(_lvColorTop, "颜色 Top10（7天）");

            layout.Controls.Add(header, 0, 0);
            layout.SetColumnSpan(header, 2);

            layout.Controls.Add(_lvSizeTop, 0, 1);
            layout.Controls.Add(_lvColorTop, 1, 1);

            page.Controls.Add(layout);
            _tabs.TabPages.Add(page);
        }

        void SetupList(ListView lv, string title)
        {
            lv.Dock = DockStyle.Fill;
            lv.View = View.Details;
            lv.FullRowSelect = true;
            lv.GridLines = true;
            lv.Columns.Add(title, 220);
            lv.Columns.Add("数量", 80, HorizontalAlignment.Right);
        }

        void BuildDetailsTab()
        {
            var page = new TabPage("明细");
            _grid.Dock = DockStyle.Fill;
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.RowHeadersVisible = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            page.Controls.Add(_grid);
            _tabs.TabPages.Add(page);
        }

        void BuildRawTab()
        {
            var page = new TabPage("原文");
            _boxRaw.ReadOnly = true;
            _boxRaw.BorderStyle = BorderStyle.None;
            _boxRaw.Dock = DockStyle.Fill;
            _boxRaw.WordWrap = true;
            _boxRaw.ScrollBars = RichTextBoxScrollBars.Vertical;
            _boxRaw.Font = new Font("Microsoft YaHei UI", _cfg.window.fontSize);
            page.Controls.Add(_boxRaw);
            _tabs.TabPages.Add(page);
        }

        void SetLoading()
        {
            _lblTitle.Text = "查询中...";
            _lblYesterday.Text = "";
            _lblSum7d.Text = "";
            _lvSizeTop.Items.Clear();
            _lvColorTop.Items.Clear();
            _grid.DataSource = null;
            _boxRaw.Text = "查询中...";
        }

        void ApplyRawText(string raw)
        {
            _boxRaw.Text = Formatter.Prettify(raw);
            _parsed = PayloadParser.Parse(raw);

            _lblTitle.Text = string.IsNullOrEmpty(_parsed.Title) ? "—" : _parsed.Title;
            _lblYesterday.Text = string.IsNullOrEmpty(_parsed.Yesterday) ? "" : _parsed.Yesterday;
            _lblSum7d.Text = _parsed.Sum7d.HasValue ? $"近7天销量汇总：{_parsed.Sum7d.Value:N0}" : "";

            // Top10 根据尺码
            _lvSizeTop.Items.Clear();
            var bySize = _parsed.Records
                .GroupBy(r => string.IsNullOrWhiteSpace(r.Size) ? "(未知)" : r.Size)
                .Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty).Take(10).ToList();
            foreach (var it in bySize)
            {
                var li = new ListViewItem(it.Key);
                li.SubItems.Add(it.Qty.ToString());
                _lvSizeTop.Items.Add(li);
            }

            // Top10 根据颜色
            _lvColorTop.Items.Clear();
            var byColor = _parsed.Records
                .GroupBy(r => string.IsNullOrWhiteSpace(r.Color) ? "(未知)" : r.Color)
                .Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty).Take(10).ToList();
            foreach (var it in byColor)
            {
                var li = new ListViewItem(it.Key);
                li.SubItems.Add(it.Qty.ToString());
                _lvColorTop.Items.Add(li);
            }

            // 明细
            var list = _parsed.Records
                .OrderByDescending(r => r.Date).ThenByDescending(r => r.Qty)
                .Select(r => new
                {
                    日期 = r.Date.ToString("yyyy-MM-dd"),
                    款式 = r.Name,
                    尺码 = string.IsNullOrWhiteSpace(r.Size) ? "(未知)" : r.Size,
                    颜色 = string.IsNullOrWhiteSpace(r.Color) ? "(未知)" : r.Color,
                    数量 = r.Qty
                }).ToList();
            _grid.DataSource = list;
        }
    }
}
