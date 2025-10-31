using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

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
        readonly Label _lblTitle = new Label();
        readonly Label _lblYesterday = new Label();
        readonly Label _lblSum7d = new Label();
        readonly Chart _chartDaily = new Chart();
        readonly Chart _chartSizeTop = new Chart();
        readonly Chart _chartColorTop = new Chart();
        readonly DataGridView _grid = new DataGridView();

        ParsedPayload _parsed = new ParsedPayload();

        public ResultForm(AppConfig cfg, string input, string result)
        {
            _cfg = cfg;
            Width = Math.Max(cfg.window.width, 1100);
            Height = Math.Max(cfg.window.height, 720);
            Text = "StyleWatcher";
            KeyPreview = true;

            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, Padding = new Padding(6,6,6,0) };
            var lbl = new Label { Text = "选中文本：", AutoSize = true, Margin = new Padding(0,8,0,0) };
            _boxInput.Width = Width - 430;
            _boxInput.Text = input;
            _btnQuery.Text = "查询(Enter)";
            _btnCopy.Text = "复制原文";
            _btnClose.Text = "关闭(Esc)";
            top.Controls.AddRange(new Control[]{ lbl, _boxInput, _btnQuery, _btnCopy, _btnClose });
            Controls.Add(top);

            _tabs.Dock = DockStyle.Fill;
            Controls.Add(_tabs);

            BuildOverviewTab();
            BuildDetailsTab();
            BuildRawTab();

            ApplyRawText(result);

            var hint = new Label {
                Dock = DockStyle.Bottom, Height = 22, ForeColor = Color.Gray,
                Text = $"提示：全局热键 {_cfg.hotkey}；Esc 关闭；Ctrl+C 复制原文；回车可再次查询。"
            };
            Controls.Add(hint);

            _btnQuery.Click += async (s,e)=> {
                SetLoading();
                var textNow = _boxInput.Text.Trim();
                var raw = await ApiHelper.QueryAsync(_cfg, textNow);
                ApplyRawText(raw);
            };
            _btnCopy.Click += (s,e)=> { try { Clipboard.SetText(_boxRaw.Text); } catch {} };
            _btnClose.Click += (s,e)=> Close();

            KeyDown += (s,e)=> {
                if (e.KeyCode == Keys.Escape) Close();
                else if (e.KeyCode == Keys.Enter) _btnQuery.PerformClick();
                else if (e.Control && e.KeyCode == Keys.C) { try { Clipboard.SetText(_boxRaw.Text); } catch {} }
            };
        }

        void BuildOverviewTab()
        {
            var page = new TabPage("概览");
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(8) };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

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

            SetupChart(_chartDaily, "最近天数销量", SeriesChartType.Column);
            SetupChart(_chartSizeTop, "尺码 Top10（7天）", SeriesChartType.Bar);
            SetupChart(_chartColorTop, "颜色 Top10（7天）", SeriesChartType.Bar);

            panel.Controls.Add(header, 0, 0);
            panel.SetColumnSpan(header, 2);
            panel.Controls.Add(_chartDaily, 0, 1);
            panel.Controls.Add(_chartSizeTop, 1, 1);

            var sub = new Panel { Dock = DockStyle.Bottom, Height = 260, Padding = new Padding(8, 8, 8, 0) };
            _chartColorTop.Dock = DockStyle.Fill;
            sub.Controls.Add(_chartColorTop);
            page.Controls.Add(panel);
            page.Controls.Add(sub);

            _tabs.TabPages.Add(page);
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

        void SetupChart(Chart chart, string title, SeriesChartType type)
        {
            chart.Dock = DockStyle.Fill;
            chart.ChartAreas.Clear();
            chart.Series.Clear();
            var area = new ChartArea();
            area.AxisX.MajorGrid.Enabled = false;
            area.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            area.AxisX.Interval = 1;
            chart.ChartAreas.Add(area);
            var s = new Series { ChartType = type, IsValueShownAsLabel = true };
            chart.Series.Add(s);
            chart.Titles.Clear();
            chart.Titles.Add(title);
        }

        void SetLoading()
        {
            _lblTitle.Text = "查询中...";
            _lblYesterday.Text = "";
            _lblSum7d.Text = "";
            _chartDaily.Series[0].Points.Clear();
            _chartSizeTop.Series[0].Points.Clear();
            _chartColorTop.Series[0].Points.Clear();
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

            var daily = _parsed.Records
                .GroupBy(r => r.Date)
                .Select(g => new { Day = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderBy(x => x.Day)
                .ToList();
            _chartDaily.Series[0].Points.Clear();
            foreach (var d in daily)
                _chartDaily.Series[0].Points.AddXY(d.Day.ToString("MM-dd"), d.Qty);

            var bySize = _parsed.Records
                .GroupBy(r => r.Size)
                .Select(g => new { Size = string.IsNullOrWhiteSpace(g.Key) ? "(未知)" : g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty)
                .Take(10)
                .ToList();
            _chartSizeTop.Series[0].Points.Clear();
            foreach (var it in bySize)
                _chartSizeTop.Series[0].Points.AddXY(it.Size, it.Qty);

            var byColor = _parsed.Records
                .GroupBy(r => r.Color)
                .Select(g => new { Color = string.IsNullOrWhiteSpace(g.Key) ? "(未知)" : g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty)
                .Take(10)
                .ToList();
            _chartColorTop.Series[0].Points.Clear();
            foreach (var it in byColor)
                _chartColorTop.Series[0].Points.AddXY(it.Color, it.Qty);

            var list = _parsed.Records
                .OrderByDescending(r => r.Date)
                .ThenByDescending(r => r.Qty)
                .Select(r => new { 日期 = r.Date.ToString("yyyy-MM-dd"), 款式 = r.Name, 尺码 = string.IsNullOrWhiteSpace(r.Size) ? "(未知)" : r.Size, 颜色 = string.IsNullOrWhiteSpace(r.Color) ? "(未知)" : r.Color, 数量 = r.Qty })
                .ToList();
            _grid.DataSource = list;
        }
    }
}
