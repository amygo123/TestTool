using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.WindowsForms;

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

        // 概览：标题/摘要
        readonly Label _lblTitle = new Label();
        readonly Label _lblYesterday = new Label();
        readonly Label _lblSum7d = new Label();

        // 图表（OxyPlot）
        readonly PlotView _pvTrend = new PlotView();   // 7日趋势
        readonly PlotView _pvSizeTop = new PlotView(); // 尺码 Top10
        readonly PlotView _pvColorTop = new PlotView();// 颜色 Top10

        // 明细
        readonly DataGridView _grid = new DataGridView();

        ParsedPayload _parsed = new ParsedPayload();

        public ResultForm(AppConfig cfg)
        {
            _cfg = cfg;
            Width = Math.Max(cfg.window.width, 1100);
            Height = Math.Max(cfg.window.height, 720);
            Text = "StyleWatcher";
            KeyPreview = true;
            StartPosition = FormStartPosition.Manual;

            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, Padding = new Padding(6, 6, 6, 0) };
            var lbl = new Label { Text = "选中文本：", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            _boxInput.Width = Width - 430;
            _btnQuery.Text = "查询(Enter)";
            _btnCopy.Text = "复制原文";
            _btnClose.Text = "隐藏(Esc)";
            top.Controls.AddRange(new Control[] { lbl, _boxInput, _btnQuery, _btnCopy, _btnClose });
            Controls.Add(top);

            _tabs.Dock = DockStyle.Fill;
            Controls.Add(_tabs);

            BuildOverviewTab();
            BuildDetailsTab();
            BuildRawTab();

            var hint = new Label
            {
                Dock = DockStyle.Bottom, Height = 22, ForeColor = Color.Gray,
                Text = $"提示：热键可复用此窗口；Esc 隐藏；Ctrl+C 复制原文；Enter 重新查询。"
            };
            Controls.Add(hint);

            _btnQuery.Click += async (s, e) =>
            {
                SetLoading("查询中...");
                var textNow = _boxInput.Text.Trim();
                var raw = await ApiHelper.QueryAsync(_cfg, textNow);
                ApplyRawText(textNow, raw);
            };
            _btnCopy.Click += (s, e) => { try { Clipboard.SetText(_boxRaw.Text); } catch { } };
            _btnClose.Click += (s, e) => this.Hide();

            KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape) this.Hide();
                else if (e.KeyCode == Keys.Enter) _btnQuery.PerformClick();
                else if (e.Control && e.KeyCode == Keys.C) { try { Clipboard.SetText(_boxRaw.Text); } catch { } }
            };
        }

        void BuildOverviewTab()
        {
            var page = new TabPage("概览");

            // 头部信息
            var header = new Panel { Dock = DockStyle.Top, Height = 100, Padding = new Padding(8) };
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

            // 图表区：上 = 趋势（横跨两列），下 = 尺码Top10 / 颜色Top10
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(8) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 40));

            _pvTrend.Dock = DockStyle.Fill;
            _pvSizeTop.Dock = DockStyle.Fill;
            _pvColorTop.Dock = DockStyle.Fill;

            layout.Controls.Add(_pvTrend, 0, 0);
            layout.SetColumnSpan(_pvTrend, 2);
            layout.Controls.Add(_pvSizeTop, 0, 1);
            layout.Controls.Add(_pvColorTop, 1, 1);

            page.Controls.Add(layout);
            page.Controls.Add(header);
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

        // 供 Tray 侧调用
        public void ShowAndFocusNearCursor(bool topMost)
        {
            var p = Cursor.Position;
            var targetX = Math.Max(0, Math.Min(Screen.PrimaryScreen.WorkingArea.Width - Width, p.X + 12));
            var targetY = Math.Max(0, Math.Min(Screen.PrimaryScreen.WorkingArea.Height - Height, p.Y + 12));
            Location = new Point(targetX, targetY);

            TopMost = topMost;
            if (!Visible) Show();
            WindowState = FormWindowState.Normal;
            Activate();
            BringToFront();
            Focus();
        }
        public void ShowNoActivateAtCursor()
        {
            var p = Cursor.Position;
            var targetX = Math.Max(0, Math.Min(Screen.PrimaryScreen.WorkingArea.Width - Width, p.X + 12));
            var targetY = Math.Max(0, Math.Min(Screen.PrimaryScreen.WorkingArea.Height - Height, p.Y + 12));
            Location = new Point(targetX, targetY);
            if (!Visible) Show();
        }
        public void FocusInput() => _boxInput.Focus();

        public void SetLoading(string message)
        {
            _boxRaw.Text = message;
            _lblTitle.Text = "—";
            _lblYesterday.Text = "";
            _lblSum7d.Text = "";
            _pvTrend.Model = new PlotModel { Title = "最近日销量（加载中）" };
            _pvSizeTop.Model = new PlotModel { Title = "尺码 Top10（加载中）" };
            _pvColorTop.Model = new PlotModel { Title = "颜色 Top10（加载中）" };
            _grid.DataSource = null;
        }

        public void ApplyRawText(string input, string raw)
        {
            _boxInput.Text = input ?? "";
            _boxRaw.Text = Formatter.Prettify(raw ?? "");

            _parsed = PayloadParser.Parse(raw ?? "");

            _lblTitle.Text = string.IsNullOrEmpty(_parsed.Title) ? "—" : _parsed.Title;
            _lblYesterday.Text = string.IsNullOrEmpty(_parsed.Yesterday) ? "" : _parsed.Yesterday;
            _lblSum7d.Text = _parsed.Sum7d.HasValue ? $"近7天销量汇总：{_parsed.Sum7d.Value:N0}" : "";

            RenderTrend();     // 7日趋势
            RenderTopCharts(); // Top10（尺码、颜色）

            // 明细表
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

        // ————— 图表渲染 —————
        void RenderTrend()
        {
            // 按日聚合最近 7 天（保证缺失日期显示 0）
            if (_parsed.Records.Count == 0)
            {
                _pvTrend.Model = new PlotModel { Title = "最近日销量（无数据）" };
                return;
            }

            var minDay = _parsed.Records.Min(x => x.Date);
            var maxDay = _parsed.Records.Max(x => x.Date);
            var start = maxDay.AddDays(-6); // 最近 7 天

            var dict = _parsed.Records
                .GroupBy(r => r.Date)
                .ToDictionary(g => g.Key.Date, g => g.Sum(x => x.Qty));

            var days = Enumerable.Range(0, 7).Select(i => start.Date.AddDays(i)).ToList();

            var model = new PlotModel { Title = "最近 7 天销量趋势" };
            var cat = new CategoryAxis { Position = AxisPosition.Bottom, IsPanEnabled = false, IsZoomEnabled = false };
            foreach (var d in days) cat.Labels.Add(d.ToString("MM-dd"));
            model.Axes.Add(cat);
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, MinorGridlineStyle = LineStyle.Dot });

            var series = new LineSeries { MarkerType = MarkerType.Circle };
            for (int i = 0; i < days.Count; i++)
            {
                var qty = dict.TryGetValue(days[i], out var v) ? v : 0;
                series.Points.Add(new DataPoint(i, qty));
            }
            model.Series.Add(series);
            _pvTrend.Model = model;
        }

        void RenderTopCharts()
        {
            // 尺码 Top10（横向条形图）
            var bySize = _parsed.Records
                .GroupBy(r => string.IsNullOrWhiteSpace(r.Size) ? "(未知)" : r.Size)
                .Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty).Take(10).ToList();

            var modelSize = new PlotModel { Title = "尺码 Top10（7天）" };
            var catSize = new CategoryAxis { Position = AxisPosition.Left, IsPanEnabled = false, IsZoomEnabled = false };
            var serSize = new BarSeries { LabelPlacement = LabelPlacement.Outside, LabelFormatString = "{0}" };
            foreach (var it in bySize)
            {
                catSize.Labels.Add(it.Key);
                serSize.Items.Add(new BarItem(it.Qty));
            }
            modelSize.Axes.Add(catSize);
            modelSize.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinorGridlineStyle = LineStyle.Dot });
            modelSize.Series.Add(serSize);
            _pvSizeTop.Model = modelSize;

            // 颜色 Top10（横向条形图）
            var byColor = _parsed.Records
                .GroupBy(r => string.IsNullOrWhiteSpace(r.Color) ? "(未知)" : r.Color)
                .Select(g => new { Key = g.Key, Qty = g.Sum(x => x.Qty) })
                .OrderByDescending(x => x.Qty).Take(10).ToList();

            var modelColor = new PlotModel { Title = "颜色 Top10（7天）" };
            var catColor = new CategoryAxis { Position = AxisPosition.Left, IsPanEnabled = false, IsZoomEnabled = false };
            var serColor = new BarSeries { LabelPlacement = LabelPlacement.Outside, LabelFormatString = "{0}" };
            foreach (var it in byColor)
            {
                catColor.Labels.Add(it.Key);
                serColor.Items.Add(new BarItem(it.Qty));
            }
            modelColor.Axes.Add(catColor);
            modelColor.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, MinorGridlineStyle = LineStyle.Dot });
            modelColor.Series.Add(serColor);
            _pvColorTop.Model = modelColor;
        }
    }
}
