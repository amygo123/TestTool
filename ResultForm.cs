using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Collections.Generic;

namespace StyleWatcherWin
{
    public class ResultForm : Form
    {
        readonly TextBox _boxInput = new TextBox();
        readonly TextBox _boxResult = new TextBox();
        readonly Button _btnQuery = new Button();
        readonly Button _btnCopy = new Button();
        readonly Button _btnClose = new Button();
        readonly Button _btnRefreshInv = new Button();

        readonly AppConfig _cfg;
        readonly InventoryClient _invClient;

        // 库存数据缓存（按输入）
        static readonly Dictionary<string,(DateTime ts, List<InventoryRow> rows)> _cache = new();

        // UI：库存区域
        readonly Panel _invPanel = new Panel();
        readonly Label _lblInvTitle = new Label();
        readonly Label _lblInvUpdate = new Label();
        readonly DataGridView _gridMain = new DataGridView();
        readonly FlowLayoutPanel _othersPanel = new FlowLayoutPanel();
        readonly Label _lblInvSummary = new Label();

        public ResultForm(AppConfig cfg, string input, string result)
        {
            _cfg = cfg;
            _invClient = new InventoryClient(cfg);

            Text = "StyleWatcher";
            Width = cfg.window.width;
            Height = Math.Max(cfg.window.height, 520);
            KeyPreview = true;

            // 顶部操作区
            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, WrapContents = false, AutoScroll = true };
            var lbl = new Label { Text = "选中文本：", AutoSize = true, Margin = new Padding(6,10,0,0) };
            _boxInput.Width = cfg.window.width - 460;
            _boxInput.Text = input;
            _btnQuery.Text = "查询";
            _btnCopy.Text = "复制结果";
            _btnClose.Text = "关闭(Esc)";
            _btnRefreshInv.Text = "刷新库存";
            top.Controls.AddRange(new Control[]{ lbl, _boxInput, _btnQuery, _btnCopy, _btnRefreshInv, _btnClose });
            Controls.Add(top);

            // 结果区
            _boxResult.Multiline = true;
            _boxResult.ReadOnly = true;
            _boxResult.ScrollBars = ScrollBars.Vertical;
            _boxResult.Dock = DockStyle.Top;
            _boxResult.Font = new Font("Consolas", cfg.window.fontSize);
            _boxResult.Height = 180;
            _boxResult.Text = result;
            Controls.Add(_boxResult);

            // 库存区域（同页签，默认展开）
            BuildInventoryArea();
            Controls.Add(_invPanel);

            var hint = new Label { Dock = DockStyle.Bottom, Height = 22, ForeColor = Color.Gray, Text = $"提示：热键 {cfg.hotkey}；Esc 关闭；Ctrl+C 复制结果；Enter 再次查询；Ctrl+Enter 刷新库存。" };
            Controls.Add(hint);

            // 事件绑定
            _btnQuery.Click += async (s,e)=> await RunMainQueryAsync();
            _btnCopy.Click += (s,e)=> { try { Clipboard.SetText(_boxResult.Text); } catch {} };
            _btnClose.Click += (s,e)=> Close();
            _btnRefreshInv.Click += async (s,e)=> await LoadInventoryAsync(force:true);

            KeyDown += async (s,e)=> {
                if (e.KeyCode == Keys.Escape) Close();
                else if (e.KeyCode == Keys.Enter && !e.Control) await RunMainQueryAsync();
                else if (e.Control && e.KeyCode == Keys.Enter) await LoadInventoryAsync(force:true);
                else if (e.Control && e.KeyCode == Keys.C) { try { Clipboard.SetText(_boxResult.Text); } catch {} }
            };

            // 初次显示时拉一次库存（可见且用户主动打开窗口的语义）
            Shown += async (s,e)=> await LoadInventoryAsync(force:false);
        }

        void BuildInventoryArea()
        {
            _invPanel.Dock = DockStyle.Fill;
            _invPanel.Padding = new Padding(8,6,8,6);

            var header = new Panel { Dock = DockStyle.Top, Height = 28 };
            _lblInvTitle.Text = "库存（颜色 × 尺码）";
            _lblInvTitle.AutoSize = true; _lblInvTitle.Font = new Font(Font, FontStyle.Bold);
            _lblInvUpdate.Text = "最近更新：--"; _lblInvUpdate.AutoSize = true; _lblInvUpdate.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            header.Controls.Add(_lblInvTitle);
            header.Controls.Add(_lblInvUpdate);
            _lblInvUpdate.Left = _invPanel.Width - 220; _lblInvUpdate.Top = 6; header.Resize += (s,e)=>{ _lblInvUpdate.Left = header.Width - _lblInvUpdate.Width - 4; };
            _invPanel.Controls.Add(header);

            // 摘要
            _lblInvSummary.Dock = DockStyle.Top; _lblInvSummary.Height = 22; _lblInvSummary.ForeColor = Color.DimGray;
            _invPanel.Controls.Add(_lblInvSummary);

            // 主矩阵
            _gridMain.Dock = DockStyle.Top;
            _gridMain.ReadOnly = true;
            _gridMain.AllowUserToAddRows = false;
            _gridMain.AllowUserToDeleteRows = false;
            _gridMain.RowHeadersVisible = false;
            _gridMain.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            _gridMain.Height = 220;
            _gridMain.CellFormatting += GridMain_CellFormatting;
            _invPanel.Controls.Add(_gridMain);

            // 其它仓
            var otherTitle = new Label { Text = "其它仓（展开查看矩阵）", Dock = DockStyle.Top, Height = 20, ForeColor = Color.Gray };
            _invPanel.Controls.Add(otherTitle);
            _othersPanel.Dock = DockStyle.Fill;
            _othersPanel.AutoScroll = true;
            _othersPanel.WrapContents = false;
            _othersPanel.FlowDirection = FlowDirection.TopDown;
            _invPanel.Controls.Add(_othersPanel);
        }

        async Task RunMainQueryAsync()
        {
            _boxResult.Text = "查询中...";
            var textNow = _boxInput.Text.Trim();
            _boxResult.Text = await ApiHelper.QueryAsync(_cfg, textNow);
            await LoadInventoryAsync(force:true); // 主查询后同步刷新库存
        }

        async Task LoadInventoryAsync(bool force)
        {
            var key = _boxInput.Text.Trim();
            if (string.IsNullOrEmpty(key)) return;

            // 缓存
            if (!force && _cache.TryGetValue(key, out var cached))
            {
                if ((DateTime.Now - cached.ts).TotalSeconds < _cfg.inventory_cache_ttl_seconds)
                {
                    BindInventory(cached.rows);
                    return;
                }
            }

            try
            {
                _lblInvUpdate.Text = "正在加载…";
                var rows = await _invClient.FetchAsync(key, CancellationToken.None);
                _cache[key] = (DateTime.Now, rows);
                BindInventory(rows);
                _lblInvUpdate.Text = $"最近更新：{DateTime.Now:HH:mm:ss}";
            }
            catch (Exception ex)
            {
                _lblInvUpdate.Text = $"库存获取失败：{ex.Message}";
            }
        }

        void BindInventory(List<InventoryRow> rows)
        {
            if (rows == null || rows.Count == 0) { _gridMain.DataSource = null; _othersPanel.Controls.Clear(); _lblInvSummary.Text = "无库存数据"; return; }

            // 统计各仓可用合计
            var byWh = rows.GroupBy(r => r.Warehouse)
                           .Select(g => new { Warehouse = g.Key, InSum = g.Sum(x=>x.QtyIn), OutSum = g.Sum(x=>x.QtyOut), Rows = g.ToList() })
                           .OrderByDescending(x => x.OutSum)
                           .ToList();

            // 主仓 = 可用最多的前三个（按你的要求）
            var top3 = byWh.Take(3).ToList();
            var main = top3.FirstOrDefault();

            // 主仓矩阵
            if (main != null)
            {
                var dt = PivotHelper.ToMatrix(main.Rows);
                _gridMain.DataSource = dt;
                _lblInvTitle.Text = $"库存（主仓：{main.Warehouse}）";
            }
            else
            {
                _gridMain.DataSource = null;
                _lblInvTitle.Text = "库存";
            }

            // 摘要（主仓 + 其它仓合计 + 异常）
            int lowTh = Math.Max(1, _cfg.inventory_low_threshold);
            int mainLow = main?.Rows.Count(r=> r.QtyOut <= lowTh && r.QtyOut >= 0) ?? 0;
            int mainBad = main?.Rows.Count(r=> r.QtyOut < 0) ?? 0;
            int othersOut = byWh.Skip(1).Sum(x=>x.OutSum);
            _lblInvSummary.Text = $"主仓可用合计：{main?.OutSum ?? 0}｜低库存(≤{lowTh})：{mainLow}｜异常(<0)：{mainBad}｜其它仓可用合计：{othersOut}";

            // 其它仓列表（不使用下拉/筛选，仅折叠式列表项）
            _othersPanel.Controls.Clear();
            for (int i=0;i<byWh.Count;i++)
            {
                var x = byWh[i];
                if (main != null && x.Warehouse == main.Warehouse) continue;
                var panel = new Panel { Width = _othersPanel.ClientSize.Width - 28, Height = 28, BackColor = (i%2==0? Color.White: Color.WhiteSmoke) };
                panel.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
                var lab = new Label { AutoSize = true, Text = $"{x.Warehouse}｜在库/可用：{x.InSum}/{x.OutSum}｜颜色：{x.Rows.Select(r=>r.Color).Distinct().Count()}｜尺码：{x.Rows.Select(r=>r.Size).Distinct().Count()}", Left = 6, Top = 6 };
                var btn = new Button { Text = "查看矩阵", Width = 88, Height = 22, Left = panel.Width - 96, Top = 3, Anchor = AnchorStyles.Right | AnchorStyles.Top };
                btn.Click += (s,e)=> ShowWarehouseMatrix(x.Warehouse, x.Rows);
                panel.Controls.Add(lab); panel.Controls.Add(btn);
                _othersPanel.Controls.Add(panel);
            }
        }

        void ShowWarehouseMatrix(string warehouse, List<InventoryRow> rows)
        {
            var dt = PivotHelper.ToMatrix(rows);
            var f = new Form { Text = $"{warehouse}：库存矩阵", StartPosition = FormStartPosition.CenterParent, Width = Math.Max(800, Width - 60), Height = Math.Max(500, Height - 60) };
            var g = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells, DataSource = dt, RowHeadersVisible = false, AllowUserToAddRows=false, AllowUserToDeleteRows=false };
            g.CellFormatting += GridMain_CellFormatting;
            f.Controls.Add(g);
            f.ShowDialog(this);
        }

        void GridMain_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex <= 0) return;
            var grid = (DataGridView)sender!;
            var val = grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
            if (string.IsNullOrWhiteSpace(val)) return;
            // 形如 "12 / 8"
            var parts = val.Split('/');
            if (parts.Length != 2) return;
            if (!int.TryParse(parts[1].Trim(), out var outQty)) return;
            if (outQty < 0) { e.CellStyle.BackColor = Color.MistyRose; e.CellStyle.ForeColor = Color.DarkRed; }
            else if (outQty == 0) { e.CellStyle.ForeColor = Color.Gray; }
            else if (outQty <= _cfg.inventory_low_threshold) { e.CellStyle.BackColor = Color.LemonChiffon; }
        }
    }
}
