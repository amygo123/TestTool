using System;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace StyleWatcherWin
{
    public class TrayApp : ApplicationContext
    {
        private NotifyIcon _tray;
        private AppConfig _cfg;

        public TrayApp()
        {
            _cfg = AppConfig.Load();
            _tray = new NotifyIcon
            {
                Text = "StyleWatcher",
                Visible = true,
                Icon = System.Drawing.SystemIcons.Information
            };
            var menu = new ContextMenuStrip();
            var miInput = new ToolStripMenuItem("手动输入查询…", null, async (s,e)=> await ShowInputAndQueryAsync());
            var miExit = new ToolStripMenuItem("退出", null, (s,e)=> ExitThread());
            menu.Items.Add(miInput);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(miExit);
            _tray.ContextMenuStrip = menu;

            // 首次提示
            _tray.BalloonTipTitle = "StyleWatcher";
            _tray.BalloonTipText = $"已启动。热键：{_cfg.hotkey}";
            _tray.ShowBalloonTip(3000);
        }

        private async Task ShowInputAndQueryAsync()
        {
            string text = Microsoft.VisualBasic.Interaction.InputBox("输入要查询的文本", "StyleWatcher", "");
            if (string.IsNullOrWhiteSpace(text)) return;
            var result = await ApiHelper.QueryAsync(_cfg, text);
            using var f = new ResultForm(_cfg, text, result);
            f.ShowDialog();
        }
    }
}
