# StyleWatcherWin（C# WinForms，真正独立 .exe）

**运行时不依赖任何环境**：通过 .NET 8 **自包含**模式发布，生成单个 `StyleWatcher.exe`。

## 一键构建（需要 .NET SDK 仅用于打包）
1. 安装 [.NET 8 SDK](https://dotnet.microsoft.com/download)（仅打包机需要；最终给用户的 exe 不再依赖）
2. 在项目目录执行：
   ```bat
   build_win_x64.bat
   ```
3. 产物：`publish\win-x64\StyleWatcher.exe`（双击即用）

## 使用
- 双击 `StyleWatcher.exe` 后驻留托盘，**选中任意文本**按 **Alt+S** 即查（默认热键，可在 `appsettings.json` 修改为 `Ctrl+Shift+S` 等）。
- 托盘右键：手动输入查询 / 打开配置 / 退出。

## 可配置
- `appsettings.json`：`api_url`、`json_key`、`timeout_seconds`、`hotkey`、窗口设置等。

## 说明
- 热键冲突会弹窗提醒；可改热键后重启。
- 结果窗口支持 Esc 关闭、Ctrl+C 复制、回车再次查询。
