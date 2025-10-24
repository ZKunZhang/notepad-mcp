**MCP Notepad（Windows）**

- 这是一个基于 Model Context Protocol 的服务端（stdio），用 PowerShell 自动化控制 Windows 记事本（Notepad）：打开/聚焦、粘贴文本、发送按键、保存/另存为、关闭、列出窗口等。

**环境要求**
- Windows 10/11（交互式桌面会话）
- Node.js 18+
- PowerShell（系统自带 `powershell.exe`）
- Notepad（系统自带）

**安装**
- 进入项目：`cd notepad-mcp`
- 安装依赖：`npm install`

**运行方式（建议由 MCP 客户端启动）**
- 该服务为 stdio 服务器，通常由 MCP 客户端（如 Claude Desktop 或 MCP Inspector）以命令行进程方式启动并通过标准输入/输出通信。
- 手动运行仅用于调试：`npm start`（不会有可读输出，因为需要 MCP 客户端握手）。

**快速自测（MCP Inspector）**
1) 启动 Inspector：`npx @modelcontextprotocol/inspector`
2) UI → “New Connection”
   - Command: `node`
   - Args: `C:\\Users\\admin\\Company\\github\\notepad-mcp\\servers\\notepad-mcp.mjs`
3) 连接后左侧会列出工具；点选调用并填写参数即可。

**在 Claude Desktop 中配置**
- 配置文件（Windows）：`%APPDATA%\Claude\claude_desktop_config.json`
- 在 `mcpServers` 中新增：
```
{
  "mcpServers": {
    "notepad": {
      "command": "node",
      "args": [
        "C:\\Users\\admin\\Company\\github\\notepad-mcp\\servers\\notepad-mcp.mjs"
      ],
      "env": {}
    }
  }
}
```
- 重启 Claude Desktop，新建会话后可让 Claude 使用 `notepad` 工具。

**工具一览**
- `open_notepad` { `filePath?`: string }
  - 打开记事本（可指定文件路径）。返回文本包含 pid。
- `list_notepad_windows` {}
  - 列出当前打开的记事本窗口，包含 `Id`（pid）与 `MainWindowTitle`。
- `paste_text` { `text`: string, `pid?`: number }
  - 聚焦指定或最近的记事本窗口，使用剪贴板 + Ctrl+V 粘贴文本（对特殊字符更鲁棒）。
- `send_keys` { `keys`: string, `pid?`: number }
  - 发送按键，采用 .NET `SendKeys` 语法，例如：`^s`（Ctrl+S），`%{F4}`（Alt+F4），`{ENTER}`。
- `save_file` { `filePath?`: string, `pid?`: number }
  - 不带 `filePath`：执行保存（Ctrl+S）。带 `filePath`：执行“另存为”并回车。
- `close_notepad` { `pid?`: number, `dontSave?`: boolean }
  - 关闭窗口；`dontSave` 省略或为 `true` 时尝试选择“不保存”。

**端到端示例（用 Inspector）**
1) 打开记事本：
   - 工具：`open_notepad`
   - 参数：`{}`（留空）
   - 期望响应：`Opened Notepad (pid=12345)`
2) 粘贴文本：
   - 工具：`paste_text`
   - 参数：`{"text":"Hello from MCP Notepad 🚀"}`
3) 另存为：
   - 工具：`save_file`
   - 参数：`{"filePath":"C:\\Users\\admin\\Desktop\\demo.txt"}`
4) 关闭：
   - 工具：`close_notepad`
   - 参数：`{"dontSave":true}` 或 `{"dontSave":false}`

**注意事项 / 常见问题**
- 前台窗口与焦点：SendKeys 仅作用于前台窗口，请确保桌面未被锁屏/遮挡；必要时手动点一下 Notepad 窗口再调用工具。
- 非英文系统的快捷键：关闭时的“不保存/保存”的快捷键字母在不同语言下可能不同。若 `close_notepad` 处理不生效，可用 `send_keys` 组合 `{TAB}`/方向键 + `{ENTER}` 实现。
- 路径与权限：另存为请使用存在的目录的绝对路径；如遇 UAC/权限问题，建议保存到用户目录下。
- 剪贴板占用：`paste_text` 使用剪贴板；若有剪贴板管理器或安全策略拦截，可能失败，可改用 `send_keys` 逐字输入（但对特殊字符较脆弱）。
- 多窗口选择：未传 `pid` 时默认聚焦最近的 Notepad。可先 `list_notepad_windows` 获取 pid 以精准控制。
- PowerShell 限制：脚本以 `-ExecutionPolicy Bypass` 运行，如企业策略限制严格，需与管理员协作放行。

**开发/运行提示**
- 入口脚本：`servers/notepad-mcp.mjs`
- 本地直接运行（调试）：`npm start` 或 `node servers/notepad-mcp.mjs`
- Node 要求：>= 18

**已知限制**
- 依赖 GUI 前台与键盘布局，无法在无头/锁屏环境工作。
- SendKeys 不保证时序绝对稳定，复杂窗口切换下需适当 `Start-Sleep`（已在脚本中做基础等待）。
- 未提供富 UI 自动化（如控件枚举/精准点击）；本项目定位为“轻量快捷自动化”。

