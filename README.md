# webmail-cli

本地单用户 Outlook Web CLI。默认通过 `playwright-core` 操作系统中已经安装的 Edge、Chrome 或 Chromium，
支持 Windows、macOS 和 Linux；Ego Lite 后端保留为启动失败时的兼容 fallback。CLI 操作
`https://partner.outlook.cn/mail/`，不使用 Graph、EWS、IMAP、SMTP，
也不读取 Cookie、Bearer Token 或浏览器 Storage。

## 快速开始

运行环境：Node.js 24+，以及已经安装的 Edge、Chrome 或 Chromium。`playwright-core` 不会自动下载浏览器。

```bash
git clone https://github.com/guanglinhuang99/outlookwebmail-cli.git
cd outlookwebmail-cli
npm install
npm run build
npm link
webmail --version
```

`npm link` 会在本机注册 `webmail` 和 `webmail-mcp` 两个命令。如果不想使用全局命令，可将下面的
`webmail` 替换为 `node dist/cli.js`，将 `webmail-mcp` 替换为 `node dist/mcp/stdio.js`。

第一次使用时先检查状态：

```bash
webmail status --json
```

CLI 会使用专用持久化 Profile 打开 `https://partner.outlook.cn/mail/`。如果返回 `AUTH_REQUIRED`，
请在打开的浏览器窗口中手工登录，然后再次运行 `webmail status --json`。登录成功后建议执行：

```bash
webmail doctor --json
webmail folders --json
webmail list --json
```

不要把 `WEBMAIL_PROFILE_DIR` 指向日常浏览器的默认 Profile。CLI 不读取用户名、密码、验证码、Cookie
或 localStorage。

### 使用发布安装包

不需要克隆源码。到 GitHub Release 下载对应压缩包并完整解压：

- macOS：下载 `webmail-cli-<版本>-macos.zip`，双击 `install.command`。
- Windows：下载 `webmail-cli-<版本>-windows.zip`，双击 `install.cmd`。

两个安装器都会把 CLI 安装到当前用户目录，不要求管理员权限；安装时需要联网下载生产依赖。
安装完成后打开一个新终端，运行 `webmail status --json`。安装包附带 `SHA256SUMS.txt`，可用于校验下载文件。

## CLI 使用

一个典型的读取流程是先列目录，再列邮件，最后使用返回的 `stableId` 读取邮件：

```bash
# 1. 查看 Inbox 下的目录
webmail folders --json

# 2. 列出今天 Inbox 中的邮件
webmail list --json

# 也可以指定日期和目录
webmail list --date 2026-08-20 --dir "收件箱/投后" --json

# 3. 使用列表返回的 stableId 读取正文和附件元数据
webmail read m_xxxxxxxxxxxxxxxxxxxx --json

# 4. 下载所有附件
webmail download-all m_xxxxxxxxxxxxxxxxxxxx --output ./downloads --json
```

`list` 的 `--date` 默认是上海时区的今天，`--dir` 默认是 Inbox。返回 `nextCursor` 时，把它原样传给
下一次命令即可继续翻页：

```bash
webmail list --date 2026-08-20 --limit 20 --json
webmail list --date 2026-08-20 --limit 20 --cursor '<nextCursor>' --json
```

列表中的每封邮件同时返回数字短 `id` 和形如 `m_...` 的不透明 `stableId`。数字 `id` 只代表当前列表页；
`stableId` 会在列表刷新和分页后保留，脚本和 Agent 应优先使用 `stableId`。发送、放弃草稿、移动、归档
或删除成功后会清空当前 Session，后续操作前需要重新执行 `list`、`search`、`inbox` 或 `drafts`。

常用写操作：

```bash
# 默认只创建草稿，由用户在浏览器中检查和发送
webmail compose --to "alice@example.com" --subject "周报" --content "请查收。" --json
webmail reply m_xxxxxxxxxxxxxxxxxxxx --content "收到，谢谢。" --json
webmail reply m_xxxxxxxxxxxxxxxxxxxx --content "请相关同事查看。" --replyall true --json

# 自动发送必须显式关闭 draft，并提供本次操作唯一的 request ID
webmail reply m_xxxxxxxxxxxxxxxxxxxx --content "收到，谢谢。" \
  --draft false --request-id reply-20260820-001 --json

# 移动和删除还必须显式确认
webmail move m_xxxxxxxxxxxxxxxxxxxx "投后" --yes --request-id move-20260820-001 --json
webmail delete m_xxxxxxxxxxxxxxxxxxxx --yes --request-id delete-20260820-001 --json
```

同一个 `request-id` 和同一组参数可以安全重试，不会重复发送或重复操作；不要把同一个 `request-id` 用于
不同参数。状态保存在 `~/.webmail-cli/mutations.json`，脱敏审计保存在
`~/.webmail-cli/audit.jsonl`。

Obsidian 增量同步和新邮件监控：

```bash
# 重复运行会跳过未变化邮件，并更新已有 Markdown
webmail sync-obsidian --date 2026-08-20 --dir "收件箱/投后" \
  --output /path/to/obsidian-vault/邮件 --json

# 首次运行只建立基线，随后每 30 秒输出新增邮件 JSONL
webmail watch --dir "收件箱/投后" --interval 30

# 首次运行时也输出当前已有邮件
webmail watch --emit-existing --interval 30
```

完整命令示例：

```bash
webmail status
webmail status --json
webmail doctor --json
webmail inspect
webmail inspect --json
webmail inbox --limit 20 --json
webmail inbox --unread --json
webmail inbox --dir "投后" --limit 20 --json
webmail list --json
webmail list --date 2026-08-19 --dir "收件箱/投后" --json
webmail list --from-date 2026-08-01 --to-date 2026-08-20 --sender "risk" --unread --limit 20 --json
webmail list --date 2026-08-20 --has-attachments --cursor '<nextCursor>' --json
webmail today --json
webmail today --dir "收件箱/投后" --json
webmail search "风险报告" --limit 20 --json
webmail folders --json
webmail inspect-message --json
webmail read m_xxxxxxxxxxxxxxxxxxxx --json
webmail attachments 3 --json
webmail compose --to "alice@example.com" --cc "team@example.com" --subject "周报" --content "请查收。" --attach ./weekly.xlsx --json
webmail compose --to "alice@example.com" --subject "周报" --content "请查收。" --draft false --request-id compose-20260820-001 --json
webmail forward 3 --to "alice@example.com" --content "供参考。" --json
webmail reply 3 --content "收到，谢谢。" --json
webmail reply 3 --content "请相关同事一并查看。" --replyall true --draft true --json
webmail reply 3 --content "收到，谢谢。" --draft false --request-id reply-20260820-001 --json
webmail drafts --limit 20 --json
webmail draft-read 1 --json
webmail draft-update 1 --subject "更新后的主题" --content "更新后的正文" --attach ./appendix.pdf --json
webmail draft-send 1 --request-id draft-send-20260820-001 --json
webmail draft-discard 1 --yes --request-id draft-discard-20260820-001 --json
webmail mark-read 3 --json
webmail mark-unread 3 --json
webmail flag 3 --state true --json
webmail category 3 --category "项目A" --applied true --json
webmail archive 3 --yes --request-id archive-20260820-001 --json
webmail conversation 3 --json
webmail download 3 1 --output ./downloads --json
webmail download-all 3 --output ./downloads --json
webmail export 3 --output /path/to/obsidian-vault/邮件 --json
webmail export-batch --date 2026-08-20 --dir "收件箱/投后" --output /path/to/obsidian-vault/邮件 --json
webmail sync-obsidian --date 2026-08-20 --dir "收件箱/投后" --output /path/to/obsidian-vault/邮件 --json
webmail watch --dir "收件箱/投后" --interval 30
webmail move 3 "投后" --yes --request-id move-20260820-001 --json
webmail delete 3 --yes --request-id delete-20260820-001 --json
```

### 浏览器后端配置

默认 `auto` 会优先使用 Playwright，仅在 Playwright 启动失败时尝试 Ego Lite fallback。

macOS/Linux：

```bash
WEBMAIL_BACKEND=auto webmail status --json
export WEBMAIL_BACKEND=playwright
export WEBMAIL_BROWSER=edge
webmail status --json
```

Windows PowerShell（通常会自动发现 Edge，也可显式指定）：

```powershell
$env:WEBMAIL_BACKEND = "playwright"
$env:WEBMAIL_BROWSER = "edge"
webmail status --json
```

显式使用旧 Ego Lite 后端：

```bash
WEBMAIL_BACKEND=ego-lite webmail status --json
```

可用环境变量：`WEBMAIL_BACKEND=auto|playwright|ego-lite`、
`WEBMAIL_BROWSER=auto|edge|chrome|chromium`、`WEBMAIL_EXECUTABLE_PATH`、`WEBMAIL_PROFILE_DIR`、
`WEBMAIL_HEADLESS`、`WEBMAIL_CDP_ENDPOINT`、`WEBMAIL_BROWSER_TIMEOUT_MS` 和 `WEBMAIL_URL`。
macOS/Linux 使用 `export WEBMAIL_BACKEND=playwright`；Windows `cmd.exe` 使用
`set WEBMAIL_BACKEND=playwright`。环境变量只需要在自动发现失败或需要固定后端时配置。

- `folders`：递归展开并列出 Inbox 下的全部子目录，返回可供 `--dir` 使用的 `name` 和 `path`。
- `doctor`：只读检查 Node.js、当前浏览器后端、Outlook 登录状态和关键 DOM；任一检查失败时进程返回非零退出码。
- `inbox`：不传 `--dir` 时列 Inbox；传入时列指定 Inbox 子目录。完整 `path` 优先于名称。
- `list`：按日期和目录分页列出邮件。`--date` 省略或为空时使用上海时区的今天；也可成对使用
  `--from-date/--to-date`。`--dir` 省略或为空时使用 Inbox；`--sender`、`--subject`、`--unread`、
  `--has-attachments` 可组合筛选。每页最多 100 封，继续翻页时原样传回 `nextCursor`；游标与目录和筛选条件绑定。
- `today`：不传 `--dir` 时列 Inbox 当日邮件；传入时列指定子目录的当日邮件。
- `read`：返回邮件头、纯文本正文和附件元数据。
- `attachments`：只返回附件名称和大小。
- `compose`：支持 `to/cc/bcc/subject/content` 和可重复的 `--attach`。默认保存草稿并把编辑窗口交给用户；
  只有 `--draft false` 且提供唯一 `--request-id` 时才自动发送。
- `forward`：在保留原邮件附件的基础上转发，可追加本地附件；默认保存草稿，自动发送规则与 `compose` 相同。
- `reply`：`--content` 提供正文；`--replyall` 默认 `false`，设为 `true` 时全部答复；`--draft` 默认 `true`，生成并打开草稿交给用户手工发送。只有显式传入 `--draft false` 并提供唯一 `--request-id` 才会自动点击发送。
- `drafts/draft-read/draft-update/draft-send/draft-discard`：列出、读取、局部修改、发送或放弃草稿。
  发送需要唯一 `--request-id`；放弃还必须显式提供 `--yes`。
- `mark-read/mark-unread/flag/category`：设置已读状态、旗标和已有 Outlook 分类；重复设置相同状态会返回
  `changed: false`，不会反向切换。
- `archive`：移动到 Outlook 的“存档”目录，要求 `--yes` 和唯一 `--request-id`。
- `conversation`：返回阅读窗格中已经加载的整段会话；`complete: false` 表示页面仍存在未展开的会话内容，Agent 不应把结果当作完整线程。
- `download`：将指定附件下载到本地目录并返回绝对路径和字节数。
- `download-all`：下载一封邮件的全部附件，计算每个文件的 SHA-256；遇到本地同名文件时自动增加序号，绝不覆盖原文件。
- `export`：将一封邮件导出为带 YAML 属性的 Obsidian Markdown；全部附件下载到相对的 `attachments/<邮件标识>/` 目录，并在 Markdown 中生成可点击的相对链接。同名导出自动增加序号，不覆盖已有文件。
- `export-batch`：复用 `list` 的日期、目录和筛选参数，逐页导出全部匹配邮件及附件；最多处理 100 页，防止异常页面导致无限循环。
- `sync-obsidian`：增量同步指定日期/范围和目录中的邮件。使用目标目录下的
  `.webmail-cli-index.json` 按 `stableId` 去重，未变化邮件不会重复下载；变化邮件更新原 Markdown，
  并保留 `<!-- webmail-cli:managed-end -->` 后的用户笔记。图片附件会内嵌显示，所有附件计算
  SHA-256 并汇总到 `_attachments-index.md`。
- `watch`：轮询今天的邮件并把新增邮件逐行输出为 JSONL。首次运行默认只建立基线；
  `--emit-existing` 可输出现有邮件。状态默认保存在 `~/.webmail-cli/watch-state.json`，按日期和目录隔离；
  `Ctrl+C` 安全停止。`--iterations` 可用于计划任务和测试，`0` 表示持续运行。
- `move`：只接受移动菜单中完全匹配的目录名，并要求 `--yes` 与唯一 `--request-id`。
- `delete`：将邮件移入“已删除邮件”，并要求 `--yes` 与唯一 `--request-id`。

自动新建/转发/回复、发送或放弃草稿、移动、归档和删除使用 request ID 防止 Agent 重试造成重复动作。
审计不记录邮件主题、地址、正文或回复内容。

## MCP Server

MCP Server 通过 stdio 向支持 MCP 的 Agent 提供标准邮件工具。先按照“快速开始”完成安装和构建；
MCP 进程通常应由 MCP 客户端启动，不需要手工常驻运行。

如果已经执行 `npm link`，MCP 客户端可以直接启动 `webmail-mcp`：

```json
{
  "mcpServers": {
    "outlook-webmail": {
      "command": "webmail-mcp",
      "env": {
        "WEBMAIL_BACKEND": "auto"
      }
    }
  }
}
```

如果 MCP 客户端找不到 npm 全局命令，使用 Node 和构建文件的绝对路径。macOS/Linux 示例：

```json
{
  "mcpServers": {
    "outlook-webmail": {
      "command": "node",
      "args": ["/absolute/path/to/webmail-cli/dist/mcp/stdio.js"],
      "env": { "WEBMAIL_BACKEND": "auto" }
    }
  }
}
```

Windows 示例（JSON 中的反斜杠需要写成 `\\`）：

```json
{
  "mcpServers": {
    "outlook-webmail": {
      "command": "C:\\Program Files\\nodejs\\node.exe",
      "args": ["C:\\work\\outlookwebmail-cli\\dist\\mcp\\stdio.js"],
      "env": {
        "WEBMAIL_BACKEND": "playwright",
        "WEBMAIL_BROWSER": "edge"
      }
    }
  }
}
```

修改 MCP 配置后重启 MCP 客户端。第一次调用 `status` 时，如果尚未登录，服务器会打开专用浏览器窗口；
手工登录后让 Agent 再次调用 `status` 或 `list_messages`。

主要工具包括 `list_folders`、`list_messages`、`search_messages`、`get_message`、
`get_conversation`、`download_attachments`、`sync_obsidian`、`create_message`、
`reply_message`、`forward_message`、草稿操作，以及移动、归档和删除。
发送类工具默认创建草稿；自动发送需要 `draft=false` 和唯一 `requestId`。移动、归档、删除和放弃草稿
必须同时提供 `confirmed=true` 与唯一 `requestId`。MCP 进程的 stdout 只用于协议消息，错误写到 stderr。

可以直接对 Agent 这样说：

- “列出今天 Inbox 的邮件。”——Agent 调用 `list_messages`。
- “列出收件箱/投后目录 2026-08-20 的邮件。”——传入 `date` 和 `directory`。
- “读取这封邮件和附件信息。”——Agent 使用列表返回的 `stableId` 调用 `get_message`。
- “回复这封邮件，内容是‘收到，谢谢’，先保存草稿。”——调用 `reply_message`，保持 `draft=true`。
- “把今天投后目录的邮件同步到这个 Obsidian 目录。”——调用 `sync_obsidian`。

对于自动发送、移动、归档、删除和放弃草稿，Agent 应先向用户确认。MCP 工具不会因为自然语言里出现
“确认”就跳过参数检查：调用时仍必须提供相应的 `requestId`，破坏性操作还必须提供
`confirmed=true`。

请先在当前浏览器后端中打开一封无敏感内容的测试邮件，再生成详情勘察文件：

```bash
node dist/cli.js inspect-message --json > outlook-message-inspect.json
```

## 开发

要求 Node.js 24+，并安装 Edge、Chrome 或 Chromium：

```bash
npm install
npm run check
npm test
npm run build
node dist/cli.js --help
```

生成可发布的 npm 包、macOS ZIP、Windows ZIP 和 SHA-256 校验文件：

```bash
npm run package:release
ls release/
```

GitHub Actions 的 `Build release packages` 可手工运行并下载构建产物；推送 `v*` 标签时，工作流还会创建
对应 GitHub Release 并上传所有安装包。

确认已经在 Playwright 专用浏览器 Profile 中登录 Outlook 后：

```bash
node dist/cli.js status --json
node dist/cli.js inspect --json > outlook-inbox-inspect.json
```

`outlook-inbox-inspect.json` 可能包含邮件主题、发件人和预览文本，仅应保存在本机，
并已通过 `.gitignore` 排除。

JSON 命令成功时返回：

```json
{"ok":true,"data":{}}
```

失败时返回非零退出码和：

```json
{"ok":false,"error":{"code":"AUTH_REQUIRED","message":"..."}}
```
