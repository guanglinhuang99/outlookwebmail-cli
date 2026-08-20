# webmail-cli

本地单用户 Outlook Web CLI。它通过 Ego Lite 操作已经登录的
`https://partner.outlook.cn/mail/`，不使用 Graph、EWS、IMAP、SMTP，
也不读取 Cookie、Bearer Token 或浏览器 Storage。

当前已提供：

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
webmail move 3 "投后" --yes --request-id move-20260820-001 --json
webmail delete 3 --yes --request-id delete-20260820-001 --json
```

列表中的每封邮件同时返回数字短 `id` 和形如 `m_...` 的不透明 `stableId`。数字 ID 只代表当前页；
`stableId` 会在列表刷新和分页后保留，推荐 Agent 保存并使用它。所有读取和单封邮件操作均兼容两种 ID。
发送、放弃草稿、移动、归档或删除成功后会清空当前 Session，必须重新运行相应列表命令。

首次执行任意需要 Outlook 的命令时，如果 Ego Lite 中尚未打开 Webmail，CLI 会自动创建或复用
`webmail-cli` 任务空间并打开 `https://partner.outlook.cn/mail/`。如果尚未登录，CLI 会将页面控制权交给用户；
请在 Ego Lite 中手工完成登录（CLI 不读取用户名、密码或验证码），然后将控制权交还给 Agent，再重新执行原命令。

- `folders`：递归展开并列出 Inbox 下的全部子目录，返回可供 `--dir` 使用的 `name` 和 `path`。
- `doctor`：只读检查 Node.js、Ego Lite 连接、Outlook 登录状态和关键 DOM；任一检查失败时进程返回非零退出码。
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
- `move`：只接受移动菜单中完全匹配的目录名，并要求 `--yes` 与唯一 `--request-id`。
- `delete`：将邮件移入“已删除邮件”，并要求 `--yes` 与唯一 `--request-id`。

自动新建/转发/回复、发送或放弃草稿、移动、归档和删除使用 request ID 防止 Agent 重试造成重复动作。相同 request ID 与相同参数再次调用时不会再次操作
Outlook；request ID 不允许复用于不同参数。状态保存在 `~/.webmail-cli/mutations.json`，脱敏审计写入
`~/.webmail-cli/audit.jsonl`；审计不记录邮件主题、地址、正文或回复内容。

请先在 Ego Lite 中打开一封无敏感内容的测试邮件，再生成详情勘察文件：

```bash
node dist/cli.js inspect-message --json > outlook-message-inspect.json
```

## 开发

要求 Node.js 24+ 和已完成初始化的 Ego Lite：

```bash
npm install
npm run check
npm test
npm run build
node dist/cli.js --help
```

确认已经在 Ego Lite 中登录 Outlook 后：

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
