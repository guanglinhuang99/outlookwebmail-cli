# webmail-cli

本地单用户 Outlook Web CLI。它通过 Ego Lite 操作已经登录的
`https://partner.outlook.cn/mail/`，不使用 Graph、EWS、IMAP、SMTP，
也不读取 Cookie、Bearer Token 或浏览器 Storage。

当前已提供：

```bash
webmail status
webmail status --json
webmail inspect
webmail inspect --json
webmail inbox --limit 20 --json
webmail inbox --unread --json
webmail inbox --dir "投后" --limit 20 --json
webmail today --json
webmail today --dir "收件箱/投后" --json
webmail search "风险报告" --limit 20 --json
webmail folders --json
webmail inspect-message --json
webmail read 3 --json
webmail attachments 3 --json
webmail download 3 1 --output ./downloads --json
webmail move 3 "投后" --yes --json
webmail delete 3 --yes --json
```

`read`、`download`、`move`、`delete` 使用最近一次 `inbox/search/today` 产生的 Session 短 ID；
再次运行列表命令会重新生成这些 ID。移动或删除成功后会清空旧 Session，必须重新运行列表命令。

- `folders`：递归展开并列出 Inbox 下的全部子目录，返回可供 `--dir` 使用的 `name` 和 `path`。
- `inbox`：不传 `--dir` 时列 Inbox；传入时列指定 Inbox 子目录。完整 `path` 优先于名称。
- `today`：不传 `--dir` 时列 Inbox 当日邮件；传入时列指定子目录的当日邮件。
- `read`：返回邮件头、纯文本正文和附件元数据。
- `attachments`：只返回附件名称和大小。
- `download`：将指定附件下载到本地目录并返回绝对路径和字节数。
- `move`：只接受移动菜单中完全匹配的目录名，并要求 `--yes`。
- `delete`：将邮件移入“已删除邮件”，并要求 `--yes`。

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
